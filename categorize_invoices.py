#!/usr/bin/env python3
"""
Categorize Onia PDF invoices into CustomerName/Year/Month folders.

Usage:
    python categorize_invoices.py                          # dry run on ./invoices -> ./output
    python categorize_invoices.py --run                    # actually copy files
    python categorize_invoices.py --input-dir /path/to/pdfs --output-dir /path/to/output --run
"""

import argparse
import csv
import logging
import os
import re
import shutil
import sys
from collections import Counter, defaultdict
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF

from invoice_processor import apply_invoice_number_dedupe, extract_invoice_number

# OCR support (optional - used only if PDF has no extractable text)
try:
    import pytesseract
    from PIL import Image
    import io
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# ---------------------------------------------------------------------------
# Fallback mapping: filename customer code -> cleaned customer name
# Used only when PDF text extraction fails
# ---------------------------------------------------------------------------
CUSTOMER_CODE_MAP = {
    "C928":    "NORDSTROMCOM DROPSHIP",
    "URBN":    "URBN DROPSHIP",
    "POS":     "Onia Retail Store - Madison",
    "MADISON": "MADISON STORE",
    "STR001":  "STARBOARD HOLDINGS LTD",
    "RE703":   "EMINENT INC DBA REVOLVE CLOTHING",
    "RVRS001": "RANCHO VALENCIA RESORT and SPA",
    "YAAMA01": "YAAMAVA RESORT and CASINO AT SAN MIGUEL",
    "VERON01": "VERONICA BEARD",
    "VIC04":   "THE VINOY",
    "MAY001":  "MAYFLOWER INN AND SPA",
    "POP8":    "POP UP WESTCHESTER",
    "MRTIQ":   "MRTIQUE",
    "SAKS0001": "SAKS FIFTH AVENUE",
}


def clean_customer_name(raw_name: str) -> str:
    """Clean a raw BILL TO name into a safe folder name."""
    name = raw_name.strip()
    # Remove parenthetical account codes: "(SAKS0001)"
    name = re.sub(r'\s*\([^)]*\)\s*', '', name)
    name = name.rstrip('.')
    name = name.replace('&', 'and')
    name = name.replace("'", '')
    name = name.replace(',', '')
    name = name.replace('.', '')
    name = name.replace(':', '')
    name = name.replace('/', '-')
    name = re.sub(r'\s+', ' ', name)
    return name.strip()


def ocr_page(pdf_path: str) -> str:
    """OCR fallback: render page 1 as image and run Tesseract."""
    if not OCR_AVAILABLE:
        return ""
    try:
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        doc.close()
        text = pytesseract.image_to_string(img)
        return text
    except Exception as e:
        logging.warning(f"OCR failed for {pdf_path}: {e}")
        return ""


def extract_invoice_data(pdf_path: str) -> dict | None:
    """Extract customer name and date from page 1 of an invoice PDF."""
    try:
        doc = fitz.open(pdf_path)
        text = doc.load_page(0).get_text()
        doc.close()
    except Exception as e:
        logging.warning(f"Cannot read PDF {pdf_path}: {e}")
        return None

    # If no text extracted (scanned image PDF), try OCR
    if len(text.strip()) < 50:
        logging.info(f"Low text content, trying OCR: {pdf_path}")
        text = ocr_page(pdf_path)
        if len(text.strip()) < 50:
            logging.warning(f"OCR also returned no text: {pdf_path}")
            return None

    # --- Extract date (first MM/DD/YY in the first 200 chars) ---
    date_match = re.search(r'(\d{2}/\d{2}/\d{2})', text[:200])
    if date_match:
        try:
            dt = datetime.strptime(date_match.group(1), "%m/%d/%y")
            year, month = dt.year, dt.month
        except ValueError:
            year, month = None, None
    else:
        year, month = None, None

    # --- Extract customer name ---
    # In PyMuPDF text output, the SHIP TO and BILL TO address blocks appear
    # BEFORE the "SHIP TO\nBILL TO" labels. The customer name line always
    # appears exactly twice (once in each block). We find it by looking for
    # the first line (after header fields) that appears exactly 2 times.
    ship_label_pos = text.find("SHIP TO\n")
    customer_name = None

    if ship_label_pos != -1:
        pre = text[:ship_label_pos]
        lines = pre.strip().split('\n')
        stripped = [l.strip() for l in lines]
        counts = Counter(stripped)

        for i, s in enumerate(stripped):
            if i > 15 and counts[s] >= 2 and len(s) > 3:
                # Skip lines that look like street addresses or city/zip lines
                if re.search(r'^\d+\s+.*\b(ST|AVE|DR|RD|STREET|DRIVE|BLVD|WAY|LANE|LN|HWY)\b', s, re.I):
                    continue
                if re.search(r',\s+\w+\s+\d+.*\s+[A-Z]{2}$', s):
                    continue
                customer_name = s
                break

    if not customer_name:
        logging.warning(f"Could not extract customer from {pdf_path}")
        return None

    # --- Check for duplicate marker ---
    is_duplicate = "*** DUPLICATE ***" in text
    basename = os.path.basename(pdf_path)
    inv = extract_invoice_number(text, basename)

    return {
        "customer_raw": customer_name,
        "customer_clean": clean_customer_name(customer_name),
        "year": year,
        "month": month,
        "is_duplicate": is_duplicate,
        "invoice_number": inv,
        "source": "pdf",
    }


def extract_from_filename(pdf_path: str) -> dict | None:
    """Fallback: extract customer from the filename code."""
    basename = os.path.basename(pdf_path)
    parts = basename.replace('.PDF', '').replace('.pdf', '').split('_')
    if len(parts) >= 2:
        code = parts[1]
        inv = extract_invoice_number("", basename)
        if code in CUSTOMER_CODE_MAP:
            return {
                "customer_raw": code,
                "customer_clean": CUSTOMER_CODE_MAP[code],
                "year": None,
                "month": None,
                "invoice_number": inv,
                "source": "filename",
            }
        else:
            return {
                "customer_raw": code,
                "customer_clean": code,
                "year": None,
                "month": None,
                "invoice_number": inv,
                "source": "filename_unknown",
            }
    return None


def process_single_invoice(args: tuple) -> dict:
    """Process one PDF: extract metadata and destination path (copy happens in main)."""
    pdf_path, output_dir = args
    basename = os.path.basename(pdf_path)
    result = {
        "source_path": pdf_path,
        "dest_path": None,
        "customer": None,
        "year": None,
        "month": None,
        "invoice_number": None,
        "status": "error",
        "method": None,
    }

    # Try PDF extraction first, then filename fallback
    data = extract_invoice_data(pdf_path)
    if data is None:
        data = extract_from_filename(pdf_path)
        if data is None:
            logging.error(f"FAILED completely: {pdf_path}")
            return result

    result["invoice_number"] = data.get("invoice_number")

    # Skip PDFs marked *** DUPLICATE *** (Onia print copy)
    if data.get("is_duplicate", False):
        result["status"] = "skipped_duplicate"
        result["method"] = data["source"]
        return result

    customer_folder = data["customer_clean"]
    year = data["year"]
    month = data["month"]

    if year and month:
        dest_dir = os.path.join(output_dir, customer_folder, str(year), f"{month:02d}")
    else:
        dest_dir = os.path.join(output_dir, customer_folder, "_UNKNOWN_DATE")

    dest_path = os.path.join(dest_dir, basename)

    result["dest_path"] = dest_path
    result["customer"] = customer_folder
    result["year"] = year
    result["month"] = month
    result["method"] = data["source"]

    result["status"] = "ok"
    return result


def discover_pdfs(input_dir: str) -> list[str]:
    """Find all PDF files recursively."""
    pdfs = []
    for ext in ("*.PDF", "*.pdf"):
        pdfs.extend(str(p) for p in Path(input_dir).rglob(ext))
    return sorted(set(pdfs))


def generate_report(results: list[dict], output_dir: str, do_copy: bool):
    """Print summary and write report.csv + errors.log."""
    # Counts
    ok = [r for r in results if r["status"] == "ok"]
    errors = [r for r in results if r["status"] == "error"]
    duplicates = [r for r in results if r["status"] == "skipped_duplicate"]
    dup_inv = [r for r in results if r["status"] == "skipped_duplicate_invoice"]
    fallbacks = [r for r in ok if r["method"] and r["method"].startswith("filename")]

    # Customer summary
    customer_counts = defaultdict(int)
    year_month_counts = defaultdict(int)
    customer_year_month = defaultdict(int)

    for r in ok:
        customer_counts[r["customer"]] += 1
        if r["year"] and r["month"]:
            ym = f"{r['year']}/{r['month']:02d}"
            year_month_counts[ym] += 1
            customer_year_month[(r["customer"], ym)] += 1

    # Console output
    print(f"\n{'='*60}")
    print(f"  Invoice Categorization {'REPORT' if not do_copy else 'COMPLETE'}")
    print(f"{'='*60}")
    print(f"  Total processed:  {len(results)}")
    print(f"  Originals copied: {len(ok)}")
    print(f"  Duplicates skip:  {len(duplicates)} (PDF marked *** DUPLICATE ***)")
    print(f"  Same invoice #: {len(dup_inv)} (extra files with same invoice number)")
    print(f"  Failed:           {len(errors)}")
    print(f"  Filename fallback:{len(fallbacks)}")
    print()

    print(f"  {'Customer':<45s} {'Count':>6s}")
    print(f"  {'-'*45} {'-'*6}")
    for cust in sorted(customer_counts, key=lambda c: -customer_counts[c]):
        print(f"  {cust:<45s} {customer_counts[cust]:>6d}")
    print()

    print(f"  {'Year/Month':<15s} {'Count':>6s}")
    print(f"  {'-'*15} {'-'*6}")
    for ym in sorted(year_month_counts):
        print(f"  {ym:<15s} {year_month_counts[ym]:>6d}")
    print()

    if not do_copy:
        print("  [DRY RUN] No files were copied. Use --run to copy.\n")

    # Write report.csv
    if do_copy:
        report_path = os.path.join(output_dir, "report.csv")
        with open(report_path, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["customer", "year_month", "count"])
            for (cust, ym) in sorted(customer_year_month):
                writer.writerow([cust, ym, customer_year_month[(cust, ym)]])
        print(f"  Report saved: {report_path}")

    # Write errors.log
    if errors:
        errors_path = os.path.join(output_dir if do_copy else ".", "errors.log")
        if do_copy:
            os.makedirs(output_dir, exist_ok=True)
        with open(errors_path, "w") as f:
            for r in errors:
                f.write(f"FAILED: {r['source_path']}\n")
        print(f"  Errors logged: {errors_path}")

    # Preview first 15 moves in dry-run
    if not do_copy and ok:
        print(f"\n  Preview (first 15 of {len(ok)}):")
        for r in ok[:15]:
            src = os.path.basename(r["source_path"])
            dst = r["dest_path"].replace(output_dir + "/", "")
            print(f"    {src}  ->  {dst}")
        if len(ok) > 15:
            print(f"    ... and {len(ok) - 15} more")
        print()


def main():
    parser = argparse.ArgumentParser(
        description="Categorize Onia invoices into CustomerName/Year/Month folders"
    )
    parser.add_argument("--input-dir", default="./invoices",
                        help="Directory containing invoice PDFs (default: ./invoices)")
    parser.add_argument("--output-dir", default="./output",
                        help="Output directory (default: ./output)")
    parser.add_argument("--run", action="store_true",
                        help="Actually copy files (default is dry run)")
    parser.add_argument("--workers", type=int, default=os.cpu_count(),
                        help="Parallel workers (default: CPU count)")
    parser.add_argument("--no-invoice-dedupe", action="store_true",
                        help="Keep all files even when invoice number repeats (default: dedupe)")
    args = parser.parse_args()

    # Discover PDFs
    print(f"Scanning {args.input_dir} for PDFs...")
    pdf_files = discover_pdfs(args.input_dir)
    print(f"Found {len(pdf_files)} PDF files")

    if not pdf_files:
        print("No PDF files found. Check --input-dir path.")
        sys.exit(1)

    do_copy = args.run
    tasks = [(pdf, args.output_dir) for pdf in pdf_files]

    results = []
    with ProcessPoolExecutor(max_workers=args.workers) as executor:
        futures = {executor.submit(process_single_invoice, t): t for t in tasks}
        for future in as_completed(futures):
            try:
                results.append(future.result())
            except Exception as e:
                pdf = futures[future][0]
                logging.error(f"Worker crashed on {pdf}: {e}")
                results.append({
                    "source_path": pdf, "dest_path": None,
                    "customer": None, "year": None, "month": None,
                    "invoice_number": None,
                    "status": "error", "method": None,
                })

    if not args.no_invoice_dedupe:
        apply_invoice_number_dedupe(results)

    if do_copy:
        for r in results:
            if r["status"] == "ok" and r.get("dest_path"):
                os.makedirs(os.path.dirname(r["dest_path"]), exist_ok=True)
                shutil.copy2(r["source_path"], r["dest_path"])

    generate_report(results, args.output_dir, do_copy)


if __name__ == "__main__":
    main()
