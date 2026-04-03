"""
Core invoice processing logic.
Extracts customer name, date, and duplicate status from Onia PDF invoices.
"""

import csv
import io
import logging
import os
import re
import tempfile
import zipfile
from collections import Counter, defaultdict
from datetime import datetime

import fitz  # PyMuPDF

# OCR support (optional)
try:
    import pytesseract
    from PIL import Image
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# Fallback mapping (used only when PDF text extraction fails)
CUSTOMER_CODE_MAP = {
    "C928": "NORDSTROMCOM DROPSHIP",
    "URBN": "URBN DROPSHIP",
    "POS": "Onia Retail Store - Madison",
    "MADISON": "MADISON STORE",
    "STR001": "STARBOARD HOLDINGS LTD",
    "RE703": "EMINENT INC DBA REVOLVE CLOTHING",
    "RVRS001": "RANCHO VALENCIA RESORT and SPA",
    "YAAMA01": "YAAMAVA RESORT and CASINO AT SAN MIGUEL",
    "VERON01": "VERONICA BEARD",
    "VIC04": "THE VINOY",
    "MAY001": "MAYFLOWER INN AND SPA",
    "POP8": "POP UP WESTCHESTER",
    "MRTIQ": "MRTIQUE",
    "SAKS0001": "SAKS FIFTH AVENUE",
}


def clean_customer_name(raw_name: str) -> str:
    """Clean a raw BILL TO name into a safe folder name."""
    name = raw_name.strip()
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


def _ocr_page(pdf_path: str) -> str:
    """OCR fallback: render page 1 as image and run Tesseract."""
    if not OCR_AVAILABLE:
        return ""
    try:
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        doc.close()
        return pytesseract.image_to_string(img)
    except Exception:
        return ""


def _extract_from_pdf(pdf_path: str) -> dict | None:
    """Extract customer name, date, and duplicate flag from page 1."""
    try:
        doc = fitz.open(pdf_path)
        text = doc.load_page(0).get_text()
        doc.close()
    except Exception:
        return None

    if len(text.strip()) < 50:
        text = _ocr_page(pdf_path)
        if len(text.strip()) < 50:
            return None

    # Date
    date_match = re.search(r'(\d{2}/\d{2}/\d{2})', text[:200])
    if date_match:
        try:
            dt = datetime.strptime(date_match.group(1), "%m/%d/%y")
            year, month = dt.year, dt.month
        except ValueError:
            year, month = None, None
    else:
        year, month = None, None

    # Customer name
    ship_label_pos = text.find("SHIP TO\n")
    customer_name = None

    if ship_label_pos != -1:
        pre = text[:ship_label_pos]
        lines = pre.strip().split('\n')
        stripped = [l.strip() for l in lines]
        counts = Counter(stripped)

        for i, s in enumerate(stripped):
            if i > 15 and counts[s] >= 2 and len(s) > 3:
                if re.search(r'^\d+\s+.*\b(ST|AVE|DR|RD|STREET|DRIVE|BLVD|WAY|LANE|LN|HWY)\b', s, re.I):
                    continue
                if re.search(r',\s+\w+\s+\d+.*\s+[A-Z]{2}$', s):
                    continue
                customer_name = s
                break

    if not customer_name:
        return None

    return {
        "customer_raw": customer_name,
        "customer_clean": clean_customer_name(customer_name),
        "year": year,
        "month": month,
        "is_duplicate": "*** DUPLICATE ***" in text,
        "source": "pdf",
    }


def _extract_from_filename(filename: str) -> dict | None:
    """Fallback: extract customer from the filename code."""
    parts = filename.replace('.PDF', '').replace('.pdf', '').split('_')
    if len(parts) >= 2:
        code = parts[1]
        return {
            "customer_raw": code,
            "customer_clean": CUSTOMER_CODE_MAP.get(code, code),
            "year": None,
            "month": None,
            "is_duplicate": False,
            "source": "filename" if code in CUSTOMER_CODE_MAP else "filename_unknown",
        }
    return None


def process_single_file(pdf_bytes: bytes, filename: str, skip_duplicates: bool = True) -> dict:
    """Process a single PDF from in-memory bytes."""
    result = {
        "filename": filename,
        "customer_clean": None,
        "year": None,
        "month": None,
        "is_duplicate": False,
        "status": "error",
        "method": None,
        "error_message": None,
    }

    # Write to temp file for fitz
    try:
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            tmp.write(pdf_bytes)
            tmp_path = tmp.name

        data = _extract_from_pdf(tmp_path)
        if data is None:
            data = _extract_from_filename(filename)
            if data is None:
                result["error_message"] = "Could not extract data from PDF or filename"
                return result
    except Exception as e:
        result["error_message"] = str(e)
        return result
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass

    if skip_duplicates and data.get("is_duplicate", False):
        result["status"] = "skipped_duplicate"
        result["method"] = data["source"]
        result["customer_clean"] = data["customer_clean"]
        return result

    result["customer_clean"] = data["customer_clean"]
    result["year"] = data["year"]
    result["month"] = data["month"]
    result["method"] = data["source"]
    result["status"] = "ok"
    return result


def compute_summary(results: list[dict]) -> dict:
    """Compute summary statistics from processing results."""
    ok = [r for r in results if r["status"] == "ok"]
    duplicates = [r for r in results if r["status"] == "skipped_duplicate"]
    errors = [r for r in results if r["status"] == "error"]

    customer_counts = defaultdict(int)
    year_month_counts = defaultdict(int)

    for r in ok:
        customer_counts[r["customer_clean"]] += 1
        if r["year"] and r["month"]:
            ym = f"{r['year']}/{r['month']:02d}"
            year_month_counts[ym] += 1

    return {
        "total": len(results),
        "ok_count": len(ok),
        "duplicate_count": len(duplicates),
        "error_count": len(errors),
        "customer_counts": dict(sorted(customer_counts.items(), key=lambda x: -x[1])),
        "year_month_counts": dict(sorted(year_month_counts.items())),
        "errors": [{"File": r["filename"], "Error": r.get("error_message", "Unknown")} for r in errors],
    }


def build_zip(ok_results: list[dict], uploaded_files: dict[str, bytes]) -> bytes:
    """Build a ZIP with CustomerName/Year/Month/ folder structure."""
    buf = io.BytesIO()
    customer_year_month = defaultdict(int)

    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for r in ok_results:
            filename = r["filename"]
            customer = r["customer_clean"]

            if r["year"] and r["month"]:
                path = f"{customer}/{r['year']}/{r['month']:02d}/{filename}"
                ym = f"{r['year']}/{r['month']:02d}"
                customer_year_month[(customer, ym)] += 1
            else:
                path = f"{customer}/_UNKNOWN_DATE/{filename}"

            if filename in uploaded_files:
                zf.writestr(path, uploaded_files[filename])

        # Add report.csv
        csv_buf = io.StringIO()
        writer = csv.writer(csv_buf)
        writer.writerow(["customer", "year_month", "count"])
        for (cust, ym) in sorted(customer_year_month):
            writer.writerow([cust, ym, customer_year_month[(cust, ym)]])
        zf.writestr("report.csv", csv_buf.getvalue())

    buf.seek(0)
    return buf.getvalue()
