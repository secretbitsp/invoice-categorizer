# Invoice Categorization Tool

Automatically sorts Onia PDF invoices into folders by **Customer Name / Year / Month**.

## Before You Start

Make sure Python 3 is installed, then install the required library:

```
pip install PyMuPDF
```

Optional (only needed if some PDFs are scanned images instead of digital):
```
pip install pytesseract Pillow
```

## How to Use

### Step 1 - Place your invoices

Put all your PDF invoices inside a folder (e.g., `invoices/`).
They can be in subfolders - the script finds them all automatically.

### Step 2 - Preview first (dry run)

```
python3 categorize_invoices.py --input-dir ./invoices
```

This shows you what the script **would do** without moving anything.
Review the customer names and counts to make sure everything looks right.

### Step 3 - Run for real

```
python3 categorize_invoices.py --input-dir ./invoices --run
```

This copies all invoices into the `output/` folder, organized like this:

```
output/
  NORDSTROMCOM DROPSHIP/
    2025/
      11/
        100000026_URBN_20260318.PDF
      12/
        ...
    2026/
      01/
        ...
  STARBOARD HOLDINGS LTD/
    2026/
      02/
        ...
  ...
```

### Step 4 - Check the results

- Open the `output/` folder and browse the customer folders
- Check `output/report.csv` for a full breakdown (customer, year/month, count)
- If any invoices failed, check `output/errors.log`

## Options

| Option | Description |
|--------|-------------|
| `--input-dir PATH` | Where your PDFs are (default: `./invoices`) |
| `--output-dir PATH` | Where to save sorted invoices (default: `./output`) |
| `--run` | Actually copy files (without this, it only previews) |
| `--workers N` | Number of parallel workers for speed (default: auto) |

## Examples

Preview 18,000 invoices:
```
python3 categorize_invoices.py --input-dir /path/to/all-invoices
```

Sort them for real into a custom folder:
```
python3 categorize_invoices.py --input-dir /path/to/all-invoices --output-dir /path/to/sorted --run
```

## What the Script Does

1. Opens each PDF and reads page 1 only
2. Extracts the **customer name** from the BILL TO section
3. Extracts the **invoice date**
4. Copies the file into `CustomerName/Year/Month/` folder
5. If a PDF can't be read, it tries OCR (if installed) or uses the filename as fallback
6. Generates a summary report when done

## Notes

- Original files are **copied, not moved** - your originals stay untouched
- The script handles multi-page invoices (only reads page 1)
- You can safely re-run it - existing files in output will be overwritten
- Works with both `.PDF` and `.pdf` file extensions
