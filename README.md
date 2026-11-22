# PDF Bill Extractor

Pipeline that extracts invoice data from PDF bills (OCR for scanned pages),
classifies and de-duplicates them, re-files the PDFs by year/site, and
generates Excel + HTML reports.

## Prerequisites

- **Python 3.10+**
- Two system binaries (the pip packages are only wrappers around them):
  - **Tesseract** — OCR engine
  - **Poppler** (`pdftoppm`) — renders PDF pages to images for OCR

Both are auto-detected on Windows, Linux and macOS (see `ocr_deps.py`).
If they are not found, the scripts fail fast with exact install steps.

## Setup

### 1. Clone and install Python packages (all platforms)

```bash
git clone <repo-url> pdf-bill-extractor
cd pdf-bill-extractor
python -m venv .venv
```

Activate the virtual environment:

- **Windows (PowerShell):** `.venv\Scripts\Activate.ps1`
- **Windows (cmd):** `.venv\Scripts\activate.bat`
- **Linux/macOS:** `source .venv/bin/activate`

Then:

```bash
pip install -r requirements.txt
```

### 2. Install the system binaries

#### Windows

**Tesseract**

1. Download the installer from <https://github.com/UB-Mannheim/tesseract/wiki>
2. Run it with the default options (installs to `C:\Program Files\Tesseract-OCR`)
3. Restart your terminal — it is auto-detected, no configuration needed

**Poppler**

1. Download the latest `Release-xxxx-64bit.zip` from
   <https://github.com/oschwartz10612/poppler-windows/releases>
2. Extract it to a permanent location, e.g. `C:\poppler`
3. Point the pipeline at it with an environment variable:

   ```bat
   setx POPPLER_PATH "C:\poppler\bin"
   ```

4. Restart your terminal (or add `C:\poppler\bin` to your PATH instead)

#### Linux (Debian/Ubuntu)

```bash
sudo apt-get update
sudo apt-get install -y tesseract-ocr poppler-utils
```

#### macOS

```bash
brew install tesseract poppler
```

### 3. Verify

```bash
python setup_check.py
```

It checks the Python version, all pip packages, and both binaries, and prints
the exact fix for anything missing. Exit code `0` means you're ready.

> Overrides: set `TESSERACT_CMD` to a full path of the tesseract executable,
> and/or `POPPLER_PATH` to the directory containing `pdftoppm`.

## Usage

```bash
python extractpdfs.py     # main pipeline: extract, classify, re-file, reports
python classifypdfs.py    # layout/document-type analysis (writes CSVs)
python merge_reports.py   # merge old + fresh extraction reports
python check_merge.py     # diagnose rows lost/changed between report versions
```

Inputs come from `Input/` (or the year folders for `classifypdfs.py`); outputs
go to `Output/` — Excel reports, an HTML dashboard, exceptions report, and
re-filed PDFs.

Note: the HTML dashboard loads jQuery/DataTables from a CDN, so opening it
requires internet access.

## Project layout

| File | Purpose |
|---|---|
| `extractpdfs.py` | Main extraction pipeline (multi-process) |
| `classifypdfs.py` | Invoice layout / document type classifier |
| `merge_reports.py` | Merges old + fresh Excel reports |
| `check_merge.py` | Row-level diff between report versions |
| `excel_report.py` | Shared Excel writer (merges, widths, heights) |
| `html_report.py` | Interactive HTML dashboard generator |
| `exceptions_report.py` | HTML exceptions report (dups, errors, mismatches) |
| `report_common.py` | Shared helpers for grouping / mismatch detection |
| `ocr_deps.py` | Tesseract/Poppler detection + friendly error messages |
| `setup_check.py` | Environment diagnostic (`python setup_check.py`) |
| `scripts/debug/` | One-off per-PDF inspection scripts |

## Troubleshooting

- **`ERROR: missing OCR dependency(ies)`** — follow the printed instructions;
  on Windows make sure you restarted the terminal after installing.
- **OCR slow / no text** — Tesseract is only used for scanned PDFs; digitally
  created PDFs are read directly with pdfplumber.
- **Reports open but no filter boxes** — the page needs internet access for
  the CDN scripts.
