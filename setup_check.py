#!/usr/bin/env python3
"""Clone-and-run diagnostic for this repository.

Checks, in order:
  1. Python version
  2. Required Python packages (pip)
  3. Required system binaries (Tesseract + Poppler)

Usage:
    python setup_check.py

Exit code 0 means everything is installed and the pipeline is ready to run.
"""

import importlib
import sys

import ocr_deps

MIN_PYTHON = (3, 10)

REQUIRED_PACKAGES = [
    "pandas",
    "numpy",
    "pdfplumber",
    "pytesseract",
    "pdf2image",
    "openpyxl",
]


def main() -> int:
    ok = True
    print("=" * 62)
    print("pdf-bill-extractor — environment check")
    print("=" * 62)

    # 1. Python version
    ver = sys.version_info
    if (ver.major, ver.minor) >= MIN_PYTHON:
        print(f"[ok]   Python {ver.major}.{ver.minor}.{ver.micro}")
    else:
        ok = False
        print(f"[FAIL] Python {ver.major}.{ver.minor} — need {MIN_PYTHON[0]}.{MIN_PYTHON[1]} or newer")
        print("       https://www.python.org/downloads/")

    # 2. Python packages
    missing_pkgs = []
    for pkg in REQUIRED_PACKAGES:
        try:
            mod = importlib.import_module(pkg)
            name = getattr(mod, "__name__", pkg)
            version = getattr(mod, "__version__", "")
            print(f"[ok]   {pkg}{f' {version}' if version else ''}")
        except ImportError:
            missing_pkgs.append(pkg)
            ok = False
            print(f"[FAIL] {pkg} — not installed")
    if missing_pkgs:
        print(f"       Fix with:  pip install {' '.join(missing_pkgs)}")
        print("       (or: pip install -r requirements.txt)")

    # 3. System binaries
    report = ocr_deps.apply_ocr_configuration()
    for name, path in report.items():
        if path:
            print(f"[ok]   {name:<10} {path}")
            print(f"       ({ocr_deps._version_line(path)})")
        else:
            ok = False
            print(f"[FAIL] {name:<10} not found")

    print("-" * 62)
    if ok:
        print("All checks passed — the pipeline is ready to run.")
        print("Next:  python extractpdfs.py")
    else:
        print("Some checks failed. Install instructions:")
        print()
        missing = [n for n, p in report.items() if not p]
        if missing_pkgs:
            print("  * Python packages:  pip install -r requirements.txt")
        if missing:
            print(ocr_deps._install_instructions(missing))
    print("=" * 62)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
