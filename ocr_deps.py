#!/usr/bin/env python3
"""Cross-platform detection of the external binaries the OCR pipeline needs.

pdf2image and pytesseract are thin wrappers around two system tools:

  * Tesseract  (tesseract / tesseract.exe)  -- the OCR engine
  * Poppler    (pdftoppm  / pdftoppm.exe)   -- renders PDF pages to images

This module locates both binaries on Windows, Linux and macOS, points the
wrapper libraries at them, and can fail fast with friendly install
instructions when something is missing.

Environment variable overrides (checked first):
  TESSERACT_CMD   full path to the tesseract executable
  POPPLER_PATH    directory containing pdftoppm
"""

import os
import shutil
import subprocess
from pathlib import Path

IS_WINDOWS = os.name == "nt"

# Common install locations, checked when the binary is not found on PATH.
TESSERACT_CANDIDATES = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]
POPPLER_CANDIDATES = [
    r"C:\poppler\bin\pdftoppm.exe",
    r"C:\Program Files\poppler\bin\pdftoppm.exe",
]


def _find_tesseract():
    env = os.environ.get("TESSERACT_CMD")
    if env and Path(env).exists():
        return env
    found = shutil.which("tesseract")
    if found:
        return found
    for cand in TESSERACT_CANDIDATES:
        if Path(cand).exists():
            return cand
    return None


def _find_pdftoppm():
    env_dir = os.environ.get("POPPLER_PATH")
    if env_dir:
        exe = Path(env_dir) / ("pdftoppm.exe" if IS_WINDOWS else "pdftoppm")
        if exe.exists():
            return str(exe)
    found = shutil.which("pdftoppm")
    if found:
        return found
    for cand in POPPLER_CANDIDATES:
        if Path(cand).exists():
            return cand
    return None


def _version_line(exe) -> str:
    """Best-effort first line of `--version` output (tools print to stderr)."""
    try:
        out = subprocess.run(
            [str(exe), "--version"],
            capture_output=True, text=True, timeout=15,
        )
        first = (out.stderr or out.stdout or "").strip().splitlines()
        return first[0] if first else "version unknown"
    except Exception as e:
        return f"found, but not runnable ({e})"


def apply_ocr_configuration():
    """Locate the binaries and point the wrapper libraries at them.

    Safe to call repeatedly (it runs at import time so that every worker
    process started by ProcessPoolExecutor picks up the configuration too).
    Wrapper packages are imported lazily so this module stays importable on a
    fresh clone before `pip install -r requirements.txt` has been run.
    Returns {"tesseract": path|None, "pdftoppm": path|None}.
    """
    report = {"tesseract": _find_tesseract(), "pdftoppm": _find_pdftoppm()}
    if report["tesseract"]:
        try:
            import pytesseract
            pytesseract.pytesseract.tesseract_cmd = report["tesseract"]
        except ImportError:
            pass
    if report["pdftoppm"]:
        try:
            from pdf2image import configuration as pdf2image_config
            pdf2image_config.poppler_path = str(Path(report["pdftoppm"]).parent)
        except ImportError:
            pass
    return report


def _install_instructions(missing: list) -> str:
    lines = []
    if IS_WINDOWS:
        lines.append("Windows setup — install the missing tools, then re-run:")
        if "tesseract" in missing:
            lines += [
                "",
                "  * Tesseract OCR:",
                "      1. Download the Windows installer from",
                "         https://github.com/UB-Mannheim/tesseract/wiki",
                "      2. Run it with the default options",
                "         (it installs to C:\\Program Files\\Tesseract-OCR and is auto-detected),",
                "      3. Restart your terminal.",
                "      Alternative: set TESSERACT_CMD to the full path of tesseract.exe",
            ]
        if "pdftoppm" in missing:
            lines += [
                "",
                "  * Poppler (pdftoppm):",
                "      1. Download the latest 'Release-xxxx-64bit.zip' from",
                "         https://github.com/oschwartz10612/poppler-windows/releases",
                "      2. Extract it somewhere permanent, e.g. C:\\poppler",
                "      3. Set the environment variable POPPLER_PATH to the bin folder,",
                '         e.g.  setx POPPLER_PATH "C:\\poppler\\bin"   (then restart your terminal),',
                "      4. Restart your terminal.",
            ]
    elif os.uname().sysname == "Darwin":
        lines += [
            "macOS setup:",
            "    brew install tesseract poppler",
        ]
    else:
        lines += [
            "Linux setup:",
            "    sudo apt-get update && sudo apt-get install -y tesseract-ocr poppler-utils",
            "    (Debian/Ubuntu — for other distros use the equivalent packages)",
        ]
    return "\n".join(lines)


def ensure_ocr_ready() -> None:
    """Print the detected OCR tools, or fail fast with install instructions."""
    report = apply_ocr_configuration()
    missing = [name for name, path in report.items() if not path]
    if missing:
        print("ERROR: missing OCR dependency(ies): " + ", ".join(missing))
        print()
        print(_install_instructions(missing))
        raise SystemExit(1)
    print(f"[ocr_deps] tesseract : {report['tesseract']}  ({_version_line(report['tesseract'])})")
    print(f"[ocr_deps] pdftoppm  : {report['pdftoppm']}  ({_version_line(report['pdftoppm'])})")


# Apply configuration at import time so worker processes are covered too.
apply_ocr_configuration()
