#!/usr/bin/env python3


import os
import re
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed

import pandas as pd
import pdfplumber
import pytesseract
from pdf2image import convert_from_path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from html import escape

import shutil
from collections import defaultdict

from excel_report import write_excel_report
from html_report import generate_dashboard_html
from exceptions_report import create_exceptions_report
from report_common import (
    build_total_mismatch_files,
    choose_preferred_invoice_file,
)

# =============================================================================
# 1. CONFIGURATION & CONSTANTS
# =============================================================================

class Config:
    BASE_DIR = Path(__file__).resolve().parent

    INPUT_DIR = BASE_DIR / "Input"
    OUTPUT_DIR = BASE_DIR / "Output"
    DUPLICATES_DIR = OUTPUT_DIR / "Duplicates"
    INDIVIDUAL_FLATS_DIR = OUTPUT_DIR / "Individual Flats"
    NOT_MORELLO_DIR = OUTPUT_DIR / "Not Morello"
    FAILED_DIR = OUTPUT_DIR / "Failed"
    DEBUG_DIR = BASE_DIR / "_debug"

    OUTPUT_XLSX = OUTPUT_DIR / "invoice_extraction.xlsx"
    OUTPUT_HTML = OUTPUT_DIR / "invoice_extraction.html"
    EXCEPTIONS_HTML = OUTPUT_DIR / "extraction_exceptions.html"

    OCR_TEXT_THRESHOLD = 80
    MAX_WORKERS = max(1, min((os.cpu_count() or 4) - 1, 4))
    TESSERACT_CONFIG = r"--oem 3 --psm 6"


class RegexPatterns:
    DATE_VALUES = [
        r"\b\d{4}-\d{2}-\d{2}\b",
        r"\b\d{2}/\d{2}/\d{4}\b",
        r"\b\d{2}/\d{2}/\d{2}\b",
        r"\b\d{2}-\d{2}-\d{4}\b",
        r"\b\d{1,2}(?:st|nd|rd|th)?\.?\s+[A-Za-z]{3,9}\s+\d{4}\b",
    ]
    DATE_STRICT_LABEL = r"\b(credit\s+date|invoice\s+date|document\s+date|tax\s+point|issued|date)\b"
    DATE_FALSE_LABELS = [
        r"date of invoice",
        r"days from date of invoice",
    ]

    DOCNO_VALUE = r"\b(?=[A-Z0-9./-]{2,40}\b)(?=[A-Z0-9./-]*\d)[A-Z0-9][A-Z0-9./-]{1,40}\b"
    INVOICE_STRICT_LABEL = r"(?:invoice\s*(?:number|no\.?|#)|\binv\s*(?:number|no\.?|#)\b)"
    INVOICE_FALLBACKS = [
        r"\binvoice\s*(?:number|no\.?|#)\s*[:.\-]?\s*(" + DOCNO_VALUE + r")\b",
        r"\binv\s*(?:number|no\.?|#)\s*[:.\-]?\s*(" + DOCNO_VALUE + r")\b",
    ]
    CREDIT_STRICT_LABEL = r"(?:credit\s*number|credit\s*no\.?|credit\s*#|credit\s*memo\s*no\.?)"
    CREDITED_INVOICE_PATTERNS = [
        r"(?:credit\s+for\s+invoice|credit\s+against\s+invoice|credited\s+invoice|original\s+invoice)\s*[:.\-]?\s*(" + DOCNO_VALUE + r")",
        r"(?:credit\s+for\s+invoice|credit\s+against\s+invoice|credited\s+invoice|original\s+invoice)\s*[:.\-]?\s*\n+\s*(" + DOCNO_VALUE + r")",
        r"\b(?:credit\s+(?:for\s+)?)(SI\d{4,8}|INV\d{4,8}|[A-Z]{2,4}\d{4,8})\b",
    ]

    DETAIL_HEADERS = [
        r"\bitem\s+detail\s+quantity\s+unit\s+line\s+value\b",
        r"\bdetail\s+quantity\s+unit\s+line\s+value\b",
        r"\bdetail\s+quantity\s+line\s+value\b",
        r"\bno\.\s+description\s+quantity\s+unit\s+price\s+amount\b",
        r"\bno\.\s+description.*",
        r"\bdescription\s+of\s+work\b",
        r"\bproduct\s*/\s*service\s+description\b",
        r"^[\W_]*detail[\W_]*$",
        r"^[\W_]*item\s+detail[\W_]*$",
        r"^[\W_]*description[\W_]*$",
        r"^[\W_]*item\s+description[\W_]*$",
        r"^[\W_]*charges[\W_]*$",
    ]
    DETAIL_STOPS = [
        r"^[\W_]*sub\s*total\b", r"^[\W_]*subtotal\b", r"^[\W_]*total\s+ex\s+vat\b",
        r"^[\W_]*total\s*(?:[£$€])?\s*excl\.?\s*vat\b", r"^[\W_]*total\s*(?:[£$€])?\s*excluding\s*vat\b",
        r"^[\W_]*vat\b", r"^[\W_]*grand\s+total\b", r"^[\W_]*amount\s+due\b",
        r"^[\W_]*balance\s+due\b", r"^[\W_]*job\s+total\b", r"^[\W_]*invoice\s+total\b",
        r"^[\W_]*terms\b", r"^[\W_]*strictly\b", r"^[\W_]*payment\s+terms\b",
        r"^[\W_]*our\s+banking\s+details\b", r"^[\W_]*account\s+name\b", r"^[\W_]*account\s+number\b",
        r"^[\W_]*sort\s+code\b", r"^[\W_]*please\s+quote\s+invoice\b", r"^[\W_]*any\s+problems\s+regarding\s+payment\b",
        r"^[\W_]*invoice\s+date\b", r"^[\W_]*credit\s+date\b", r"^[\W_]*ship[\-\s]*to\b", r"^[\W_]*delivery\s+address\b",
        r"^[\W_]*please\s+remit\s+payments\b", r"^[\W_]*page\s+\d+\s+of\s+\d+\b", r"^[\W_]*for\s+finance\s+queries\b",
        r"^[\W_]*for\s+any\s+urgent\s+repairs\b", r"^[\W_]*customer\s+services\b", r"^[\W_]*fax\s*[:\-]", r"^[\W_]*email\s*[:\-]",
        r"^[\W_]*total\s*[:\-]?\s*(?:[£$€]|\d+)\b", r"^[\W_]*total\s+amount\b"
    ]

    AMOUNT_PRIORITIES = [
        ("total_incl_vat", r"\btotal\s*(?:[£$€])?\s*(?:incl\.?|inc\.?)\s*vat\b"),
        ("total_inc_vat", r"\btotal\s*(?:[£€])?\s*inc\.?\s*vat\b"),
        ("total_including_vat", r"\btotal\s*(?:[£€])?\s*including\s*vat\b"),
        ("grand_total", r"\bgrand\s+total\b"), ("amount_due", r"\bamount\s+due\b"),
        ("balance_due", r"\bbalance\s+due\b"), ("invoice_total", r"\binvoice\s+total\b"), ("total", r"\btotal\b"),
    ]
    AMOUNT_REJECTS = [
        r"\btotal\s+ex\s+vat\b", r"\btotal\s*(?:[£$€])?\s*excl\.?\s*vat\b", r"\btotal\s*(?:[£$€])?\s*excluding\s*vat\b",
        r"^[\W_]*vat\b", r"^[\W_]*sub\s*total\b", r"^[\W_]*subtotal\b", r"^[\W_]*net\b",
        r"^[\W_]*unit\b", r"^[\W_]*rate\b",
    ]

    LINE_ITEM_SKIP = [
        r"^Engineer Comments\b", r"^Attended\b", r"^Job Number\b", r"^Initial Call Type\b",
        r"^Caller Name\b", r"^Time Details\b", r"^Time Charges\b", r"^Travel Between\b",
        r"^Work Between\b", r"^Rounded\b", r"^Rounded upto\b", r"^Minutes\b",
        r"^Split due to charge rate boundary\b", r"^Please remit payments\b", r"^Page\s+\d+",
        r"^For finance queries\b", r"^For any urgent repairs\b", r"^Fax\b", r"^Email\b", r"^Customer Services\b",

        # --- Switch2 Boilerplate & Page Break text ---
        r"^Switch2'?\s+terms\b", r"^Registered\s+Office\b", r"^Sales\s+Invoice\b", r"^INVOICE\s+TO",
        r"^MORELLO\s+QUARTER\b", r"^All\s+correspondence\b", r"^Brandon\s+Schmidt\b",
        r"^C/o\s+Rendall\b", r"^Portsoken\s+House\b", r"^\d+/\d+\s+Minories\b", r"^London\s+EC3N\b",
        r"^Great\s+Britain\b", r"^Switch2\s+Energy\b", r"^The\s+Waterfront\b", r"^Salts\s+Mill\b",
        r"^Shipley\b", r"^West\s+Yorkshire\b", r"^BD17\s+7EZ\b", r"^Bill-to\s+Customer\b",
        r"^Document\s+Date\b", r"^Due\s+Date\b", r"^VAT\s+Registration\b", r"^Bank\s+Barclays\b",
        # --- Short Location/Section Headers ---
        r"^Heat\s*$", r"^Roof\s*\d+\s*$"
    ]
    LINE_ITEM_SUMMARY = [
        r"\bcost of this work\b", r"\btotal cost of this work\b", r"\bthis work will be\b",
        r"\btotal cost\b", r"\btotal\b.*\bexc\b.*\bvat\b", r"\btotal\b.*\bex\b.*\bvat\b",
        r"\bexc\b\.?\s*vat\b", r"\bex\b\.?\s*vat\b",
    ]
    LINE_ITEM_FALLBACK_AMOUNT = r"(?<![\d.])-?\s*[£$€]?\s*[\d,]+\.\d{2}(?=\s|$|[)\],;:])"

    ANCHOR_CHECKS = [
        ("issuer_switch2", r"switch2 energy limited"),
        ("accounts_payable", r"accounts payable"),
        ("invoice_label", INVOICE_STRICT_LABEL),
        ("credit_label", CREDIT_STRICT_LABEL),
        ("line_items", r"(item detail quantity unit line value|detail quantity unit line value|detail quantity line value|no\.\s+description|^[\W_]*detail[\W_]*$|^[\W_]*description[\W_]*$|^[\W_]*charges[\W_]*$)"),
    ]

    CONTRACT_SUMMARY_MARKERS = [
        re.compile(r"^\s*covering\b", re.I),
        re.compile(r"\bannual\s+contract\s+value\b", re.I),
        re.compile(r"\bannual\s+fee\b", re.I),
        re.compile(r"\bper\s+unit\s+per\s+month\b", re.I),
        re.compile(r"\bcontract\s+energy\s+management\b", re.I),
        re.compile(r"\bresponse\s+maintenance\s+services\b", re.I),
        re.compile(r"\bplanned\s+preventative\s+maintenance\b", re.I),
        re.compile(r"\bmanagement,\s*data\s*&\s*credit\s+billing\s+services\b", re.I),
        re.compile(r"\bsite\s*-\s*", re.I),
    ]

    FUEL_MARKERS = [
        re.compile(r"\bfuel\s+delivery\b", re.I),
        re.compile(r"\bfuel\s+supply\b", re.I),
        re.compile(r"\boil\s+delivery\b", re.I),
        re.compile(r"\bred\s+diesel\b", re.I),
        re.compile(r"\bbulk\s+fuel\b", re.I),
        re.compile(r"\bgenerator\s+fuel\b", re.I),
    ]

    DOMESTIC_MARKERS = [
        re.compile(r"\bflat\s+\d+\b", re.IGNORECASE),
        re.compile(r"\bapt\s+\d+\b", re.IGNORECASE),
        re.compile(r"\bapartment\b", re.IGNORECASE),
        re.compile(r"\bamarelle\b", re.IGNORECASE),
        re.compile(r"\bsantina\b", re.IGNORECASE),
        re.compile(r"\brainier\b", re.IGNORECASE),
        re.compile(r"\bmaraschino\b", re.IGNORECASE),
    ]

    PLANT_ROOM_MARKERS = [
        re.compile(r"\bplant\b", re.I),
        re.compile(r"\benergy\s+centre\b", re.I),
        re.compile(r"\bcommunal\b", re.I),
        re.compile(r"\bheat\s+station\b", re.I),
        re.compile(r"\bchw\b", re.I),
        re.compile(r"\blthw\b", re.I),
        re.compile(r"\bwater\s+sample\b", re.I),
        re.compile(r"\bwater\s+technician\b", re.I),
        re.compile(r"\bpumps?\b", re.I),
        re.compile(r"\bbms\s+panel\b", re.I),
        re.compile(r"\bvalves?\b", re.I),
        re.compile(r"\bactuators?\b", re.I),
    ]

    # --- Isolated Dosing Categories ---
    DOSING_MARKERS = [
        re.compile(r"\bdosing\b", re.I),
        re.compile(r"\bdose\b", re.I),
        re.compile(r"\binhibitor\b", re.I),
        re.compile(r"\bbiocide\b", re.I),
    ]

    BOILER_MARKERS = [
        re.compile(r"\bboiler(?:s|room|house)?\b", re.IGNORECASE),
        re.compile(r"\bhamworthy\b", re.IGNORECASE),
    ]

    NOT_MORELLO_MARKERS = [
        re.compile(r"\bthe\s+dumont\b", re.IGNORECASE),
        re.compile(r"\bthe\s+corniche\b", re.IGNORECASE),
        re.compile(r"\bmono\s+tower\b", re.IGNORECASE),
    ]

    PIPEWORK_MARKERS = [
        re.compile(r"\bmapress\b", re.I),
        re.compile(r"\bcarbon\s+pipe\b", re.I),
        re.compile(r"\bsteel\s+pipe\b", re.I),
        re.compile(r"\bcopper\s+pipe\b", re.I),
        re.compile(r"\bcoupling\b", re.I),
        re.compile(r"\breducer\b", re.I),
        re.compile(r"\btee\b", re.I),
    ]


# =============================================================================
# 2. TEXT EXTRACTION & CLEANING
# =============================================================================

def clean_text(text: str) -> str:
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def get_lines(text: str):
    return [x.strip() for x in text.splitlines() if x.strip()]

def normalise_date(date_str: str) -> str:
    if not date_str: return ""
    try:
        clean_str = re.sub(r"(?<=\d)(st|nd|rd|th)\b", "", date_str.lower())
        d = pd.to_datetime(clean_str, dayfirst=True, errors="coerce")
        return d.strftime("%d/%m/%y") if pd.notna(d) else date_str
    except Exception:
        return date_str

def extract_text_pdf(pdf_path: Path) -> str:
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            if txt := page.extract_text(layout=True):
                if txt.strip(): parts.append(txt)
    return clean_text("\n".join(parts))

def extract_text_ocr(pdf_path: Path) -> str:
    images = convert_from_path(str(pdf_path), dpi=300)
    parts = []
    for img in images:
        if txt := pytesseract.image_to_string(img, config=Config.TESSERACT_CONFIG):
            if txt.strip(): parts.append(txt)
    return clean_text("\n".join(parts))


# =============================================================================
# 3. PARSING LOGIC & CANDIDATE MATCHING
# =============================================================================

def pick_best(candidates, key="value"):
    return next(((c[key] or "").strip() for c in candidates if (c[key] or "").strip()), "")

def candidate_invoice_numbers(lines):
    candidates = []
    text = "\n".join(lines)

    for pat in RegexPatterns.INVOICE_FALLBACKS:
        for m in re.finditer(pat, text, re.I | re.M):
            if val := (m.group(1) or "").strip():
                candidates.append({"value": val})

    for i, line in enumerate(lines):
        if not re.search(RegexPatterns.INVOICE_STRICT_LABEL, line, re.I):
            continue

        same_line = re.search(RegexPatterns.INVOICE_STRICT_LABEL + r"\s*[:.\-]?\s*(" + RegexPatterns.DOCNO_VALUE + r")", line, re.I)
        next_line = lines[i + 1] if i + 1 < len(lines) else ""

        if same_line:
            candidates.append({"value": same_line.group(1)})
            continue

        if next_match := re.search(r"^" + RegexPatterns.DOCNO_VALUE + r"$", next_line.strip(), re.I):
            candidates.append({"value": next_match.group(0)})
            continue

        if i + 2 < len(lines):
            joined_next = f"{next_line} {lines[i+2]}".strip()
            if joined_match := re.search(r"\b(" + RegexPatterns.DOCNO_VALUE + r")\b", joined_next, re.I):
                candidates.append({"value": joined_match.group(1)})

    seen, cleaned = set(), []
    for c in candidates:
        v = (c.get("value") or "").strip().upper()
        if v and not v.lower().startswith("invoice") and v not in seen:
            seen.add(v)
            cleaned.append({"value": v})
    return cleaned

def candidate_credit_numbers(lines):
    candidates = []
    for i, line in enumerate(lines):
        if not re.search(RegexPatterns.CREDIT_STRICT_LABEL, line, re.I):
            continue
        same_line = re.search(RegexPatterns.CREDIT_STRICT_LABEL + r"\s*[:.\-]?\s*(" + RegexPatterns.DOCNO_VALUE + r")", line, re.I)
        next_line = lines[i + 1] if i + 1 < len(lines) else ""

        if same_line:
            candidates.append({"value": same_line.group(1)})
        elif next_match := re.search(r"^" + RegexPatterns.DOCNO_VALUE + r"$", next_line.strip(), re.I):
            candidates.append({"value": next_match.group(0)})
    return candidates

def candidate_credited_invoice_numbers(lines):
    text = "\n".join(lines)
    candidates = []
    for pat in RegexPatterns.CREDITED_INVOICE_PATTERNS:
        for m in re.finditer(pat, text, re.I | re.M):
            candidates.append({"value": m.group(1).strip()})
    return candidates

def candidate_dates(lines):
    candidates = []
    for i, line in enumerate(lines):
        if any(re.search(pat, line, re.I) for pat in RegexPatterns.DATE_FALSE_LABELS): continue
        if not re.search(RegexPatterns.DATE_STRICT_LABEL, line, re.I): continue

        same_line = next((re.search(pat, line) for pat in RegexPatterns.DATE_VALUES if re.search(pat, line)), None)
        next_line = lines[i + 1] if i + 1 < len(lines) else ""
        next_match = None

        if not same_line and not any(re.search(pat, next_line, re.I) for pat in RegexPatterns.DATE_FALSE_LABELS):
            next_match = next((re.search(pat, next_line) for pat in RegexPatterns.DATE_VALUES if re.search(pat, next_line)), None)

        value = same_line.group(0) if same_line else (next_match.group(0) if next_match else "")
        if value:
            candidates.append({
                "anchor_line": line,
                "value": re.sub(r"^(\d{1,2})\.\s+", r"\1 ", value.strip())
            })
    return candidates

def candidate_amounts(lines):
    candidates = []
    for i, line in enumerate(lines):
        txt = (line or "").strip()
        matched_label = next((label for label, pat in RegexPatterns.AMOUNT_PRIORITIES if re.search(pat, txt.lower(), re.I)), None)

        if not matched_label or any(re.search(p, txt, re.I) for p in RegexPatterns.AMOUNT_REJECTS):
            continue

        raw_vals = re.findall(r"-?\s*[£$€]?\s*[\d,]+\.\d{2}", txt)
        value = re.sub(r"[£$€\s,]", "", raw_vals[-1]) if raw_vals else ""

        if not value and i + 1 < len(lines):
            next_line = (lines[i + 1] or "").strip()
            if not any(re.search(p, next_line, re.I) for p in RegexPatterns.AMOUNT_REJECTS):
                next_vals = re.findall(r"-?\s*[£$€]?\s*[\d,]+\.\d{2}", next_line)
                if next_vals: value = re.sub(r"[£$€\s,]", "", next_vals[-1])

        if value:
            candidates.append({"label": matched_label, "value": value})
    return candidates

def candidate_detail_blocks(lines):
    candidates = []
    i = 0
    while i < len(lines):
        line = (lines[i] or "").strip()
        clean_line = re.sub(r"\s+", " ", line)
        if any(re.search(p, clean_line.lower(), re.I) for p in RegexPatterns.DETAIL_HEADERS):
            block = []
            j = i + 1
            while j < len(lines):
                nxt = re.sub(r"\s+", " ", (lines[j] or "")).strip()
                if nxt and any(re.search(p, nxt.lower(), re.I) for p in RegexPatterns.DETAIL_STOPS):
                    break
                if nxt:
                    block.append(nxt)
                j += 1
            if block:
                candidates.append({"value": "\n".join(block).strip()})
            i = j
        else:
            i += 1
    return candidates

def pick_best_detail(candidates):
    combined_filtered = []
    for c in candidates:
        if not (value := (c.get("value") or "").strip()): continue
        for ln in [x for x in value.splitlines() if x.strip()]:
            low = ln.strip().lower()
            if low in {"charges", "charge", "description", "detail", "item detail", "price"}:
                continue
            if any(re.search(p, low, re.I) for p in RegexPatterns.DETAIL_HEADERS):
                continue
            if re.fullmatch(r"\d{2}/\d{2}/\d{2,4}", low):
                continue
            if re.search(r"^(qty|quantity|unit|unit price|price|value|vat rate|vat identifier|discount %|receipt date|measure)$", low, re.I):
                continue
            clean_line = re.sub(r"^([A-Za-z0-9.-]+)\s{3,}", "", ln.lstrip()).strip()
            if clean_line:
                combined_filtered.append(clean_line)
    return "\n".join(combined_filtered).strip()

def pick_best_date(candidates):
    preferred = ["tax point", "document date", "credit date", "invoice date", "issued", "date"]
    for label in preferred:
        for c in candidates:
            if c["value"] and label in c["anchor_line"].lower():
                return c["value"]
    return pick_best(candidates)

def pick_best_amount(candidates):
    preferred = ["total_incl_vat", "total_inc_vat", "total_including_vat", "grand_total", "invoice_total", "amount_due", "balance_due", "total"]
    for wanted in preferred:
        for c in reversed(candidates):
            if c.get("label") == wanted and c.get("value"):
                return c["value"]
    return ""

def detect_document_type(lines, credit_number, invoice_number):
    text = "\n".join(lines[:120]).lower()
    if credit_number or re.search(r"\bcredit\s+(note|memo)\b", text): return "credit_note"
    if invoice_number or re.search(r"\binvoice\b", text): return "invoice"
    return "unknown"

def detect_flat_site(detail_text: str) -> str:
    text = (detail_text or "").lower()
    site_map = {
        "Amarelle": [r"\bamarelle\b"],
        "Rainier": [r"\brainier\b"],
        "Santina": [r"\bsantina\b"],
        "Maraschino": [r"\bmaraschino\b"],
    }
    for site_name, patterns in site_map.items():
        if any(re.search(p, text, re.I) for p in patterns):
            return site_name
    return "Unknown"


# =============================================================================
# 4. CORE ORCHESTRATION & LINE ITEMS
# =============================================================================

def extract_line_items(detail_text: str):
    items, matched_time_charge_descs = [], set()
    lines = [line.strip() for line in (detail_text or "").splitlines() if line.strip()]

    clean_amt = lambda val: re.sub(r"[£$€\s,]", "", (val or "").strip())
    norm_space = lambda s: re.sub(r"\s+", " ", (s or "").strip())

    def is_valid_charge(desc, amt_str):
        if len(desc) < 5: return False
        if "DOMESTIC COMMISSIONING" in desc and not any(char.isdigit() for char in desc):
            return False
        return True

    SPLIT_ANCHORS = [
        # FIXED: explicitly checks for both, and refuses to split them apart
        r"(?=Plant\s+Planned\s+Preventative\s+Maintenance)",
        r"(?<!Plant\s)(?=Planned\s+Preventative\s+Maintenance)",

        r"(?=Prepayment\s+Services,\s*Incontro)",
        r"(?=Response\s+Maintenance\s+Services)",
        r"(?=Credit\s+Billing\s+Services)",
        r"(?=Debt\s+Management)",
        r"(?=Hire\s+Contract)",
        r"(?=Management\s+Service)",
        r"(?=Data\s+Services)",
        r"(?=Property\s+Cyclical\s+Maintenance)",
        r"(?=Property\s+Standby\s+Engineer)"
    ]
    split_pattern = re.compile("|".join(SPLIT_ANCHORS))

    raw_processed_lines = []
    for line in lines:
        if any(re.search(p, line, re.I) for p in RegexPatterns.LINE_ITEM_SKIP + RegexPatterns.LINE_ITEM_SUMMARY):
            continue
        parts = [p.strip() for p in split_pattern.split(line) if p.strip()]
        raw_processed_lines.extend(parts)

    for line in raw_processed_lines:
        matched = False

        # Pattern 1: Time Charges
        if m := re.match(r"^(Time charge(?: using .*?)?\s*\(.*?\)).*?(?<![\d.])-?[£$€]?\s*(-?[\d,]+\.\d{2})\s*$", line, re.I):
            desc = m.group(1).strip()
            matched_time_charge_descs.add(norm_space(desc).lower())
            items.append({"line_description": desc, "line_amount": clean_amt(m.group(2))})
            matched = True
            continue

        if norm_space(line).lower() in matched_time_charge_descs:
            continue

        # Pattern 2: Standard Items
        if m := re.match(r"^(.*?\D)\s+(\d+(?:\.\d+)?)\s+[£$€]?\s*(-?[\d,]+\.\d{2})\s+[£$€]?\s*(-?[\d,]+\.\d{2})(?:\s+\d+(?:\.\d+)?\s*%?)?\s*$", line, re.I):
            desc = f"{m.group(1).strip()} {m.group(2).strip()}".strip()
            items.append({"line_description": desc, "line_amount": clean_amt(m.group(4))})
            matched = True
            continue

        # Pattern 3: Dates mixed with values
        if m := re.search(r"^(.*?\S)\s+\d+(?:\.\d+)?\s+\d{2}/\d{2}/\d{2,4}\s+[A-Za-z]+\s+(-?\d+(?:\.\d{2})?)\s+[A-Za-z0-9]+\s+(-?\d+(?:\.\d{2})?)\s*$", line, re.I):
            desc = m.group(1).strip()
            items.append({"line_description": desc, "line_amount": clean_amt(m.group(3))})
            matched = True
            continue

        # Fallback Pattern
        fallback_match = re.search(r"(?<![\d.])-?\s*[£$€]?\s*(-?[\d,]+\.\d{2})(?:\s*(?:[A-Za-z]{1,3}\d?|\d+(?:\.\d+)?\s*%))?\s*$", line)
        if fallback_match and "=" not in line:
            amt = clean_amt(fallback_match.group(1))
            desc = line[:fallback_match.start()].strip()
            desc = re.sub(r"[\s=:\-]+$", "", desc).strip()
            if is_valid_charge(desc, amt):
                items.append({"line_description": desc, "line_amount": amt})
                matched = True
                continue

        zero_fee_match = re.search(r"Annual\s+Fee\s*=\s*[£$€]?\s*0\.00", line, re.I)

        if not matched and zero_fee_match:
            items.append({"line_description": line, "line_amount": "0.00"})
        elif not matched and items and len(line) > 2:
            skip_patterns = [
                r"^(Invoice|Credit)\s+No",
                r"\b(schedule|contract|covering|page)\b",
                r"^Planned\s+Preventative\s+Maintenance",
                r"^Prepayment\s+Services",
                r"^Response\s+Maintenance\s+Services",
                r"^Credit\s+Billing\s+Services",
            ]
            if any(re.search(p, line, re.I) for p in skip_patterns):
                continue
            items[-1]["line_description"] += f" {line}"

    return items

def parse_lines(lines):
    if not lines: return {}
    invoice_candidates = candidate_invoice_numbers(lines)
    credited_invoice_number = pick_best(candidate_credited_invoice_numbers(lines))
    invoice_number = next(((c.get("value") or "").strip() for c in invoice_candidates if (c.get("value") or "").strip() and (c.get("value") or "").strip() != credited_invoice_number), "")
    return {
        "invoice_number": invoice_number or credited_invoice_number,
        "credit_number": pick_best(candidate_credit_numbers(lines)),
        "invoice_date": normalise_date(pick_best_date(candidate_dates(lines))),
        "invoice_amount": pick_best_amount(candidate_amounts(lines)),
        "detail": pick_best_detail(candidate_detail_blocks(lines)),
        "lines_ref": lines,
    }

def has_any(patterns, text: str) -> bool:
    text = text or ""
    return any(pattern.search(text) for pattern in patterns)

# =============================================================================
# 5. EXECUTION & REPORTING
# =============================================================================

def debug_dump(pdf_path, parsed, text_pdf="", text_ocr="", chosen_source=""):
    Config.DEBUG_DIR.mkdir(exist_ok=True)
    out_file = Config.DEBUG_DIR / f"{pdf_path.stem.replace('/', '_').replace(chr(92), '_')}.txt"
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(f"FILE: {pdf_path}\nCHOSEN_SOURCE: {chosen_source}\n\n{'='*80}\nPARSED SUMMARY\n{'='*80}\n")
        for k in ['invoice_number', 'credit_number', 'invoice_date', 'invoice_amount']:
            f.write(f"{k}: {parsed.get(k, '')}\n")
        f.write(f"\n{'='*80}\nDETAIL\n{'='*80}\n{parsed.get('detail', '')}\n")

        line_items = extract_line_items(parsed.get("detail", ""))
        line_sum, vat_sum = 0.0, 0.0
        f.write(f"\n{'='*80}\nLINE ITEMS FROM DETAIL\n{'='*80}\n")
        for idx, item in enumerate(line_items, 1):
            f.write(f"[{idx}] DESC: {item.get('line_description', '')}\n    AMT : {item.get('line_amount', '')}\n")
            try:
                amt = float(str(item.get("line_amount", "")).replace(",", "").strip())
                line_sum += amt
                vat_sum += round(amt * 0.20, 2)
            except ValueError: pass

        f.write(f"\n{'='*80}\nTOTAL CHECK\n{'='*80}\n")
        f.write(f"line_amount_sum: {line_sum:.2f}\nvat_amount_sum: {vat_sum:.2f}\ncalculated_total: {line_sum + vat_sum:.2f}\ninvoice_amount: {parsed.get('invoice_amount', '')}\n")
        f.write(f"\n{'='*80}\nLINES_REF\n{'='*80}\n")
        for i, line in enumerate(parsed.get("lines_ref", []), 1): f.write(f"{i:04d}: {line}\n")
        if text_pdf: f.write(f"\n{'='*80}\nRAW PDF TEXT\n{'='*80}\n{text_pdf}\n")
        if text_ocr: f.write(f"\n{'='*80}\nRAW OCR TEXT\n{'='*80}\n{text_ocr}\n")


def extract_and_parse(pdf_path: Path):
    text_pdf = extract_text_pdf(pdf_path)
    parsed_pdf = parse_lines(get_lines(text_pdf))
    has_data = any([parsed_pdf.get("invoice_number"), parsed_pdf.get("credit_number"), parsed_pdf.get("invoice_amount")])
    use_ocr = len(text_pdf.strip()) < Config.OCR_TEXT_THRESHOLD or not has_data or (not parsed_pdf.get("detail") and len(text_pdf) > 100)
    parsed_final, text_ocr, chosen_source = parsed_pdf, "", "pdf"

    if use_ocr:
        text_ocr = extract_text_ocr(pdf_path)
        parsed_ocr = parse_lines(get_lines(text_ocr))
        score = lambda p: sum(1 for k in ["invoice_number", "credit_number", "invoice_date", "invoice_amount", "detail"] if p.get(k))
        if score(parsed_ocr) >= score(parsed_pdf) and score(parsed_ocr) > 0:
            parsed_final, chosen_source = parsed_ocr, "ocr"
        elif not parsed_pdf.get("detail") and parsed_ocr.get("detail"):
            parsed_pdf["detail"] = parsed_ocr["detail"]
            parsed_final, chosen_source = parsed_pdf, "pdf+ocr_detail"

    debug_dump(pdf_path, parsed_final, text_pdf, text_ocr, chosen_source)
    return parsed_final

def process_pdf(pdf_path_str: str):
    pdf = Path(pdf_path_str)
    rel_path = str(pdf.relative_to(Config.BASE_DIR))
    try:
        parsed = extract_and_parse(pdf)
        doc_type = detect_document_type(parsed.get("lines_ref", []), parsed.get("credit_number"), parsed.get("invoice_number"))

        base_row = {
            "file": rel_path,
            "document_type": doc_type,
            "invoice_date": parsed.get("invoice_date", ""),
            "invoice_number": parsed.get("invoice_number", ""),
            "credit_number": parsed.get("credit_number", ""),
            "invoice_amount": parsed.get("invoice_amount", ""),
            "detail": parsed.get("detail", ""),
            "category": "",
            "in_scope": "",
        }

        missing_fields = [k for k in ["invoice_number", "invoice_date", "invoice_amount", "detail"] if not base_row[k]]
        line_items = extract_line_items(base_row["detail"])
        rows = []

        if not line_items:
            row = {
                **base_row,
                "line_description": "",
                "line_amount": "",
                "vat_amount": "",
                "line_total": "",
                "line_index": 0
            }
            missing_fields.extend(["line_description", "line_amount"])
            rows.append(row)
        else:
            # --- SMART HEADER CHECK (The first 10 lines) ---
            # Identifies dedicated Job Sheets without triggering on buried schedule lines
            detail_lines = [ln.strip() for ln in (base_row.get("detail") or "").splitlines() if ln.strip()]
            detail_header = "\n".join(detail_lines[:10]).lower()
            detail_lower = (base_row.get("detail") or "").lower()

            global_category = ""
            is_not_morello = has_any(RegexPatterns.NOT_MORELLO_MARKERS, detail_lower)

            if has_any(RegexPatterns.DOMESTIC_MARKERS, detail_header):
                global_category = "Domestic"
            elif has_any(RegexPatterns.PLANT_ROOM_MARKERS, detail_header):
                global_category = "Plant Room"
            elif has_any(RegexPatterns.BOILER_MARKERS, detail_header):
                global_category = "Boiler"
            elif has_any(RegexPatterns.FUEL_MARKERS, detail_header):
                global_category = "Fuel"

            for idx, item in enumerate(line_items):
                try:
                    val_check = float(item["line_amount"])
                    if val_check == 0.0: continue
                except (ValueError, TypeError): pass

                row = {
                    **base_row,
                    "line_description": item["line_description"],
                    "line_amount": item["line_amount"],
                    "vat_amount": "",
                    "line_total": "",
                    "line_index": idx
                }

                desc_lower = item["line_description"].lower()

                if is_not_morello or has_any(RegexPatterns.NOT_MORELLO_MARKERS, desc_lower):
                    row["is_not_morello_file"] = True
                    row["category"] = "Not Morello"
                else:
                    row["is_not_morello_file"] = False

                    # 1. STRICT LINE-LEVEL EXCEPTIONS
                    # Contract phrases and specific tasks always take priority
                    if "preventative" in desc_lower:
                        row["category"] = "Preventative"
                        row["in_scope"] = "Y"
                    elif any(term in desc_lower for term in ["cyclical"]):
                        row["category"] = "Routine"
                    elif any(term in desc_lower for term in ["billing", "debt management"]):
                        row["category"] = "Billing"
                        row["in_scope"] = "N"
                    elif has_any(RegexPatterns.DOSING_MARKERS, desc_lower):
                        row["category"] = "Dosing"

                    # 2. GLOBAL JOB SHEET INHERITANCE
                    # If the invoice header specifies a location ("Flat 41" or "Energy Centre"),
                    # generic line items (like Time Charges or parts) inherit that location.
                    elif global_category:
                        row["category"] = global_category

                    # 3. LINE-LEVEL FALLBACKS FOR MIXED INVOICES
                    # If there's no global header location (e.g. Schedule/Contract invoice),
                    # evaluate the physical markers directly on the line item.
                    elif has_any(RegexPatterns.BOILER_MARKERS, desc_lower):
                        row["category"] = "Boiler"
                    elif has_any(RegexPatterns.DOMESTIC_MARKERS, desc_lower):
                        row["category"] = "Domestic"
                    elif has_any(RegexPatterns.PLANT_ROOM_MARKERS, desc_lower):
                        row["category"] = "Plant Room"
                    elif has_any(RegexPatterns.FUEL_MARKERS, desc_lower) or "fuel" in desc_lower:
                        row["category"] = "Fuel"
                    else:
                        row["category"] = ""

                try:
                    amt_float = float(item["line_amount"])
                    row.update({
                        "line_amount": f"{amt_float:.2f}",
                        "vat_amount": f"{amt_float * 0.20:.2f}",
                        "line_total": f"{amt_float * 1.20:.2f}"
                    })
                except ValueError: pass
                rows.append(row)

        return {"pdf": rel_path, "rows": rows, "error": None, "missing_fields": missing_fields}
    except Exception as e:
        return {"pdf": rel_path, "rows": [], "error": str(e), "missing_fields": []}


# =============================================================================
# 6. FILE SYSTEM UTILITIES
# =============================================================================

def clear_output_directory():
    output_dir = Config.OUTPUT_DIR
    if output_dir.exists():
        print(f"Clearing old extraction assets from: {output_dir}")
        for item in output_dir.iterdir():
            try:
                if item.is_dir(): shutil.rmtree(item)
                else: item.unlink()
            except Exception as e:
                print(f"Warning: Could not remove old artifact {item.name}: {e}")

def duplicate_keep_priority(path_str: str):
    name = Path(path_str).name.lower()
    return (1 if "copy of" in name else 0, len(Path(path_str).parts), len(name), name)

def deduplicate_inputs(pdfs):
    unique_pdfs, seen_filenames, dup_names = [], {}, {}
    for pdf in pdfs:
        rel_path = str(pdf.relative_to(Config.BASE_DIR))
        if pdf.name in seen_filenames:
            if pdf.name not in dup_names:
                dup_names[pdf.name] = {"key": pdf.name, "kept": seen_filenames[pdf.name], "discarded": []}
            dup_names[pdf.name]["discarded"].append(rel_path)
            continue
        seen_filenames[pdf.name] = rel_path
        unique_pdfs.append(pdf)
    return unique_pdfs, dup_names

def ensure_output_dirs():
    Config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    Config.DUPLICATES_DIR.mkdir(parents=True, exist_ok=True)
    Config.INDIVIDUAL_FLATS_DIR.mkdir(parents=True, exist_ok=True)
    Config.NOT_MORELLO_DIR.mkdir(parents=True, exist_ok=True)
    Config.FAILED_DIR.mkdir(parents=True, exist_ok=True)

def safe_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]+', "_", name).strip()

def unique_destination_path(dest_dir: Path, filename: str) -> Path:
    dest_dir.mkdir(parents=True, exist_ok=True)
    candidate = dest_dir / safe_filename(filename)
    if not candidate.exists(): return candidate
    stem, suffix, i = candidate.stem, candidate.suffix, 2
    while True:
        alt = dest_dir / f"{stem} ({i}){suffix}"
        if not alt.exists(): return alt
        i += 1

def invoice_year_from_row(row: dict) -> str:
    raw = (row.get("invoice_date") or "").strip()
    if not raw: return "Unknown Year"
    dt = pd.to_datetime(raw, dayfirst=True, errors="coerce")
    return str(dt.year) if pd.notna(dt) else "Unknown Year"

def is_flat_visit_result(result: dict) -> bool:
    if not result.get("rows"): return False
    return any((r.get("category") or "") == "Domestic" for r in result["rows"])

def copy_pdf_to_output(source_rel_path: str, category: str, row: dict | None = None) -> str:
    source_path = Config.BASE_DIR / source_rel_path
    if category == "duplicate": dest_dir = Config.DUPLICATES_DIR
    elif category == "not_morello": dest_dir = Config.NOT_MORELLO_DIR
    elif category == "failed": dest_dir = Config.FAILED_DIR
    elif category == "flat": dest_dir = Config.INDIVIDUAL_FLATS_DIR / invoice_year_from_row(row or {}) / detect_flat_site((row or {}).get("detail", ""))
    else: dest_dir = Config.OUTPUT_DIR / invoice_year_from_row(row or {})
    dest_path = unique_destination_path(dest_dir, source_path.name)
    shutil.copy2(source_path, dest_path)
    return str(dest_path.relative_to(Config.OUTPUT_DIR))

def make_html_link(rel_to_output: str) -> str:
    return f'<a href="{escape(rel_to_output).replace("\\", "/")}" target="_blank">{escape(rel_to_output)}</a>'


# =============================================================================
# 7. MAIN ENTRY POINT
# =============================================================================

def main():
    clear_output_directory()
    ensure_output_dirs()

    pdfs = sorted(Config.INPUT_DIR.rglob("*.[pP][dD][fF]"))
    if not pdfs:
        print(f"No PDFs found in {Config.INPUT_DIR}")
        return

    unique_pdfs, duplicate_filename_groups = deduplicate_inputs(pdfs)
    duplicate_invoice_groups, seen_invoice_keys = {}, {}
    error_files, missing_data_files = [], []
    kept_results, flat_results, duplicate_results, not_morello_results = [], [], [], []
    failed_files = []  # rel_paths of PDFs that could not be processed

    # =========================================================================
    # FIX: Reconciliation counters — track every input PDF so none disappear.
    # =========================================================================
    total_input = len(pdfs)
    processed_count = 0
    failed_count = 0
    duplicate_count = 0
    classified_count = 0

    with ProcessPoolExecutor(max_workers=min(Config.MAX_WORKERS, len(unique_pdfs))) as executor:
        futures = {executor.submit(process_pdf, str(pdf)): pdf for pdf in unique_pdfs}
        for future in as_completed(futures):
            result = future.result()

            if result["error"]:
                print(f"Failed: {result['pdf']} -> {result['error']}")
                error_files.append({"pdf": result["pdf"], "error": result["error"]})
                failed_files.append(result["pdf"])
                failed_count += 1
                continue

            if not result.get("rows"):
                # No exception was raised, but nothing was extracted at all
                # (e.g. blank scan, OCR returned nothing). Treat as failed so it
                # is visible and moved to Output/Failed instead of disappearing.
                reason = "No text/rows extracted from PDF (empty or unreadable content)"
                print(f"Failed: {result['pdf']} -> {reason}")
                error_files.append({"pdf": result["pdf"], "error": reason})
                failed_files.append(result["pdf"])
                failed_count += 1
                continue

            first_row = result["rows"][0]
            invoice_key = (
                (first_row.get("document_type") or "").strip().lower(),
                (first_row.get("invoice_number") or "").strip().upper(),
                (first_row.get("invoice_amount") or "").strip(),
            )

            # =================================================================
            # FIX: Check data quality BEFORE duplicate detection.
            #
            # The original code checked for invoice-key duplicates whenever
            # invoice_number and invoice_amount were both non-empty, WITHOUT
            # verifying that the extraction was actually good. When multiple
            # PDFs had bad extractions that produced the same wrong
            # invoice_number + invoice_amount (e.g. from common boilerplate
            # text, footers, or OCR artifacts), they were silently treated as
            # duplicates and sent to the Duplicates folder — never appearing
            # in the main output or in Failed.
            #
            # Now: if the PDF has ANY missing critical fields, we skip the
            # duplicate check entirely. The PDF falls through to normal
            # classification and appears in the output with its missing data
            # flagged in the exceptions report.
            #
            # This ensures every input PDF is accounted for in the output.
            # =================================================================
            if result.get("missing_fields"):
                # Extraction is incomplete — do NOT treat this as a duplicate
                # candidate even if invoice_number/amount happen to match.
                # Fall through to normal classification below.
                pass
            elif invoice_key[1] and invoice_key[2]:
                if invoice_key in seen_invoice_keys:
                    group_key = f"{invoice_key[0]} | {invoice_key[1]} | {invoice_key[2]}"
                    if group_key not in duplicate_invoice_groups:
                        duplicate_invoice_groups[group_key] = {"key": group_key, "kept": seen_invoice_keys[invoice_key]["pdf"], "discarded": []}
                    duplicate_invoice_groups[group_key]["discarded"].append(result["pdf"])
                    duplicate_results.append(result)
                    duplicate_count += 1
                    print(f"Skipped duplicate invoice: {result['pdf']}")
                    continue
                seen_invoice_keys[invoice_key] = result

            if any(r.get("is_not_morello_file") for r in result["rows"]):
                not_morello_results.append(result)
                classified_count += 1
                print(f"Classified as Outside Site (Not Morello): {result['pdf']}")
            elif is_flat_visit_result(result):
                flat_results.append(result)
                classified_count += 1
                print(f"Classified as Individual Flat: {result['pdf']}")
            else:
                kept_results.append(result)
                classified_count += 1

            if result.get("missing_fields"):
                missing_data_files.append({"pdf": result["pdf"], "missing": result["missing_fields"]})
            processed_count += 1
            print(f"Processed: {result['pdf']}")

        # =========================================================================
        # FIX: Reconciliation print — verify every input PDF is accounted for.
        # =========================================================================
        filename_dup_count = sum(len(g["discarded"]) for g in duplicate_filename_groups.values())
        total_accounted = (
            classified_count
            + failed_count
            + duplicate_count
            + filename_dup_count
        )
        print(f"\n{'='*60}")
        print(f"RECONCILIATION")
        print(f"{'='*60}")
        print(f"  Total input PDFs:              {total_input}")
        print(f"  Filename duplicates (to Dup):   {filename_dup_count}")
        print(f"  Unique PDFs processed:          {len(unique_pdfs)}")
        print(f"    - Classified (output):       {classified_count}")
        print(f"    - Invoice-key dups (to Dup):  {duplicate_count}")
        print(f"    - Failed (to Failed):         {failed_count}")
        print(f"  Total accounted for:            {total_accounted + filename_dup_count}")
        if total_accounted + filename_dup_count != total_input:
            print(f"  *** WARNING: MISMATCH! {total_input - total_accounted - filename_dup_count} PDFs unaccounted for! ***")
        print(f"{'='*60}\n")

        output_path_map = {}
        for result in kept_results:
            output_rel = copy_pdf_to_output(result["pdf"], "kept", result["rows"][0])
            output_path_map[result["pdf"]] = output_rel
            for r in result["rows"]: r["file"] = output_rel

        for result in flat_results:
            output_rel = copy_pdf_to_output(result["pdf"], "flat", result["rows"][0])
            output_path_map[result["pdf"]] = output_rel
            for r in result["rows"]: r["file"] = output_rel

        for result in not_morello_results:
            output_rel = copy_pdf_to_output(result["pdf"], "not_morello", result["rows"][0])
            output_path_map[result["pdf"]] = output_rel
            for r in result["rows"]: r["file"] = output_rel

        for result in duplicate_results:
            output_path_map[result["pdf"]] = copy_pdf_to_output(result["pdf"], "duplicate")

        # --- Move filename duplicates into Output/Duplicates ---
        for group in duplicate_filename_groups.values():
            for discarded_pdf in group["discarded"]:
                try:
                    output_path_map[discarded_pdf] = copy_pdf_to_output(discarded_pdf, "duplicate")
                    print(f"Moved filename duplicate to Duplicates folder: {discarded_pdf}")
                except Exception as e:
                    print(f"Warning: could not move filename duplicate {discarded_pdf}: {e}")

        # --- Move failed PDFs into Output/Failed ---
        for rel_path in failed_files:
            try:
                output_path_map[rel_path] = copy_pdf_to_output(rel_path, "failed")
                print(f"Moved failed PDF to Failed folder: {rel_path}")
            except Exception as e:
                print(f"Warning: could not move failed PDF {rel_path} to Failed folder: {e}")

        for g in [duplicate_filename_groups, duplicate_invoice_groups]:
            for k, group in g.items():
                group["kept"] = output_path_map.get(group["kept"], group["kept"])
                group["discarded"] = [output_path_map.get(x, x) for x in group["discarded"]]

        for item in error_files + missing_data_files:
            item["pdf"] = output_path_map.get(item["pdf"], item["pdf"])

        rows = []
        for res in [kept_results, flat_results, not_morello_results]:
            for r in res: rows.extend(r["rows"])

    if rows:
        df = pd.DataFrame(rows)

        final_cols = [
            "file", "document_type", "invoice_date", "invoice_number",
            "credit_number", "invoice_amount", "detail", "line_description", "category",
            "line_amount", "vat_amount", "line_total", "in_scope"
        ]

        df = df[final_cols + ["line_index"]].sort_values(["file", "line_index"]).reset_index(drop=True)
        df = df[final_cols]

        write_excel_report(df, Config.OUTPUT_XLSX)
        print(f"Extraction complete! Excel workbook saved to {Config.OUTPUT_XLSX}")

        generate_dashboard_html(df, Config.OUTPUT_HTML, make_html_link)
        mismatches = build_total_mismatch_files(df)

        create_exceptions_report(
            duplicate_filename_groups, duplicate_invoice_groups,
            error_files, missing_data_files, mismatches, Config.EXCEPTIONS_HTML,
        )
    else:
        print("No rows extracted")
        create_exceptions_report(duplicate_filename_groups, duplicate_invoice_groups, error_files, missing_data_files, [], Config.EXCEPTIONS_HTML)

if __name__ == "__main__":
    main()