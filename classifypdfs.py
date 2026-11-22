#!/usr/bin/env python3

import re
from pathlib import Path

import pandas as pd
import pdfplumber
import pytesseract
from pdf2image import convert_from_path

from ocr_deps import ensure_ocr_ready


BASE_DIR = Path(__file__).resolve().parent
YEAR_FOLDERS = ["2019", "2021", "2022"]

OUTPUT_DIAG = BASE_DIR / "invoice_diagnostic.csv"
OUTPUT_LAYOUTS = BASE_DIR / "invoice_layout_summary.csv"

OCR_TEXT_THRESHOLD = 80


DATE_VALUE_PATTERNS = [
    r"\b\d{4}-\d{2}-\d{2}\b",
    r"\b\d{2}/\d{2}/\d{4}\b",
    r"\b\d{2}/\d{2}/\d{2}\b",
    r"\b\d{2}-\d{2}-\d{4}\b",
    r"\b\d{1,2}\.?\s+[A-Za-z]{3,9}\s+\d{4}\b",  # <-- Added \.? here to catch "21."
]

# Added document date, tax point, and issued. 
# Removed ^\s* so dates can be found anywhere on the line, acting as a flexible keyword search
STRICT_DATE_LABEL_PATTERN = r"\b(credit\s+date|invoice\s+date|document\s+date|tax\s+point|issued|date)\b"

FALSE_DATE_PATTERNS = [
    r"date of invoice",
    r"days from date of invoice",
    # posting date removed from here
]

DOCNO_VALUE_PATTERN = r"\b(?=[A-Z0-9./-]{2,40}\b)(?=[A-Z0-9./-]*\d)[A-Z0-9][A-Z0-9./-]{1,40}\b"

STRICT_INVOICE_LABEL_PATTERN = (
    r"(?:invoice\s*number|invoice\s*no\.?|invoice\s*#|\binv\s*number\b|\binv\s*no\.?\b|\binv\s*#)"
)
# Added credit memo no to the strict label search
STRICT_CREDIT_LABEL_PATTERN = r"(?:credit\s*number|credit\s*no\.?|credit\s*#|credit\s*memo\s*no\.?)"

STRUCTURAL_PATTERNS = [
    r"switch2 energy limited",
    r"accounts payable",
    r"all correspondence to be sent to",
    r"ground floor west",
    r"the waterfront",
    r"salts mill road",
    r"shipley",
    r"invoice",
    r"credit note",
    r"invoice date",
    r"credit date",
    r"date",
    r"invoice number",
    r"invoice no",
    r"credit number",
    r"credit no",
    r"credit against invoice",
    r"credit for invoice",
    r"item detail quantity unit line value",
    r"detail quantity unit line value",
    r"detail quantity line value",
    r"item detail",
    r"description",
    r"detail",
    r"job total",
    r"total ex vat",
    r"grand total",
    r"total due",
    r"amount due",
    r"balance due",
    r"vat",
    r"total",
]

DETAIL_HEADER_PATTERNS = [
    # Long headers can appear anywhere on the line if there are multiple columns
    r"\bitem\s+detail\s+quantity\s+unit\s+line\s+value\b",
    r"\bdetail\s+quantity\s+unit\s+line\s+value\b",
    r"\bdetail\s+quantity\s+line\s+value\b",
    r"\bno\.\s+description\s+quantity\s+unit\s+price\s+amount\b",
    r"\bno\.\s+description.*",
    
    # Short headers should be strongly anchored to avoid matching regular sentences
    r"^[\W_]*detail[\W_]*$",
    r"^[\W_]*item\s+detail[\W_]*$",
    r"^[\W_]*description[\W_]*$",
    r"^[\W_]*item\s+description[\W_]*$",
    r"^[\W_]*charges[\W_]*$",  # <-- Added to catch the Charges blocks
]

# Updated to use ^[\W_]* to bypass punctuation/line artifacts
DETAIL_STOP_PATTERNS = [
    r"^[\W_]*total\s+ex\s+vat\b",
    r"^[\W_]*total\s*(?:£)?\s*excl\.?\s*vat\b",
    r"^[\W_]*total\s*(?:£)?\s*excluding\s*vat\b",
    r"^[\W_]*vat\b",
    r"^[\W_]*total\b",
    r"^[\W_]*grand\s+total\b",
    r"^[\W_]*amount\s+due\b",
    r"^[\W_]*balance\s+due\b",
    r"^[\W_]*job\s+total\b",
    r"^[\W_]*invoice\s+total\b",
    r"^[\W_]*terms\b",
    r"^[\W_]*invoice\s*(number|no\.?|#)\b",
    r"^[\W_]*credit\s*(number|no\.?|#)\b",
    r"^[\W_]*invoice\s+date\b",
    r"^[\W_]*credit\s+date\b",
]

# Removed the ^[\W_]* anchor so labels can be found anywhere on the line
AMOUNT_PRIORITY_PATTERNS = [
    ("total_incl_vat", r"\btotal\s*(?:£)?\s*(?:incl\.?|inc\.?)\s*vat\b"),
    ("total_inc_vat", r"\btotal\s*(?:£)?\s*inc\.?\s*vat\b"),
    ("total_including_vat", r"\btotal\s*(?:£)?\s*including\s*vat\b"),
    ("grand_total", r"\bgrand\s+total\b"),
    ("amount_due", r"\bamount\s+due\b"),
    ("balance_due", r"\bbalance\s+due\b"),
    ("invoice_total", r"\binvoice\s+total\b"),
    ("total", r"\btotal\b"),
]

# Kept the start anchor (^) on single words like "VAT" or "Net" 
# so we don't accidentally reject a line that says "Total including VAT"
AMOUNT_REJECT_PATTERNS = [
    r"\btotal\s+ex\s+vat\b",
    r"\btotal\s*(?:£)?\s*excl\.?\s*vat\b",
    r"\btotal\s*(?:£)?\s*excluding\s*vat\b",
    r"^[\W_]*vat\b",
    r"^[\W_]*sub\s*total\b",
    r"^[\W_]*subtotal\b",
    r"^[\W_]*net\b",
    r"^[\W_]*unit\b",
    r"^[\W_]*rate\b",
]

def list_pdfs():
    pdfs = []
    for folder in YEAR_FOLDERS:
        p = BASE_DIR / folder
        if p.exists() and p.is_dir():
            pdfs.extend(sorted(p.rglob("*.pdf")))
    return pdfs


def clean_text(text: str) -> str:
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_text_pdf(pdf_path: Path) -> str:
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Added layout=True to stop pdfplumber from mashing text columns together
            txt = page.extract_text(layout=True) or ""
            if txt.strip():
                parts.append(txt)
    return clean_text("\n".join(parts))

def extract_text_ocr(pdf_path: Path) -> str:
    images = convert_from_path(str(pdf_path), dpi=300)
    parts = []
    for img in images:
        txt = pytesseract.image_to_string(img)
        if txt.strip():
            parts.append(txt)
    return clean_text("\n".join(parts))


def get_lines(text: str):
    return [x.strip() for x in text.splitlines() if x.strip()]


def normalize_preview_line(line: str) -> str:
    line = line.strip().lower()
    line = line.replace("’", "'").replace("“", '"').replace("”", '"')
    line = re.sub(r"[,:;]", " ", line)
    line = re.sub(r"\.+", " ", line)
    line = re.sub(r"\s+", " ", line).strip()

    line = re.sub(r"\b\d{4}-\d{2}-\d{2}\b", "<date>", line)
    line = re.sub(r"\b\d{2}/\d{2}/\d{4}\b", "<date>", line)
    line = re.sub(r"\b\d{2}/\d{2}/\d{2}\b", "<date>", line)  # <-- NEW NORMALIZER ADDED HERE    
    line = re.sub(r"\b\d{2}-\d{2}-\d{4}\b", "<date>", line)
    line = re.sub(r"\b\d{1,2}\.?\s+[a-z]{3,9}\s+\d{4}\b", "<date>", line)

    line = re.sub(r"-?[£$€]?\s*[\d,]+\.\d{2}", "<amount>", line)
    line = re.sub(r"-?[£$€]?\s*[\d,]+", "<amount>", line)

    line = re.sub(r"\bbd\d{1,2}\s*[0-9a-z]{2,4}\b", "<postcode>", line, flags=re.I)

    line = re.sub(
        r"((?:credit|invoice|order|ref(?:erence)?)\s*(?:number|no\.?|#)\s*)[a-z0-9./-]+",
        r"\1<docno>",
        line,
        flags=re.I,
    )

    line = re.sub(r"(against\s+invoice\s+)[a-z0-9./-]+", r"\1<docno>", line, flags=re.I)
    line = re.sub(r"(credit\s+for\s+invoice\s+)[a-z0-9./-]+", r"\1<docno>", line, flags=re.I)

    line = re.sub(r"\b(?:scr|si|csms|cn|inv|ord|sms)\s*[a-z0-9./-]+\b", "<docno>", line, flags=re.I)
    line = re.sub(r"\b(?:scr|si|csms|cn|inv|ord|sms)[a-z0-9./-]+\b", "<docno>", line, flags=re.I)

    line = re.sub(r"\bquater\b", "quarter", line)
    line = re.sub(r"\biitem\b", "item", line)
    line = re.sub(r"\bdetai1\b", "detail", line)
    line = re.sub(r"\b1tem\b", "item", line)

    line = re.sub(r"\bitem detail quantity unit line value\b", "line_items", line)
    line = re.sub(r"\bdetail quantity unit line value\b", "line_items", line)
    line = re.sub(r"\bdetail quantity line value\b", "line_items", line)
    line = re.sub(r"\bitem detail quantity line value\b", "line_items", line)

    line = re.sub(r"\baccounts payable\b", "accounts_payable", line)
    line = re.sub(r"\btotal ex vat\b", "total_ex_vat", line)
    line = re.sub(r"\binvoice number\b", "invoice_number", line)
    line = re.sub(r"\binvoice no\b", "invoice_number", line)
    line = re.sub(r"\bcredit number\b", "credit_number", line)
    line = re.sub(r"\bcredit no\b", "credit_number", line)

    line = re.sub(r"\s+", " ", line).strip()
    return line


def is_structural_line(line: str) -> bool:
    return any(re.search(pat, line, re.I) for pat in STRUCTURAL_PATTERNS)


def make_top_preview(lines, max_lines=12):
    structural = []
    seen = set()

    for line in lines[:120]:
        if not is_structural_line(line):
            continue
        norm = normalize_preview_line(line)
        if not norm or norm in seen:
            continue
        seen.add(norm)
        structural.append(norm)

    if structural:
        structural = sorted(structural)
        return " | ".join(structural[:max_lines])

    fallback = [normalize_preview_line(x) for x in lines[:8] if x.strip()]
    fallback = sorted(set(fallback))
    return " | ".join(fallback[:8])


def format_candidates(candidates):
    parts = []
    for c in candidates:
        label = f" {{{c['label']}}}" if "label" in c and c["label"] else ""
        parts.append(f"L{c['line_no']} [{c['source']}] {c['anchor_line']}{label} => {c['value']}")
    return " || ".join(parts)


def candidate_invoice_numbers(lines, max_hits=8):
    candidates = []
    for i, line in enumerate(lines):
        if not re.search(STRICT_INVOICE_LABEL_PATTERN, line, re.I):
            continue

        same_line = re.search(
            STRICT_INVOICE_LABEL_PATTERN + r"\s*[:.\-]?\s*(" + DOCNO_VALUE_PATTERN + r")",
            line,
            re.I,
        )
        next_line = lines[i + 1] if i + 1 < len(lines) else ""

        if same_line:
            value = same_line.group(1)
            source = "same_line"
        else:
            next_match = re.search(r"^" + DOCNO_VALUE_PATTERN + r"$", next_line.strip(), re.I)
            value = next_match.group(0) if next_match else ""
            source = "next_line" if value else "none"

        candidates.append({
            "line_no": i + 1,
            "source": source,
            "anchor_line": line,
            "value": value,
        })
        if len(candidates) >= max_hits:
            break
    return candidates


def candidate_credit_numbers(lines, max_hits=8):
    candidates = []
    for i, line in enumerate(lines):
        if not re.search(STRICT_CREDIT_LABEL_PATTERN, line, re.I):
            continue

        same_line = re.search(
            STRICT_CREDIT_LABEL_PATTERN + r"\s*[:.\-]?\s*(" + DOCNO_VALUE_PATTERN + r")",
            line,
            re.I,
        )
        next_line = lines[i + 1] if i + 1 < len(lines) else ""

        if same_line:
            value = same_line.group(1)
            source = "same_line"
        else:
            next_match = re.search(r"^" + DOCNO_VALUE_PATTERN + r"$", next_line.strip(), re.I)
            value = next_match.group(0) if next_match else ""
            source = "next_line" if value else "none"

        candidates.append({
            "line_no": i + 1,
            "source": source,
            "anchor_line": line,
            "value": value,
        })
        if len(candidates) >= max_hits:
            break
    return candidates


def candidate_credited_invoice_numbers(lines, max_hits=8):
    candidates = []
    patterns = [
        r"(?:credit\s+against\s+invoice)\s+(" + DOCNO_VALUE_PATTERN + r")",
        r"(?:credit\s+for\s+invoice)\s+(" + DOCNO_VALUE_PATTERN + r")",
        r"(?:credited\s+invoice)\s*[:.\-]?\s*(" + DOCNO_VALUE_PATTERN + r")",
        r"(?:original\s+invoice)\s*[:.\-]?\s*(" + DOCNO_VALUE_PATTERN + r")",
    ]

    for i, line in enumerate(lines):
        for pat in patterns:
            m = re.search(pat, line, re.I)
            if m:
                candidates.append({
                    "line_no": i + 1,
                    "source": "same_line",
                    "anchor_line": line,
                    "value": m.group(1).strip(),
                })
                break
        if len(candidates) >= max_hits:
            break
    return candidates


def candidate_dates(lines, max_hits=8):
    candidates = []
    for i, line in enumerate(lines):
        if any(re.search(pat, line, re.I) for pat in FALSE_DATE_PATTERNS):
            continue
        if not re.search(STRICT_DATE_LABEL_PATTERN, line, re.I):
            continue

        same_line = None
        for pat in DATE_VALUE_PATTERNS:
            same_line = re.search(pat, line)
            if same_line:
                break

        next_line = lines[i + 1] if i + 1 < len(lines) else ""
        next_match = None

        if not same_line and not any(re.search(pat, next_line, re.I) for pat in FALSE_DATE_PATTERNS):
            for pat in DATE_VALUE_PATTERNS:
                next_match = re.search(pat, next_line)
                if next_match:
                    break

        value = same_line.group(0) if same_line else (next_match.group(0) if next_match else "")
        source = "same_line" if same_line else ("next_line" if next_match else "none")

        # --- NEW ADDITION: Clean the dot out of the final string ---
        if value:
            # Transforms "21. July 2021" into "21 July 2021"
            value = re.sub(r"^(\d{1,2})\.\s+", r"\1 ", value.strip())
        # -----------------------------------------------------------

        candidates.append({
            "line_no": i + 1,
            "source": source,
            "anchor_line": line,
            "value": value,
        })
        if len(candidates) >= max_hits:
            break
    return candidates


def normalize_money_string(value: str) -> str:
    value = (value or "").strip()
    value = re.sub(r"[£$€\s]", "", value)

    if not value:
        return ""

    negative = False
    if value.startswith("-"):
        negative = True
        value = value[1:]

    # Remove commas (thousands separators)
    value = value.replace(",", "")

    # Clean up ANY leading zeros from the raw string digits now that '-' is removed
    if re.match(r"^0+\d", value):
        value = value.lstrip("0")
        if value.startswith("."):
            value = "0" + value

    if not re.fullmatch(r"\d+\.\d{2}", value):
        return ""

    return f"-{value}" if negative else value


def extract_money_values(line: str):
    raw_vals = re.findall(r"-?\s*[£$€]?\s*[\d,]+\.\d{2}", line or "")
    cleaned = []
    for raw in raw_vals:
        val = normalize_money_string(raw)
        if val:
            cleaned.append(val)
    return cleaned


def is_amount_reject_line(line: str) -> bool:
    txt = (line or "").strip().lower()
    return any(re.search(p, txt, re.I) for p in AMOUNT_REJECT_PATTERNS)


def candidate_amounts(lines, max_hits=20):
    candidates = []

    for i, line in enumerate(lines):
        txt = (line or "").strip()
        lower_txt = txt.lower()

        matched_label = None
        for label, pat in AMOUNT_PRIORITY_PATTERNS:
            if re.search(pat, lower_txt, re.I):
                matched_label = label
                break

        if matched_label is None:
            continue

        if is_amount_reject_line(txt):
            continue

        vals = extract_money_values(txt)
        value = vals[-1] if vals else ""
        source = "same_line" if value else "none"

        if not value and i + 1 < len(lines):
            next_line = (lines[i + 1] or "").strip()
            if not is_amount_reject_line(next_line):
                next_vals = extract_money_values(next_line)
                if next_vals:
                    value = next_vals[-1]
                    source = "next_line"

        candidates.append({
            "line_no": i + 1,
            "source": source,
            "anchor_line": line,
            "value": value,
            "label": matched_label,
        })
        if len(candidates) >= max_hits:
            break

    return candidates


def clean_detail_line(line: str) -> str:
    return re.sub(r"\s+", " ", (line or "")).strip()

def looks_like_detail_header(line: str) -> bool:
    txt = clean_detail_line(line).lower()
    # Changed from re.fullmatch to re.search to allow long patterns to match
    # even if other column text shares the same horizontal line.
    return any(re.search(p, txt, re.I) for p in DETAIL_HEADER_PATTERNS)

def looks_like_detail_stop(line: str) -> bool:
    txt = clean_detail_line(line).lower()
    return any(re.search(p, txt, re.I) for p in DETAIL_STOP_PATTERNS)


def candidate_detail_blocks(lines, max_hits=5):
    candidates = []

    for i, line in enumerate(lines):
        if not looks_like_detail_header(line):
            continue

        block = []
        j = i + 1

        while j < len(lines):
            nxt = clean_detail_line(lines[j])

            if not nxt:
                j += 1
                continue

            if looks_like_detail_stop(nxt):
                break

            block.append(nxt)
            j += 1

        block_text = "\n".join(block).strip()

        candidates.append({
            "line_no": i + 1,
            "source": "block",
            "anchor_line": line,
            "value": block_text,
        })

        if len(candidates) >= max_hits:
            break

    return candidates


def pick_best_invoice_number(candidates):
    for c in candidates:
        value = (c["value"] or "").strip()
        if value:
            return value
    return ""


def pick_best_credit_number(candidates):
    for c in candidates:
        value = (c["value"] or "").strip()
        if value:
            return value
    return ""


def pick_best_credited_invoice_number(candidates):
    for c in candidates:
        value = (c["value"] or "").strip()
        if value:
            return value
    return ""


def pick_best_date(candidates):
    # Ranked in order of legal/accounting preference
    preferred = [
        "tax point",
        "document date", 
        "credit date", 
        "invoice date", 
        "issued", 
        "date"
    ]

    for label in preferred:
        for c in candidates:
            if c["value"] and label in c["anchor_line"].lower():
                return c["value"]

    # Fallback to whatever date was found if none of the preferred labels match cleanly
    for c in candidates:
        if c["value"]:
            return c["value"]

    return ""


def pick_best_amount(candidates):
    preferred_order = [
        "total_incl_vat",
        "total_inc_vat",
        "total_including_vat",
        "grand_total",
        "invoice_total",
        "amount_due",
        "balance_due",
        "total",
    ]

    for wanted in preferred_order:
        # Reversing the candidates allows us to prioritize the LAST occurrence
        # of the required label down the page (typically the grand total).
        for c in reversed(candidates):
            value = (c.get("value") or "").strip()
            if c.get("label") == wanted and value:
                return value

    return ""


def pick_best_detail(candidates):
    best = ""

    for c in candidates:
        value = (c["value"] or "").strip()
        if not value:
            continue

        lines = [x for x in value.splitlines() if x.strip()]
        filtered = []

        for ln in lines:
            low = ln.strip().lower()
            if low in {"charges", "charge", "description", "detail", "item detail"}:
                continue
            
            # --- NEW ADDITION: Strip the "No." / "Item" column ---
            # Looks for an alphanumeric string at the start, followed by 3+ spaces.
            # Using lstrip() ensures we don't accidentally delete indented descriptions.
            ln = re.sub(r"^([A-Za-z0-9.-]+)\s{3,}", "", ln.lstrip())
            # -----------------------------------------------------
            
            filtered.append(ln.strip())

        value = "\n".join(filtered).strip()
        if not value:
            continue

        if len(value) > len(best):
            best = value

    return best


def detail_needs_ocr(detail_value: str) -> bool:
    txt = (detail_value or "").strip()
    if not txt:
        return True
    if txt.lower() in {"charges", "charge"}:
        return True
    if len(txt) < 25:
        return True
    return False


def compact_multiline(text: str) -> str:
    return " | ".join([x.strip() for x in (text or "").splitlines() if x.strip()])


def estimate_line_item_count(detail: str) -> int:
    lines = [x.strip() for x in (detail or "").splitlines() if x.strip()]
    useful = []
    for ln in lines:
        low = ln.lower()
        if low in {"charges", "detail", "description"}:
            continue
        useful.append(ln)
    return len(useful)


def detect_document_type(lines, credit_number, invoice_number):
    text = "\n".join(lines[:120]).lower()

    if credit_number:
        return "credit_note"
    # Added (note|memo) so both trigger the credit_note flag
    if re.search(r"\bcredit\s+(note|memo)\b", text):
        return "credit_note"
    if invoice_number:
        return "invoice"
    if re.search(r"\binvoice\b", text):
        return "invoice"
    return ""


def docno_shape(value: str) -> str:
    value = (value or "").strip()

    if not value:
        return "blank"
    if re.fullmatch(r"[A-Za-z]+[0-9]+", value):
        return "letters_digits"
    if re.fullmatch(r"[A-Za-z0-9]+", value):
        return "alnum_no_punct"
    if re.fullmatch(r"[A-Za-z]+[0-9]+(?:/[0-9.]+)+", value):
        return "letters_digits_slash"
    if "/" in value or "-" in value or "." in value:
        return "punctuated"
    return "other"


def date_shape(value: str) -> str:
    value = (value or "").strip()

    if not value:
        return "blank"
    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", value):
        return "dd/mm/yyyy"
    if re.fullmatch(r"\d{2}/\d{2}/\d{2}", value):  # <-- NEW RULE ADDED HERE
        return "dd/mm/yy"
    if re.fullmatch(r"\d{2}-\d{2}-\d{4}", value):
        return "dd-mm-yyyy"
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        return "yyyy-mm-dd"
    if re.fullmatch(r"\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}", value):
        return "d_month_yyyy"
    return "other"

def amount_shape(value: str) -> str:
    value = (value or "").strip()

    if not value:
        return "blank"
    if re.fullmatch(r"-?\d+\.\d{2}", value):
        return "money_2dp"
    if re.fullmatch(r"-?\d+", value):
        return "whole_number"
    return "other"


def detail_shape(value: str) -> str:
    value = (value or "").strip()

    if not value:
        return "blank"
    if re.match(r"(?i)^job number\b", value):
        return "job_number_detail"
    if "\n" in value:
        return "multiline"
    if len(value) <= 20:
        return "short_text"
    return "free_text"


def anchor_flags(lines):
    text = "\n".join(lines).lower()
    flags = []

    # Updated anchors with ^[\W_]*
    checks = [
        ("issuer_switch2", r"switch2 energy limited"),
        ("accounts_payable", r"accounts payable"),
        ("correspondence", r"all correspondence to be sent to"),
        ("invoice_label", STRICT_INVOICE_LABEL_PATTERN),
        ("credit_label", STRICT_CREDIT_LABEL_PATTERN),
        ("invoice_date_label", r"\binvoice date\b"),
        ("credit_date_label", r"\bcredit date\b"),
        ("line_items", r"(item detail quantity unit line value|detail quantity unit line value|detail quantity line value|no\.\s+description|^[\W_]*detail[\W_]*$|^[\W_]*description[\W_]*$|^[\W_]*charges[\W_]*$)"),
        ("total_label", r"^[\W_]*total\b"),
        ("grand_total", r"^[\W_]*grand\s+total\b"),
        ("amount_due", r"^[\W_]*amount\s+due\b"),
        ("balance_due", r"^[\W_]*balance\s+due\b"),
        ("invoice_total", r"^[\W_]*invoice\s+total\b"),
        ("total_ex_vat", r"^[\W_]*total\s+ex\s+vat\b"),
        ("credit_against_invoice", r"(credit against invoice|credit for invoice|credited invoice|original invoice)"),
    ]

    for name, pat in checks:
        if name in {"total_label", "grand_total", "amount_due", "balance_due", "invoice_total", "total_ex_vat"}:
            if any(re.search(pat, ln, re.I) for ln in lines):
                flags.append(name)
        else:
            if re.search(pat, text, re.I | re.M):
                flags.append(name)

    return sorted(flags)


def make_layout_signature(parsed, lines):
    flags = anchor_flags(lines)

    parts = [
        f"doc={parsed.get('document_type') or 'blank'}",
        f"invshape={docno_shape(parsed.get('invoice_number', ''))}",
        f"credshape={docno_shape(parsed.get('credit_number', ''))}",
        f"credinvshape={docno_shape(parsed.get('credited_invoice_number', ''))}",
        f"dateshape={date_shape(parsed.get('invoice_date', ''))}",
        f"amountshape={amount_shape(parsed.get('invoice_amount', ''))}",
        f"detailshape={detail_shape(parsed.get('detail', ''))}",
        "flags=" + ",".join(flags),
    ]

    return " | ".join(parts)


def empty_parsed():
    return {
        "top_preview": "",
        "document_type": "",
        "invoice_number": "",
        "credit_number": "",
        "credited_invoice_number": "",
        "invoice_date": "",
        "invoice_amount": "",
        "detail": "",
        "detail_compact": "",
        "line_item_count": 0,
        "invoice_candidates": [],
        "credit_candidates": [],
        "credited_invoice_candidates": [],
        "date_candidates": [],
        "amount_candidates": [],
        "detail_candidates": [],
    }


def parse_lines(lines):
    if not lines:
        return empty_parsed()

    top_preview = make_top_preview(lines)

    invoice_candidates = candidate_invoice_numbers(lines)
    credit_candidates = candidate_credit_numbers(lines)
    credited_invoice_candidates = candidate_credited_invoice_numbers(lines)
    date_candidates = candidate_dates(lines)
    amount_candidates = candidate_amounts(lines)
    detail_candidates = candidate_detail_blocks(lines)

    invoice_number = pick_best_invoice_number(invoice_candidates)
    credit_number = pick_best_credit_number(credit_candidates)
    credited_invoice_number = pick_best_credited_invoice_number(credited_invoice_candidates)
    invoice_date = pick_best_date(date_candidates)
    invoice_amount = pick_best_amount(amount_candidates)
    detail = pick_best_detail(detail_candidates)
    detail_compact = compact_multiline(detail)
    line_item_count = estimate_line_item_count(detail)
    document_type = detect_document_type(lines, credit_number, invoice_number)

    return {
        "top_preview": top_preview,
        "document_type": document_type,
        "invoice_number": invoice_number,
        "credit_number": credit_number,
        "credited_invoice_number": credited_invoice_number,
        "invoice_date": invoice_date,
        "invoice_amount": invoice_amount,
        "detail": detail,
        "detail_compact": detail_compact,
        "line_item_count": line_item_count,
        "invoice_candidates": invoice_candidates,
        "credit_candidates": credit_candidates,
        "credited_invoice_candidates": credited_invoice_candidates,
        "date_candidates": date_candidates,
        "amount_candidates": amount_candidates,
        "detail_candidates": detail_candidates,
    }


def has_useful_extracted_data(parsed):
    fields = [
        parsed.get("invoice_number", ""),
        parsed.get("credit_number", ""),
        parsed.get("credited_invoice_number", ""),
        parsed.get("invoice_date", ""),
        parsed.get("invoice_amount", ""),
        parsed.get("detail", ""),
    ]
    return any(bool(x) for x in fields)


def extraction_score(parsed):
    score = 0
    if parsed.get("invoice_number"):
        score += 3
    if parsed.get("credit_number"):
        score += 3
    if parsed.get("credited_invoice_number"):
        score += 2
    if parsed.get("invoice_date"):
        score += 2
    if parsed.get("invoice_amount"):
        score += 3
    detail = parsed.get("detail", "")
    if detail:
        score += min(5, max(1, len(detail) // 50))
    return score


def extract_and_parse(pdf_path: Path):
    notes = []

    text_pdf = extract_text_pdf(pdf_path)
    lines_pdf = get_lines(text_pdf)
    parsed_pdf = parse_lines(lines_pdf)

    use_ocr = False

    if len(text_pdf.strip()) < OCR_TEXT_THRESHOLD:
        use_ocr = True
        notes.append("Low text from pdfplumber; OCR fallback used")
    elif text_pdf.strip() and not has_useful_extracted_data(parsed_pdf):
        use_ocr = True
        notes.append("pdfplumber extracted text but no usable label data; OCR fallback used")
    elif text_pdf.strip() and detail_needs_ocr(parsed_pdf.get("detail", "")):
        use_ocr = True
        notes.append("pdfplumber detail was weak; OCR fallback used for detail recovery")

    if use_ocr:
        text_ocr = extract_text_ocr(pdf_path)
        lines_ocr = get_lines(text_ocr)
        parsed_ocr = parse_lines(lines_ocr)

        if text_ocr.strip():
            if not text_pdf.strip():
                return text_ocr, "ocr", notes, parsed_ocr, lines_ocr

            score_pdf = extraction_score(parsed_pdf)
            score_ocr = extraction_score(parsed_ocr)

            if score_ocr >= score_pdf:
                return text_ocr, "ocr", notes, parsed_ocr, lines_ocr
            else:
                # --- NEW GRAFTING LOGIC ---
                # PDF won the score, but if OCR found detail and PDF didn't, graft it in!
                if not parsed_pdf.get("detail") and parsed_ocr.get("detail"):
                    parsed_pdf["detail"] = parsed_ocr["detail"]
                    parsed_pdf["detail_compact"] = parsed_ocr["detail_compact"]
                    notes.append("PDF lacked detail; grafted detail block from OCR")
        else:
            notes.append("OCR produced no text")

    if text_pdf.strip():
        return text_pdf, "pdfplumber", notes, parsed_pdf, lines_pdf

    notes.append("No text extracted")
    return "", "pdfplumber", notes, empty_parsed(), []


def manual_review_required(parsed):
    doc_type = parsed.get("document_type", "")
    has_date = bool(parsed.get("invoice_date", ""))
    has_amount = bool(parsed.get("invoice_amount", ""))

    if doc_type == "credit_note":
        return not (parsed.get("credit_number", "") and has_date and has_amount)
    if doc_type == "invoice":
        return not (parsed.get("invoice_number", "") and has_date and has_amount)
    return True


def build_row(pdf: Path):
    text, method, notes, parsed, lines_used = extract_and_parse(pdf)
    layout_signature = make_layout_signature(parsed, lines_used)

    diagnostic_notes = list(notes)

    if not parsed["invoice_candidates"] and not parsed["credit_candidates"]:
        diagnostic_notes.append("No strict invoice/credit number label found")
    if not parsed["date_candidates"]:
        diagnostic_notes.append("No date label found")
    if not parsed["amount_candidates"]:
        diagnostic_notes.append("No amount label found")
    if parsed["amount_candidates"] and not parsed["invoice_amount"]:
        diagnostic_notes.append("Amount labels found but no final total extracted")
    if not parsed["detail_candidates"]:
        diagnostic_notes.append("No detail block found")
    if parsed["document_type"] == "credit_note" and not parsed["credit_number"]:
        diagnostic_notes.append("Credit note detected but credit number not extracted")
    if parsed["invoice_amount"] == "":
        diagnostic_notes.append("No final total extracted")
    if parsed["detail"] == "":
        diagnostic_notes.append("No detail extracted")

    return {
        "relative_path": str(pdf.relative_to(BASE_DIR)),
        "file_name": pdf.name,
        "extract_method": method,
        "text_length": len(text),
        "top_preview": parsed["top_preview"],
        "layout_signature": layout_signature,
        "document_type": parsed["document_type"],
        "invoice_number": parsed["invoice_number"],
        "credit_number": parsed["credit_number"],
        "credited_invoice_number": parsed["credited_invoice_number"],
        "invoice_date": parsed["invoice_date"],
        "invoice_amount": parsed["invoice_amount"],
        "detail": parsed["detail"],
        "detail_compact": parsed["detail_compact"],
        "line_item_count": parsed["line_item_count"],
        "manual_review_required": "YES" if manual_review_required(parsed) else "",
        "invoice_candidates": format_candidates(parsed["invoice_candidates"]),
        "credit_candidates": format_candidates(parsed["credit_candidates"]),
        "credited_invoice_candidates": format_candidates(parsed["credited_invoice_candidates"]),
        "date_candidates": format_candidates(parsed["date_candidates"]),
        "amount_candidates": format_candidates(parsed["amount_candidates"]),
        "detail_candidates": format_candidates(parsed["detail_candidates"]),
        "diagnostic_notes": "; ".join(diagnostic_notes),
    }


def assign_layout_labels(df: pd.DataFrame) -> pd.DataFrame:
    order = (
        df.groupby("layout_signature", dropna=False)
        .size()
        .reset_index(name="files")
        .sort_values(["files", "layout_signature"], ascending=[False, True])
        .reset_index(drop=True)
    )

    order["layout_label"] = [f"layout_{i+1}" for i in range(len(order))]
    return df.merge(order[["layout_signature", "layout_label"]], on="layout_signature", how="left")


def main():
    ensure_ocr_ready()
    pdfs = list_pdfs()
    if not pdfs:
        print("No PDFs found")
        return

    rows = []
    for pdf in pdfs:
        try:
            rows.append(build_row(pdf))
            print(f"Processed: {pdf}")
        except Exception as e:
            rows.append({
                "relative_path": str(pdf.relative_to(BASE_DIR)),
                "file_name": pdf.name,
                "extract_method": "ERROR",
                "text_length": 0,
                "top_preview": "",
                "layout_signature": "",
                "document_type": "",
                "invoice_number": "",
                "credit_number": "",
                "credited_invoice_number": "",
                "invoice_date": "",
                "invoice_amount": "",
                "detail": "",
                "detail_compact": "",
                "line_item_count": 0,
                "manual_review_required": "YES",
                "invoice_candidates": "",
                "credit_candidates": "",
                "credited_invoice_candidates": "",
                "date_candidates": "",
                "amount_candidates": "",
                "detail_candidates": "",
                "diagnostic_notes": f"Script error: {e}",
            })
            print(f"Failed: {pdf} -> {e}")

    df = pd.DataFrame(rows)
    df = assign_layout_labels(df)
    df = df.sort_values(["layout_label", "relative_path"])

    diag_columns = [
        "layout_label",
        "relative_path",
        "file_name",
        "extract_method",
        "text_length",
        "top_preview",
        "layout_signature",
        "document_type",
        "invoice_number",
        "credit_number",
        "credited_invoice_number",
        "invoice_date",
        "invoice_amount",
        "detail",
        "detail_compact",
        "line_item_count",
        "manual_review_required",
        "invoice_candidates",
        "credit_candidates",
        "credited_invoice_candidates",
        "date_candidates",
        "amount_candidates",
        "detail_candidates",
        "diagnostic_notes",
    ]

    df[diag_columns].to_csv(OUTPUT_DIAG, index=False, encoding="utf-8-sig")

    summary = (
        df.groupby("layout_label", dropna=False)
        .agg(
            files=("file_name", "count"),
            methods=("extract_method", lambda s: ", ".join(sorted(set(s)))),
            sample_file=("file_name", "first"),
            layout_signature=("layout_signature", "first"),
            top_preview=("top_preview", "first"),
            sample_document_type=("document_type", "first"),
            sample_invoice_number=("invoice_number", "first"),
            sample_credit_number=("credit_number", "first"),
            sample_credited_invoice_number=("credited_invoice_number", "first"),
            sample_invoice_date=("invoice_date", "first"),
            sample_invoice_amount=("invoice_amount", "first"),
            sample_detail_compact=("detail_compact", "first"),
            sample_line_item_count=("line_item_count", "first"),
            sample_manual_review_required=("manual_review_required", "first"),
            diagnostic_notes=("diagnostic_notes", "first"),
        )
        .reset_index()
        .sort_values(["files", "layout_label"], ascending=[False, True])
    )

    summary.to_csv(OUTPUT_LAYOUTS, index=False, encoding="utf-8-sig")

    print(f"\nWrote {OUTPUT_DIAG}")
    print(f"Wrote {OUTPUT_LAYOUTS}")
    print("\nTop layouts:")
    print(
        summary[
            [
                "layout_label",
                "files",
                "sample_file",
                "sample_document_type",
                "sample_invoice_amount",
                "sample_line_item_count",
                "sample_manual_review_required",
            ]
        ].head(20).to_string(index=False)
    )


if __name__ == "__main__":
    main()