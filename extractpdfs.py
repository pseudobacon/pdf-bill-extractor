#!/usr/bin/env python3

import os
import re
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed

import pandas as pd
import pdfplumber
import pytesseract
from pdf2image import convert_from_path

# =============================================================================
# 1. CONFIGURATION & CONSTANTS
# =============================================================================

class Config:
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_FILE = BASE_DIR / "final_invoice_extraction.csv"
    OUTPUT_HTML = BASE_DIR / "final_invoice_extraction.html"
    EXCEPTIONS_HTML = BASE_DIR / "extraction_exceptions.html"
    DEBUG_DIR = BASE_DIR / "_debug"

    OCR_TEXT_THRESHOLD = 80
    MAX_WORKERS = max(1, min((os.cpu_count() or 4) - 1, 4))
    TESSERACT_CONFIG = r"--oem 3 --psm 6"


class RegexPatterns:
    # Date Patterns
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

    # Document Number Patterns
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

    # Details & Line Item Patterns
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

    # Amounts
    AMOUNT_PRIORITIES = [
        ("total_incl_vat", r"\btotal\s*(?:[£$€])?\s*(?:incl\.?|inc\.?)\s*vat\b"),
        ("total_inc_vat", r"\btotal\s*(?:[£$€])?\s*inc\.?\s*vat\b"),
        ("total_including_vat", r"\btotal\s*(?:[£$€])?\s*including\s*vat\b"),
        ("grand_total", r"\bgrand\s+total\b"), ("amount_due", r"\bamount\s+due\b"),
        ("balance_due", r"\bbalance\s+due\b"), ("invoice_total", r"\binvoice\s+total\b"), ("total", r"\btotal\b"),
    ]
    AMOUNT_REJECTS = [
        r"\btotal\s+ex\s+vat\b", r"\btotal\s*(?:[£$€])?\s*excl\.?\s*vat\b", r"\btotal\s*(?:[£$€])?\s*excluding\s*vat\b",
        r"^[\W_]*vat\b", r"^[\W_]*sub\s*total\b", r"^[\W_]*subtotal\b", r"^[\W_]*net\b",
        r"^[\W_]*unit\b", r"^[\W_]*rate\b",
    ]

    # Line Item Specific Patterns
    LINE_ITEM_SKIP = [
        r"^Engineer Comments\b", r"^Attended\b", r"^Job Number\b", r"^Initial Call Type\b",
        r"^Caller Name\b", r"^Time Details\b", r"^Time Charges\b", r"^Travel Between\b",
        r"^Work Between\b", r"^Rounded\b", r"^Rounded upto\b", r"^Minutes\b",
        r"^Split due to charge rate boundary\b", r"^Please remit payments\b", r"^Page\s+\d+\s+of\s+\d+\b",
        r"^For finance queries\b", r"^For any urgent repairs\b", r"^Fax\b", r"^Email\b", r"^Customer Services\b",
    ]
    LINE_ITEM_SUMMARY = [
        r"\bcost of this work\b", r"\btotal cost of this work\b", r"\bthis work will be\b",
        r"\btotal cost\b", r"\btotal\b.*\bexc\b.*\bvat\b", r"\btotal\b.*\bex\b.*\bvat\b",
        r"\bexc\b\.?\s*vat\b", r"\bex\b\.?\s*vat\b",
    ]
    LINE_ITEM_FALLBACK_AMOUNT = r"(?<![\d.])-?\s*[Â£$â‚¬]?\s*[\d,]+\.\d{2}(?=\s|$|[)\],;:])"

    # Layout Flag Anchors
    ANCHOR_CHECKS = [
        ("issuer_switch2", r"switch2 energy limited"),
        ("accounts_payable", r"accounts payable"),
        ("invoice_label", INVOICE_STRICT_LABEL),
        ("credit_label", CREDIT_STRICT_LABEL),
        ("line_items", r"(item detail quantity unit line value|detail quantity unit line value|detail quantity line value|no\.\s+description|^[\W_]*detail[\W_]*$|^[\W_]*description[\W_]*$|^[\W_]*charges[\W_]*$)"),
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

    # Deduplicate and clean
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

        raw_vals = re.findall(r"-?\s*[Â£$â‚¬]?\s*[\d,]+\.\d{2}", txt)
        value = re.sub(r"[Â£$â‚¬\s,]", "", raw_vals[-1]) if raw_vals else ""

        if not value and i + 1 < len(lines):
            next_line = (lines[i + 1] or "").strip()
            if not any(re.search(p, next_line, re.I) for p in RegexPatterns.AMOUNT_REJECTS):
                next_vals = re.findall(r"-?\s*[Â£$â‚¬]?\s*[\d,]+\.\d{2}", next_line)
                if next_vals: value = re.sub(r"[Â£$â‚¬\s,]", "", next_vals[-1])

        if value:
            candidates.append({"label": matched_label, "value": value})
    return candidates

def candidate_detail_blocks(lines):
    candidates = []
    i = 0

    while i < len(lines):
        line = (lines[i] or "").strip()
        clean_line = re.sub(r"\s+", " ", line)

        header_match = next(
            (p for p in RegexPatterns.DETAIL_HEADERS
             if re.search(p, clean_line.lower(), re.I)),
            None
        )

        if header_match:
            block = []
            j = i + 1

            stop_pattern = None
            stop_line = None

            while j < len(lines):
                nxt = re.sub(r"\s+", " ", (lines[j] or "")).strip()

                stop_pattern = next(
                    (p for p in RegexPatterns.DETAIL_STOPS
                     if nxt and re.search(p, nxt.lower(), re.I)),
                    None
                )

                if stop_pattern:
                    stop_line = nxt
                    break

                if nxt:
                    block.append(nxt)

                j += 1

            if block:
                candidates.append({
                    "value": "\n".join(block).strip()
                })

                # DEBUG
                print("\n" + "=" * 80)
                print(f"DETAIL BLOCK #{len(candidates)}")
                print("=" * 80)
                print(f"HEADER MATCH : {header_match}")
                print(f"START LINE   : {i + 1}")

                if stop_pattern:
                    print(f"STOP PATTERN : {stop_pattern}")
                    print(f"STOP LINE    : {stop_line}")
                    print(f"STOP INDEX   : {j + 1}")
                else:
                    print("STOP PATTERN : <EOF>")

                print(f"LINES CAPTURED: {len(block)}")

                print("\nFIRST 5 LINES:")
                for ln in block[:5]:
                    print(f"  {ln}")

                print("\nLAST 5 LINES:")
                for ln in block[-5:]:
                    print(f"  {ln}")

            i = j
        else:
            i += 1

    return candidates

def pick_best_detail(candidates):
    # Instead of picking just ONE block, we COMBINE all candidates. 
    # This handles multi-page invoices seamlessly.
    combined_filtered = []
    
    for c in candidates:
        if not (value := (c.get("value") or "").strip()): continue

        for ln in [x for x in value.splitlines() if x.strip()]:
            low = ln.strip().lower()
            
            # Skip standalone header words
            if low in {"charges", "charge", "description", "detail", "item detail", "price"}: 
                continue
                
            # Strip out header patterns that repeat on new pages (e.g., "Item Detail Quantity Unit...")
            if any(re.search(p, low, re.I) for p in RegexPatterns.DETAIL_HEADERS):
                continue
                
            if re.fullmatch(r"\d{2}/\d{2}/\d{2,4}", low): 
                continue
                
            if re.search(r"^(qty|quantity|unit|unit price|price|value|vat rate|vat identifier|discount %|receipt date|measure)$", low, re.I): 
                continue
                
            # Clean up leading formatting artifacts
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


# =============================================================================
# 4. CORE ORCHESTRATION & LINE ITEMS
# =============================================================================

def extract_line_items(detail_text: str):
    items, matched_time_charge_descs = [], set()
    lines = [line.strip() for line in (detail_text or "").splitlines() if line.strip()]

    # FIX: Allow '-' to pass through the cleaning process
    clean_amt = lambda val: re.sub(r"[£$€\s,]", "", (val or "").strip())
    norm_space = lambda s: re.sub(r"\s+", " ", (s or "").strip())

# --- Helper to validate if a line is actually a full charge row ---
    def is_valid_charge(desc, amt_str):
        # Rule: A valid charge must have a description length > 3 
        # and not look like a fragment (e.g., just "DOMESTIC COMMISSIONING SALES")
        if len(desc) < 5: return False
        
        # Rule: If it's one of these "fragment" invoices, it must contain a quantity/multiplier
        if "DOMESTIC COMMISSIONING" in desc and not any(char.isdigit() for char in desc):
            return False
        return True

    for line in lines:
        if any(re.search(p, line, re.I) for p in RegexPatterns.LINE_ITEM_SKIP + RegexPatterns.LINE_ITEM_SUMMARY):
            continue

        matched = False

    for line in lines:
        if any(re.search(p, line, re.I) for p in RegexPatterns.LINE_ITEM_SKIP + RegexPatterns.LINE_ITEM_SUMMARY):
            continue

        matched = False

        # Pattern 1: Time Charges (Support Negative)
        # Added -? before the currency/amount
        if m := re.match(r"^(Time charge(?: using .*?)?\s*\(.*?\)).*?(?<![\d.])-?[£$€]?\s*(-?[\d,]+\.\d{2})\s*$", line, re.I):
            desc = m.group(1).strip()
            matched_time_charge_descs.add(norm_space(desc).lower())
            items.append({"line_description": desc, "line_amount": clean_amt(m.group(2))})
            matched = True
            continue

        if norm_space(line).lower() in matched_time_charge_descs:
            continue

        # Pattern 2: Standard Items (Support Negative)
        # Regex updated to look for -? at the start of amount groups
        if m := re.match(r"^(.*?\D)\s+(\d+(?:\.\d+)?)\s+[£$€]?\s*(-?[\d,]+\.\d{2})\s+[£$€]?\s*(-?[\d,]+\.\d{2})(?:\s+\d+(?:\.\d+)?\s*%?)?\s*$", line, re.I):
            desc = f"{m.group(1).strip()} {m.group(2).strip()}".strip()
            items.append({"line_description": desc, "line_amount": clean_amt(m.group(4))})
            matched = True
            continue

        # Pattern 3: Dates mixed with values (Support Negative)
        if m := re.search(r"^(.*?\S)\s+\d+(?:\.\d+)?\s+\d{2}/\d{2}/\d{2,4}\s+[A-Za-z]+\s+(-?\d+(?:\.\d{2})?)\s+[A-Za-z0-9]+\s+(-?\d+(?:\.\d{2})?)\s*$", line, re.I):
            desc = m.group(1).strip()
            items.append({"line_description": desc, "line_amount": clean_amt(m.group(3))})
            matched = True
            continue

# Fallback Pattern (Updated with Validation)
        fallback_match = re.search(r"(?<![\d.])-?\s*[£$€]?\s*(-?[\d,]+\.\d{2})(?:\s*(?:[A-Za-z]{1,3}\d?|\d+(?:\.\d+)?\s*%))?\s*$", line)
        if fallback_match and "=" not in line:
            amt = clean_amt(fallback_match.group(1))
            desc = line[:fallback_match.start()].strip()
            desc = re.sub(r"[\s=:\-]+$", "", desc).strip()
            
            # Apply our new validation gate
            if is_valid_charge(desc, amt):
                items.append({"line_description": desc, "line_amount": amt})
                matched = True
                continue

        # Continuation Line Catch-All
        if not matched and items and len(line) > 2:
            if not re.search(r"^(Invoice|Credit)\s+No", line, re.I):
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

def anchor_flags(lines):
    text = "\n".join(lines).lower()
    flags = []
    for name, pat in RegexPatterns.ANCHOR_CHECKS:
        if name == "line_items":
            if any(re.search(pat, ln, re.I) for ln in lines): flags.append(name)
        elif re.search(pat, text, re.I | re.M):
            flags.append(name)
    return sorted(flags)


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
        layout_sig = f"doc={doc_type} | flags={','.join(anchor_flags(parsed.get('lines_ref', [])))}"

        base_row = {
            "layout_signature": layout_sig, "file": rel_path, "document_type": doc_type,
            "invoice_date": parsed.get("invoice_date", ""), "invoice_number": parsed.get("invoice_number", ""),
            "credit_number": parsed.get("credit_number", ""), "invoice_amount": parsed.get("invoice_amount", ""),
            "detail": parsed.get("detail", ""),
        }

        missing_fields = [k for k in ["invoice_number", "invoice_date", "invoice_amount", "detail"] if not base_row[k]]
        line_items = extract_line_items(base_row["detail"])
        rows = []

        if not line_items:
            rows.append({**base_row, "line_description": "", "line_amount": "", "vat_amount": "", "line_total": ""})
            missing_fields.extend(["line_description", "line_amount"])
        else:
            for item in line_items:
                row = {**base_row, "line_description": item["line_description"], "line_amount": item["line_amount"], "vat_amount": "", "line_total": ""}
                try:
                    amt_float = float(item["line_amount"])
                    row.update({"line_amount": f"{amt_float:.2f}", "vat_amount": f"{amt_float * 0.20:.2f}", "line_total": f"{amt_float * 1.20:.2f}"})
                except ValueError: pass
                rows.append(row)

        return {"pdf": rel_path, "rows": rows, "error": None, "missing_fields": missing_fields}
    except Exception as e:
        return {"pdf": rel_path, "rows": [], "error": str(e), "missing_fields": []}

# =============================================================================
# 6. HTML DASHBOARDS AND REPORT GENERATION
# =============================================================================

def build_total_mismatch_files(df: pd.DataFrame):
    mismatches = []
    to_float = lambda val: float(str(val).replace(",", "").strip()) if str(val).replace(",", "").strip() else None

    for file_name, grp in df.groupby("file", dropna=False):
        invoice_total = next((to_float(v) for v in grp["invoice_amount"] if to_float(v) is not None), None)
        if invoice_total is None: continue

        line_sum = sum(to_float(v) or 0.0 for v in grp["line_amount"])
        vat_sum = sum(to_float(v) or 0.0 for v in grp["vat_amount"])
        calc_total = round(line_sum + vat_sum, 2)

        if abs(invoice_total - calc_total) > 0.05:
            mismatches.append({"pdf": file_name, "invoice_amount": f"{invoice_total:.2f}", "line_amount_sum": f"{line_sum:.2f}", "vat_amount_sum": f"{vat_sum:.2f}", "calculated_total": f"{calc_total:.2f}", "difference": f"{round(invoice_total - calc_total, 2):.2f}"})
    return mismatches

def generate_dashboard_html(df: pd.DataFrame):
    html_df = df.copy()
    html_df["file"] = html_df["file"].apply(lambda x: f'<a href="{x}" target="_blank">{x}</a>' if x else "")
    html_df["detail"] = html_df["detail"].apply(lambda x: x.replace("\n", "<br>") if isinstance(x, str) else x)
    html_df["line_description"] = html_df["line_description"].apply(lambda x: x.replace("\n", "<br>") if isinstance(x, str) else x)
    
    table_html = html_df.to_html(index=False, escape=False, classes="display cell-border", table_id="invoiceTable")
    
    with open(Config.OUTPUT_HTML, "w", encoding="utf-8") as f:
        # [Preserved your exact JavaScript DataTables logic here]
        f.write(f"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><title>Invoice Extraction Dashboard</title><link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css"><style>body{{font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding:20px; background-color:#f8f9fa;}} .container{{background-color:white; padding:20px; border-radius:8px; box-shadow:0 2px 4px rgba(0,0,0,0.1); overflow-x:auto;}} h2{{color:#333;}} a{{color:#0066cc; text-decoration:none; font-weight:500;}} a:hover{{text-decoration:underline;}} td{{vertical-align:top; line-height:1.4;}} thead input, thead select{{width:100%; box-sizing:border-box; padding:4px; border:1px solid #ccc; border-radius:4px; font-weight:normal;}} .filters th{{padding:8px 4px !important; background-color:#f1f1f1;}} table.dataTable.display tbody tr.odd > .sorting_1, table.dataTable.display tbody tr.even > .sorting_1, table.dataTable.display tbody tr.odd, table.dataTable.display tbody tr.even{{background-color:transparent !important;}} table.dataTable tbody tr.file-group-0 td, table.dataTable.display tbody tr.file-group-0 td.sorting_1{{background-color:#ffffff !important;}} table.dataTable tbody tr.file-group-1 td, table.dataTable.display tbody tr.file-group-1 td.sorting_1{{background-color:#f2f7fb !important;}}</style></head><body><div class="container"><h2>Invoice Extraction Dashboard</h2>{table_html}</div><script src="https://code.jquery.com/jquery-3.7.0.min.js"></script><script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script><script>$(document).ready(function(){{ $('#invoiceTable thead tr').clone(true).addClass('filters').appendTo('#invoiceTable thead'); $.fn.dataTable.ext.search.push(function(settings, data, dataIndex) {{ var dateSearch = $('#dateSearchInput').val(); if (!dateSearch) return true; var rowDateRaw = data[3] || ""; var rowDateIso = ""; var ukMatch = rowDateRaw.match(/^(\\d{{1,2}})[\\/\\.\\-](\\d{{1,2}})[\\/\\.\\-](\\d{{2,4}})/); if (ukMatch) {{ var d = ukMatch[1].padStart(2, '0'); var m = ukMatch[2].padStart(2, '0'); var y = ukMatch[3].length === 2 ? "20" + ukMatch[3] : ukMatch[3]; rowDateIso = y + "-" + m + "-" + d; }} else {{ var dObj = new Date(rowDateRaw); if (!isNaN(dObj)) {{ rowDateIso = dObj.getFullYear() + "-" + String(dObj.getMonth() + 1).padStart(2, '0') + "-" + String(dObj.getDate()).padStart(2, '0'); }} }} return rowDateIso === dateSearch; }}); var table = $('#invoiceTable').DataTable({{ "pageLength": 50, "lengthChange": false, "orderCellsTop": true, "order": [[ 0, "asc" ]], "drawCallback": function (settings) {{ var api = this.api(); var lastFile = null; var groupToggle = 0; api.rows({{page:'current'}}).every(function(rowIdx) {{ var rowData = api.row(rowIdx).data(); var fileData = rowData[1] ? rowData[1].replace(/(<([^>]+)>)/gi, "").trim() : ""; if (fileData !== lastFile) {{ groupToggle = 1 - groupToggle; lastFile = fileData; }} $(this.node()).removeClass('file-group-0 file-group-1').addClass('file-group-' + groupToggle); }}); function applySmartMerge(colIndex) {{ var lastData = null; var lastCell = null; var lastFile = null; api.rows({{page:'current'}}).every(function(rowIdx) {{ var rowData = api.row(rowIdx).data(); var cell = api.cell(rowIdx, colIndex).node(); var fileData = rowData[1] ? rowData[1].replace(/(<([^>]+)>)/gi, "").trim() : ""; var rawCell = rowData[colIndex] || ""; var cellData = String(rawCell).trim(); if (cellData === lastData && fileData === lastFile && cellData !== "") {{ $(cell).empty().css('border-top', 'none'); if (lastCell) $(lastCell).css('border-bottom', 'none'); }} else {{ $(cell).html(rawCell).css('border-top', '').css('border-bottom', ''); lastData = cellData; lastFile = fileData; lastCell = cell; }} }}); }} for (var i = 0; i <= 7; i++) {{ applySmartMerge(i); }} }}, initComplete: function () {{ var api = this.api(); api.columns().eq(0).each(function (colIdx) {{ var cell = $('.filters th').eq($(api.column(colIdx).header()).index()); var title = $(cell).text(); if (title === 'invoice_date') {{ $(cell).html('<input type="date" id="dateSearchInput" />'); $('#dateSearchInput').on('change clear', function() {{ table.draw(); }}); }} else if (title === 'layout' || title === 'document_type') {{ var select = $('<select><option value="">All</option></select>').appendTo($(cell).empty()).on('change', function () {{ var val = $.fn.dataTable.util.escapeRegex($(this).val()); api.column(colIdx).search(val ? '^' + val + '$' : '', true, false).draw(); }}); api.column(colIdx).data().unique().sort().each(function (d) {{ var cleanData = d.replace(/(<([^>]+)>)/gi, "").trim(); if (cleanData) select.append('<option value="' + cleanData + '">' + cleanData + '</option>'); }}); }} else {{ $(cell).html('<input type="text" placeholder="Filter ' + title + '" />'); $('input', cell).on('keyup change clear', function () {{ if (api.column(colIdx).search() !== this.value) {{ api.column(colIdx).search(this.value).draw(); }} }}); }} }}); }} }}); }}); </script></body></html>""")
    print(f"Interactive dashboard saved to {Config.OUTPUT_HTML}")

def create_exceptions_report(dup_names, dup_invs, errs, missing, mismatches):
    html = f"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><title>Extraction Exceptions</title><style>body{{font-family:Arial,sans-serif; margin:20px; background:#f7f7f7;}} .section{{background:#fff; padding:16px; margin-bottom:32px; border-radius:8px;}} table{{width:100%; border-collapse:collapse; margin-top:16px;}} th,td{{border:1px solid #ddd; padding:8px; text-align:left;}} th{{background:#f0f0f0;}} .summary-card{{background:#fafafa; border:1px solid #ddd; padding:16px; display:inline-block; margin:5px;}}</style></head><body><h1>Extraction Exceptions Report</h1>"""
    
    html += f"<div class='section'><h2>Summary</h2><div class='summary-card'>Duplicates (Files): {len(dup_names)}</div><div class='summary-card'>Duplicates (Invoices): {len(dup_invs)}</div><div class='summary-card'>Errors: {len(errs)}</div><div class='summary-card'>Missing Data: {len(missing)}</div><div class='summary-card'>Mismatches: {len(mismatches)}</div></div>"
    
    html += "<div class='section'><h2>Errors</h2><table><tr><th>File</th><th>Error</th></tr>" + "".join(f"<tr><td>{x['pdf']}</td><td>{x['error']}</td></tr>" for x in errs) + "</table></div>"
    html += "<div class='section'><h2>Missing Data</h2><table><tr><th>File</th><th>Missing</th></tr>" + "".join(f"<tr><td>{x['pdf']}</td><td>{', '.join(x['missing'])}</td></tr>" for x in missing) + "</table></div>"
    html += "<div class='section'><h2>Mismatches</h2><table><tr><th>File</th><th>Invoice Amt</th><th>Calc Total</th><th>Diff</th></tr>" + "".join(f"<tr><td>{x['pdf']}</td><td>{x['invoice_amount']}</td><td>{x['calculated_total']}</td><td>{x['difference']}</td></tr>" for x in mismatches) + "</table></div></body></html>"
    
    with open(Config.EXCEPTIONS_HTML, "w", encoding="utf-8") as f: f.write(html)
    print(f"Exceptions report saved to {Config.EXCEPTIONS_HTML}")


# =============================================================================
# 7. MAIN ENTRY POINT
# =============================================================================

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

def main():
    pdfs = sorted(Config.BASE_DIR.rglob("*.pdf"))
    if not pdfs:
        print("No PDFs found"); return

    unique_pdfs, duplicate_filename_groups = deduplicate_inputs(pdfs)
    duplicate_invoice_groups, seen_invoice_keys = {}, {}
    error_files, missing_data_files, rows = [], [], []

    with ProcessPoolExecutor(max_workers=min(Config.MAX_WORKERS, len(unique_pdfs))) as executor:
        futures = {executor.submit(process_pdf, str(pdf)): pdf for pdf in unique_pdfs}

        for future in as_completed(futures):
            result = future.result()
            if result["error"]:
                print(f"Failed: {result['pdf']} -> {result['error']}")
                error_files.append({"pdf": result["pdf"], "error": result["error"]})
                continue

            if not result.get("rows"):
                print(f"Processed: {result['pdf']} (no rows)")
                continue

            first_row = result["rows"][0]
            invoice_key = ((first_row.get("document_type") or "").strip().lower(), (first_row.get("invoice_number") or "").strip().upper(), (first_row.get("invoice_amount") or "").strip())

            if invoice_key[1] and invoice_key[2]:
                if invoice_key in seen_invoice_keys:
                    group_key = f"{invoice_key[0]} | {invoice_key[1]} | {invoice_key[2]}"
                    if group_key not in duplicate_invoice_groups:
                        duplicate_invoice_groups[group_key] = {"key": group_key, "kept": seen_invoice_keys[invoice_key], "discarded": []}
                    duplicate_invoice_groups[group_key]["discarded"].append(result["pdf"])
                    print(f"Skipped duplicate invoice: {result['pdf']}")
                    continue
                seen_invoice_keys[invoice_key] = result["pdf"]

            rows.extend(result["rows"])
            if result.get("missing_fields"):
                missing_data_files.append({"pdf": result["pdf"], "missing": result["missing_fields"]})
            print(f"Processed: {result['pdf']}")

    if rows:
        df = pd.DataFrame(rows)
        # Assign Layouts
        order = df.groupby("layout_signature", dropna=False).size().reset_index(name="files").sort_values(["files", "layout_signature"], ascending=[False, True]).reset_index(drop=True)
        order["layout"] = [str(i + 1) for i in range(len(order))]
        df = df.merge(order[["layout_signature", "layout"]], on="layout_signature", how="left")

        # Sort and Format Final CSV
        final_cols = ["layout", "file", "document_type", "invoice_date", "invoice_number", "credit_number", "invoice_amount", "detail", "line_description", "line_amount", "vat_amount", "line_total"]
        df = df[final_cols].sort_values(["layout", "file"])
        
        csv_df = df.copy()
        csv_df.loc[csv_df.duplicated(subset=["file"]), ["layout", "file", "document_type", "invoice_date", "invoice_number", "credit_number", "invoice_amount", "detail"]] = ""
        csv_df.to_csv(Config.OUTPUT_FILE, index=False, encoding="utf-8-sig")
        print(f"\nExtraction complete! CSV saved to {Config.OUTPUT_FILE}")

        generate_dashboard_html(df)
        create_exceptions_report(duplicate_filename_groups, duplicate_invoice_groups, error_files, missing_data_files, build_total_mismatch_files(df))
    else:
        print("No rows extracted")
        create_exceptions_report(duplicate_filename_groups, duplicate_invoice_groups, error_files, missing_data_files, [])

if __name__ == "__main__":
    main()