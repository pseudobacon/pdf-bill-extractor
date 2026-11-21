#!/usr/bin/env python3

import re
import pdfplumber
from pathlib import Path

# --- CHANGE THIS TO YOUR SPECIFIC TEST FILE ---
TEST_PDF = "515828.pdf" 
# ----------------------------------------------

def extract_raw_text(pdf_path):
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            if txt.strip():
                parts.append(txt)
    return "\n".join(parts)

def debug_amount_extraction():
    target_file = Path(__file__).resolve().parent / TEST_PDF
    
    if not target_file.exists():
        print(f"ERROR: Could not find {target_file}")
        return

    print(f"--- EXTRACTING: {target_file.name} ---")
    raw_text = extract_raw_text(target_file)
    lines = [x.strip() for x in raw_text.splitlines() if x.strip()]

    print("\n--- LAST 15 LINES OF EXTRACTED TEXT (WHAT THE MACHINE SEES) ---")
    for i, line in enumerate(lines[-15:]):
        print(f"Line {len(lines) - 15 + i}: '{line}'")

    print("\n--- HUNTING FOR 'TOTAL' KEYWORDS ---")
    found_any = False
    for i, line in enumerate(lines):
        if "total" in line.lower() or "amount" in line.lower() or "due" in line.lower() or "vat" in line.lower():
            found_any = True
            print(f"\nFound keyword on Line {i}: '{line}'")
            
            # Show what is on the very next line in case the amount dropped down
            if i + 1 < len(lines):
                print(f"  -> Next line reads: '{lines[i+1]}'")
            
            # Test if our money regex actually sees a number here
            money_matches = re.findall(r"-?\s*[£$€]?\s*[\d,]+\.\d{2}", line)
            if money_matches:
                print(f"  -> Money regex found on SAME line: {money_matches}")
            elif i + 1 < len(lines):
                next_money = re.findall(r"-?\s*[£$€]?\s*[\d,]+\.\d{2}", lines[i+1])
                if next_money:
                    print(f"  -> Money regex found on NEXT line: {next_money}")
                else:
                    print("  -> NO money format detected on this or next line.")

    if not found_any:
        print("\nWARNING: The words 'total', 'amount', 'due', or 'vat' were not found anywhere in the text.")

if __name__ == "__main__":
    debug_amount_extraction()