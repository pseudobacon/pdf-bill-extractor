#!/usr/bin/env python3

import re
import pdfplumber
from pathlib import Path

# --- CHANGE THIS TO YOUR SPECIFIC TEST FILE ---
TEST_PDF = "736152.pdf" 
# ----------------------------------------------

def extract_raw_text(pdf_path):
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # We use layout=True to preserve horizontal spacing better, 
            # which is often where tabular dates fail to parse correctly.
            txt = page.extract_text(layout=True) or ""
            if txt.strip():
                parts.append(txt)
    return "\n".join(parts)

def debug_date_extraction():
    target_file = Path(__file__).resolve().parent / TEST_PDF
    
    if not target_file.exists():
        print(f"ERROR: Could not find {target_file}")
        return

    print(f"--- EXTRACTING: {target_file.name} ---")
    raw_text = extract_raw_text(target_file)
    lines = [x.strip() for x in raw_text.splitlines() if x.strip()]

    print("\n--- FIRST 20 LINES OF EXTRACTED TEXT (WHAT THE MACHINE SEES) ---")
    # Dates are usually near the top, so we print the first 20 lines instead of the last
    for i, line in enumerate(lines[:20]):
        print(f"Line {i}: '{line}'")

    print("\n--- HUNTING FOR 'DATE' KEYWORDS ---")
    found_any = False
    
    # Regex to catch DD/MM/YY, DD/MM/YYYY, MM.DD.YY, DD-MM-YYYY, etc.
    date_regex = r"\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}"
    
    for i, line in enumerate(lines):
        if "date" in line.lower():
            found_any = True
            print(f"\nFound keyword on Line {i}: '{line}'")
            
            # Show what is on the very next line in case the date dropped down
            if i + 1 < len(lines):
                print(f"  -> Next line reads: '{lines[i+1]}'")
            
            # Test if our date regex actually sees a date here
            date_matches = re.findall(date_regex, line)
            if date_matches:
                print(f"  -> Date regex found on SAME line: {date_matches}")
            elif i + 1 < len(lines):
                next_date = re.findall(date_regex, lines[i+1])
                if next_date:
                    print(f"  -> Date regex found on NEXT line: {next_date}")
                else:
                    print("  -> NO date format detected on this or next line.")

    if not found_any:
        print("\nWARNING: The word 'date' was not found anywhere in the text.")

if __name__ == "__main__":
    debug_date_extraction()