#!/usr/bin/env python3

import pandas as pd
from pathlib import Path

# =============================================================================
# CONFIGURATION
# =============================================================================
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "Output"

OLD_REPORT_PATH = OUTPUT_DIR / "invoice_extraction_old.xlsx"
FRESH_REPORT_PATH = OUTPUT_DIR / "invoice_extraction.xlsx"
DIAGNOSTIC_PATH = OUTPUT_DIR / "merge_diagnostics.xlsx"

def main():
    if not OLD_REPORT_PATH.exists() or not FRESH_REPORT_PATH.exists():
        print("ERROR: Could not find one or both report files.")
        return

    print("Loading data for comparison...")
    old_df = pd.read_excel(OLD_REPORT_PATH)
    fresh_df = pd.read_excel(FRESH_REPORT_PATH)

    # 1. Define the exact same columns used in the merge script
    fresh_cols = list(fresh_df.columns)
    EXCLUDED_COLS = ["detail", "line_description", "category", "in_scope"]
    merge_keys = [c for c in fresh_cols if c not in EXCLUDED_COLS]
    
    print(f"Comparing rows based on these columns:\n{merge_keys}\n")

    # 2. Build the safe keys
    def build_safe_key(row):
        vals = [str(row.get(c, "")).replace('nan', '').strip().lower() for c in merge_keys]
        return "|".join(vals)

    old_df['_merge_key'] = old_df.apply(build_safe_key, axis=1)
    fresh_df['_merge_key'] = fresh_df.apply(build_safe_key, axis=1)

    old_df['_occurrence'] = old_df.groupby('_merge_key').cumcount()
    fresh_df['_occurrence'] = fresh_df.groupby('_merge_key').cumcount()

    # Create a single unique string for easy comparison
    old_df['_final_key'] = old_df['_merge_key'] + "|Occur:" + old_df['_occurrence'].astype(str)
    fresh_df['_final_key'] = fresh_df['_merge_key'] + "|Occur:" + fresh_df['_occurrence'].astype(str)

    # 3. Find the mismatches
    old_keys = set(old_df['_final_key'])
    fresh_keys = set(fresh_df['_final_key'])

    missing_in_fresh = old_keys - fresh_keys
    missing_in_old = fresh_keys - old_keys

    print(f"Total Rows in Old: {len(old_df)}")
    print(f"Total Rows in Fresh: {len(fresh_df)}")
    print(f"--> Matches Found: {len(old_keys.intersection(fresh_keys))}")
    print(f"--> Old Rows that FAILED to match: {len(missing_in_fresh)}")
    print(f"--> Fresh Rows with no old match: {len(missing_in_old)} (Expected for new invoices)")

    # 4. Extract the broken rows
    old_orphans = old_df[old_df['_final_key'].isin(missing_in_fresh)].copy()
    fresh_unmatched = fresh_df[fresh_df['_final_key'].isin(missing_in_old)].copy()

    # Sort them by file and invoice number to make manual review easier
    if 'file' in old_orphans.columns:
        old_orphans = old_orphans.sort_values(['file', 'invoice_number'])
    if 'file' in fresh_unmatched.columns:
        fresh_unmatched = fresh_unmatched.sort_values(['file', 'invoice_number'])

    # 5. Export to Excel
    print(f"\nExporting detailed breakdown to {DIAGNOSTIC_PATH}...")
    with pd.ExcelWriter(DIAGNOSTIC_PATH, engine='openpyxl') as writer:
        old_orphans.to_excel(writer, sheet_name="Lost Old Data", index=False)
        fresh_unmatched.to_excel(writer, sheet_name="New or Changed Fresh", index=False)
        
    print("Done! Open 'merge_diagnostics.xlsx' to see exactly which rows are misbehaving.")

if __name__ == "__main__":
    main()