#!/usr/bin/env python3

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.comments import Comment

from excel_report import write_excel_report

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "Output"

OLD_REPORT_PATH = OUTPUT_DIR / "invoice_extraction_old.xlsx"
FRESH_REPORT_PATH = OUTPUT_DIR / "invoice_extraction.xlsx"
MERGED_REPORT_PATH = OUTPUT_DIR / "invoice_extraction_merged.xlsx"

MANUAL_COLS = [
    "Manual complete",
    "Comment",
    "in_scope",
]


def normalize(value):
    """
    Make matching resilient to formatting differences.

    Examples:
        15      -> 15.00
        15.0    -> 15.00
        15.00   -> 15.00
        None    -> ""
        Text    -> lowercase stripped
    """
    if pd.isna(value):
        return ""

    if isinstance(value, (int, float, np.integer, np.floating)):
        return f"{float(value):.2f}"

    text = str(value).strip()

    try:
        return f"{float(text):.2f}"
    except Exception:
        return text.lower()


def build_key(row, cols):
    return "|".join(normalize(row.get(col)) for col in cols)


def extract_metadata_from_excel(filepath):
    comments = {}

    wb = openpyxl.load_workbook(filepath, data_only=False)
    ws = wb.active

    headers = {
        cell.column: cell.value
        for cell in ws[1]
        if cell.value
    }

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.comment:
                col_name = headers.get(cell.column)

                if col_name:
                    comments[(cell.row, col_name)] = {
                        "text": cell.comment.text,
                        "author": cell.comment.author,
                    }

    return comments


def apply_metadata_to_excel(filepath, df, old_comments):
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    headers = {
        cell.value: cell.column
        for cell in ws[1]
        if cell.value
    }

    for idx, row_data in df.iterrows():

        old_row_num = row_data.get("_old_row")

        if pd.isna(old_row_num):
            continue

        new_row_num = idx + 2

        for col_name, col_idx in headers.items():

            comment_data = old_comments.get(
                (int(old_row_num), col_name)
            )

            if comment_data:
                ws.cell(
                    row=new_row_num,
                    column=col_idx
                ).comment = Comment(
                    comment_data["text"],
                    comment_data["author"] or "Author",
                )

    anchor_col = 1
    merge_start_col = 1
    merge_end_col = 7

    row = 2
    last_row = ws.max_row

    merge_groups = []

    while row <= last_row:

        value = ws.cell(row=row, column=anchor_col).value

        is_blank = (
            value is None
            or str(value).strip() == ""
        )

        if is_blank:
            row += 1
            continue

        start_row = row
        end_row = row

        next_row = row + 1

        while next_row <= last_row:

            next_value = ws.cell(
                row=next_row,
                column=anchor_col
            ).value

            next_blank = (
                next_value is None
                or str(next_value).strip() == ""
            )

            if next_blank:
                end_row = next_row
                next_row += 1
            else:
                break

        if end_row > start_row:
            merge_groups.append((start_row, end_row))

        row = next_row

    for start_row, end_row in merge_groups:

        for col_idx in range(
            merge_start_col,
            merge_end_col + 1
        ):
            ws.merge_cells(
                start_row=start_row,
                start_column=col_idx,
                end_row=end_row,
                end_column=col_idx,
            )

    wb.save(filepath)


def main():

    if not OLD_REPORT_PATH.exists():
        print(f"Missing: {OLD_REPORT_PATH}")
        return

    if not FRESH_REPORT_PATH.exists():
        print(f"Missing: {FRESH_REPORT_PATH}")
        return

    print("Loading workbooks...")

    old_df = pd.read_excel(OLD_REPORT_PATH)
    fresh_df = pd.read_excel(FRESH_REPORT_PATH)

    old_comments = extract_metadata_from_excel(
        OLD_REPORT_PATH
    )

    old_cols = list(old_df.columns)
    fresh_cols = list(fresh_df.columns)

    old_df["_old_row"] = old_df.index + 2

    #
    # Header-level invoice identity
    #
    header_key_cols = [
        c
        for c in [
            "document_type",
            "invoice_date",
            "invoice_number",
            "invoice_amount",
        ]
        if c in old_df.columns
        and c in fresh_df.columns
    ]

    #
    # Line-level identity
    #
    preferred_line_cols = [
        "line_description",
        "line_amount",
        "vat_amount",
        "line_total",
    ]

    line_key_cols = [
        c
        for c in preferred_line_cols
        if c in old_df.columns
        and c in fresh_df.columns
    ]

    if not header_key_cols:
        raise ValueError(
            "No header key columns found."
        )

    if not line_key_cols:
        raise ValueError(
            "No line key columns found."
        )

    manual_cols = [
        c
        for c in MANUAL_COLS
        if c in old_df.columns
    ]

    print(
        "Header keys:",
        header_key_cols
    )

    print(
        "Line keys:",
        line_key_cols
    )

    #
    # Fill merged invoice headers down
    #
    old_df[header_key_cols] = (
        old_df[header_key_cols].ffill()
    )

    fresh_df[header_key_cols] = (
        fresh_df[header_key_cols].ffill()
    )

    #
    # Build match keys
    #
    key_cols = (
        header_key_cols +
        line_key_cols
    )

    old_df["_match_key"] = old_df.apply(
        lambda r: build_key(r, key_cols),
        axis=1,
    )

    fresh_df["_match_key"] = fresh_df.apply(
        lambda r: build_key(r, key_cols),
        axis=1,
    )

    #
    # Handle duplicate identical lines
    #
    old_df["_match_occurrence"] = (
        old_df
        .groupby("_match_key")
        .cumcount()
    )

    fresh_df["_match_occurrence"] = (
        fresh_df
        .groupby("_match_key")
        .cumcount()
    )

    #
    # Diagnostics
    #
    old_keys = set(old_df["_match_key"])
    fresh_keys = set(fresh_df["_match_key"])

    matched_keys = len(
        old_keys & fresh_keys
    )

    missing_keys = len(
        fresh_keys - old_keys
    )

    print(
        f"Matched keys: {matched_keys}"
    )

    print(
        f"Fresh-only keys: {missing_keys}"
    )

    #
    # Remove manual columns from fresh file
    #
    fresh_base_df = fresh_df.drop(
        columns=manual_cols,
        errors="ignore",
    )

    mapping_df = old_df[
        [
            "_match_key",
            "_match_occurrence",
            "_old_row",
            *manual_cols,
        ]
    ].copy()

    print(
        "Copying manual values from old workbook..."
    )

    merged_df = fresh_base_df.merge(
        mapping_df,
        on=[
            "_match_key",
            "_match_occurrence",
        ],
        how="left",
        validate="m:1",
    )

    matched_rows = (
        merged_df["_old_row"]
        .notna()
        .sum()
    )

    print(
        f"Matched rows: {matched_rows} "
        f"of {len(merged_df)}"
    )

    #
    # Debug manual value transfer
    #
    print("\n" + "=" * 80)
    print("MANUAL VALUE DIAGNOSTICS")
    print("=" * 80)

    old_manual_count = (
        old_df["Manual complete"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("Y")
        .sum()
    )

    merged_manual_count = (
        merged_df["Manual complete"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("Y")
        .sum()
    )

    print(f"Old Manual=Y    : {old_manual_count}")
    print(f"Merged Manual=Y : {merged_manual_count}")
    print(f"Difference      : {old_manual_count - merged_manual_count}")

    unmatched_rows = merged_df[
        merged_df["_old_row"].isna()
    ].copy()

    print(f"\nRows with no match at all: {len(unmatched_rows)}")

    unmatched_rows.to_excel(
        OUTPUT_DIR / "debug_unmatched_rows.xlsx",
        index=False
    )

    print(
        f"Saved {OUTPUT_DIR / 'debug_unmatched_rows.xlsx'}"
    )

    #
    # Find Manual=Y rows from old file that did not transfer
    #
    old_manual = old_df[
        old_df["Manual complete"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("Y")
    ].copy()

    old_manual["_debug_key"] = (
        old_manual["_match_key"]
        + "|"
        + old_manual["_match_occurrence"].astype(str)
    )

    merged_manual = merged_df[
        merged_df["Manual complete"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("Y")
    ].copy()

    merged_manual_keys = set(
        merged_manual["_match_key"]
        + "|"
        + merged_manual["_match_occurrence"].astype(str)
    )

    missing_manual = old_manual[
        ~old_manual["_debug_key"].isin(merged_manual_keys)
    ].copy()

    print(f"\nLost Manual=Y rows: {len(missing_manual)}")

    if len(missing_manual):
        missing_manual.to_excel(
            OUTPUT_DIR / "debug_missing_manual_rows.xlsx",
            index=False
        )

        print(
            f"Saved {OUTPUT_DIR / 'debug_missing_manual_rows.xlsx'}"
        )

        cols = [
            c for c in [
                "invoice_number",
                "credit_number",
                "line_description",
                "line_amount",
                "vat_amount",
                "line_total"
            ]
            if c in missing_manual.columns
        ]

        print("\nFirst 20 missing rows:")
        print(
            missing_manual[cols]
            .head(20)
            .to_string()
        )

    #
    # Preserve old visible column order
    #
    helper_cols = {
        "_old_row",
        "_match_key",
        "_match_occurrence",
    }

    final_cols = []

    for col in old_cols:
        if (
            col in merged_df.columns
            and col not in helper_cols
            and col not in final_cols
        ):
            final_cols.append(col)

    for col in fresh_cols:
        if (
            col in merged_df.columns
            and col not in helper_cols
            and col not in final_cols
        ):
            final_cols.append(col)

    final_export_df = merged_df[
        final_cols
    ].copy()

    print(
        f"Saving: {MERGED_REPORT_PATH}"
    )

    write_excel_report(
        final_export_df,
        MERGED_REPORT_PATH,
    )

    print(
        "Reapplying comments and merges..."
    )

    apply_metadata_to_excel(
        MERGED_REPORT_PATH,
        merged_df,
        old_comments,
    )

    print("Done.")


if __name__ == "__main__":
    main()