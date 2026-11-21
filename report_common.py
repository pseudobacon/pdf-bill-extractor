from html import escape
from pathlib import Path
import pandas as pd


def safe_str(val) -> str:
    if val is None:
        return ""
    if isinstance(val, float) and pd.isna(val):
        return ""
    return str(val)


def group_keys_for_merge(df: pd.DataFrame):
    preferred = ["file", "invoice_number", "credit_number", "invoice_amount"]
    return [c for c in preferred if c in df.columns]


def grouped_ranges(df: pd.DataFrame):
    group_keys = group_keys_for_merge(df)
    if not group_keys or df.empty:
        return []

    ranges = []
    rows = df.reset_index(drop=True)
    start_idx = 0

    for i in range(1, len(rows) + 1):
        is_break = (
            i == len(rows)
            or any(safe_str(rows.loc[i, k]) != safe_str(rows.loc[i - 1, k]) for k in group_keys)
        )
        if is_break:
            ranges.append((start_idx, i - 1))
            start_idx = i

    return ranges


def build_total_mismatch_files(df: pd.DataFrame):
    mismatches = []
    to_float = lambda val: float(str(val).replace(",", "").strip()) if str(val).replace(",", "").strip() else None

    for file_name, grp in df.groupby("file", dropna=False):
        invoice_total = next((to_float(v) for v in grp["invoice_amount"] if to_float(v) is not None), None)
        if invoice_total is None:
            continue

        line_sum = sum(to_float(v) or 0.0 for v in grp["line_amount"])
        vat_sum = sum(to_float(v) or 0.0 for v in grp["vat_amount"])
        calc_total = round(line_sum + vat_sum, 2)

        if abs(invoice_total - calc_total) > 0.05:
            mismatches.append({
                "pdf": file_name,
                "invoice_amount": f"{invoice_total:.2f}",
                "line_amount_sum": f"{line_sum:.2f}",
                "vat_amount_sum": f"{vat_sum:.2f}",
                "calculated_total": f"{calc_total:.2f}",
                "difference": f"{round(invoice_total - calc_total, 2):.2f}",
            })

    return mismatches


def is_copy_filename(path_str: str) -> bool:
    name = Path(path_str).name.lower()
    return "copy of" in name


def duplicate_keep_priority(path_str: str):
    name = Path(path_str).name.lower()
    parent_depth = len(Path(path_str).parts)
    return (
        1 if "copy of" in name else 0,
        parent_depth,
        len(name),
        name,
    )


def choose_preferred_invoice_file(existing_path: str, candidate_path: str) -> str:
    ranked = sorted([existing_path, candidate_path], key=duplicate_keep_priority)
    return ranked[0]


def render_exception_rows(rows, headers, row_builder):
    if not rows:
        return "<p>None.</p>"
    head = "".join(f"<th>{escape(str(h))}</th>" for h in headers)
    body = "".join(row_builder(r) for r in rows)
    return f"<table><tr>{head}</tr>{body}</table>"
