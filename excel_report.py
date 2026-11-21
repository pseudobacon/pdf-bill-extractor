import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

def _safe_str(val) -> str:
    if val is None: return ""
    if isinstance(val, float) and pd.isna(val): return ""
    return str(val)

def _group_keys(df: pd.DataFrame):
    preferred = ["file", "invoice_number", "credit_number", "invoice_amount"]
    return [c for c in preferred if c in df.columns]

def _grouped_ranges(df: pd.DataFrame):
    if df.empty: return []
    keys = _group_keys(df)
    if not keys: return [(0, len(df) - 1)]

    rows = df.reset_index(drop=True)
    ranges = []
    start_idx = 0

    for i in range(1, len(rows) + 1):
        is_break = (
            i == len(rows)
            or any(_safe_str(rows.loc[i, k]) != _safe_str(rows.loc[i - 1, k]) for k in keys)
        )
        if is_break:
            ranges.append((start_idx, i - 1))
            start_idx = i
    return ranges

def write_excel_report(df: pd.DataFrame, output_path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    export_df = df.copy().reset_index(drop=True)

    wb = Workbook()
    ws = wb.active

    columns = list(export_df.columns)
    ws.append(columns)

    for row in export_df.itertuples(index=False):
        ws.append([_safe_str(val) for val in row])

    # REMOVED: "layout" text from the vertical rowspan merge array configuration
    cols_to_merge = [
        "file", "document_type", "invoice_date", "invoice_number",
        "invoice_amount", "detail", "credit_number"
    ]
    merge_col_indexes = [columns.index(col) + 1 for col in cols_to_merge if col in columns]

    ranges = _grouped_ranges(export_df)
    for start_idx, end_idx in ranges:
        if start_idx == end_idx: continue
        start_excel_row = start_idx + 2
        end_excel_row = end_idx + 2

        for col_idx in merge_col_indexes:
            ws.merge_cells(start_row=start_excel_row, start_column=col_idx, end_row=end_excel_row, end_column=col_idx)
            cell = ws.cell(row=start_excel_row, column=col_idx)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for row_num in range(2, ws.max_row + 1):
        for col_idx in range(1, len(columns) + 1):
            cell = ws.cell(row=row_num, column=col_idx)
            if col_idx in merge_col_indexes:
                if row_num == start_excel_row:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
            else:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    # Auto-fit Width Engine
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_length = 0
        for cell in col_cells:
            if cell.value:
                lines = str(cell.value).split('\n')
                longest_line = max((len(line) for line in lines), default=0)
                if longest_line > max_length: max_length = longest_line
        ws.column_dimensions[col_letter].width = min(max_length + 2.5, 50)

    # Dynamic Row Heights
    for row_idx, row_cells in enumerate(ws.iter_rows(), start=1):
        max_lines = 1
        for cell in row_cells:
            if cell.column in merge_col_indexes: continue
            if cell.value:
                col_letter = get_column_letter(cell.column)
                col_width = ws.column_dimensions[col_letter].width
                text_str = str(cell.value)
                total_lines = text_str.count('\n') + int(len(text_str) / max(col_width - 2, 1)) + 1
                if total_lines > max_lines: max_lines = total_lines
        if max_lines > 1: ws.row_dimensions[row_idx].height = max_lines * 15

    wb.save(output_path)
    print(f"Excel report saved to {output_path}")