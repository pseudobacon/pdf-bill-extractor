from html import escape
import pandas as pd
from excel_report import _grouped_ranges

def _safe_str(val) -> str:
    if val is None: return ""
    if isinstance(val, float) and pd.isna(val): return ""
    return str(val)

def generate_dashboard_html(df: pd.DataFrame, output_path, make_html_link):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    html_df = df.copy().reset_index(drop=True)

    if "file" in html_df.columns:
        html_df["file"] = html_df["file"].apply(lambda x: make_html_link(x) if x else "")

    for col in ["detail", "line_description"]:
        if col in html_df.columns:
            html_df[col] = html_df[col].apply(
                lambda x: escape(_safe_str(x)).replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")
            )

    columns = list(html_df.columns)
    thead_top = "".join(f"<th>{escape(col)}</th>" for col in columns)

    filter_cells = []
    for col in columns:
        col_lower = col.lower()
        # REMOVED: "layout" search target key
        if col_lower in {"document_type", "category", "in_scope"}:
            filter_cells.append(f'<th><select class="col-filter" data-col="{escape(col)}"><option value="">All</option></select></th>')
        elif col_lower == "invoice_date":
            filter_cells.append(f'<th><input type="text" class="col-filter date-filter" data-col="{escape(col)}" placeholder="dd/mm/yy"></th>')
        else:
            filter_cells.append(f'<th><input type="text" class="col-filter" data-col="{escape(col)}" placeholder="Filter {escape(col)}"></th>')

    ranges = _grouped_ranges(html_df)

    body_rows = []

    for start_idx, end_idx in ranges:
        for i in range(start_idx, end_idx + 1):
            row_data = html_df.loc[i]
            tds = []

            for col in columns:
                val = row_data.get(col, "")
                rendered = _safe_str(val)

                if (
                    i > start_idx
                    and col in {
                        "file",
                        "document_type",
                        "invoice_date",
                        "invoice_number",
                        "credit_number",
                        "invoice_amount",
                        "detail",
                    }
                ):
                    rendered = ""

                cell_content = (
                    rendered
                    if col == "file" or "<br>" in rendered
                    else escape(rendered)
                )

                tds.append(f"<td>{cell_content}</td>")

            body_rows.append("<tr>" + "".join(tds) + "</tr>")

    table_html = f"""
<table id="invoiceTable" class="display cell-border compact stripe order-column" style="width:100%">
    <thead>
        <tr>{thead_top}</tr>
        <tr>{''.join(filter_cells)}</tr>
    </thead>
    <tbody>
        {''.join(body_rows)}
    </tbody>
</table>
"""

    html_doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Invoice Extraction Dashboard</title>
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.8/css/jquery.dataTables.min.css">
<link rel="stylesheet" href="https://cdn.datatables.net/fixedheader/3.4.0/css/fixedHeader.dataTables.min.css">
<script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
<script src="https://cdn.datatables.net/1.13.8/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.datatables.net/fixedheader/3.4.0/js/dataTables.fixedHeader.min.js"></script>
<style>
body {{ font-family: Arial, sans-serif; margin: 20px; background: #f7f7f7; color: #222; }}
h1 {{ margin-bottom: 10px; }}
.note {{ margin-bottom: 14px; color: #555; font-size: 13px; }}
#invoiceTable {{ width: 100% !important; background: #fff; }}
table.dataTable thead th {{ white-space: nowrap; }}
table.dataTable tbody td {{ vertical-align: top; white-space: normal !important; word-break: break-word; }}
thead input, thead select {{ width: 100%; box-sizing: border-box; padding: 4px 6px; font-size: 12px; }}
.filter-row th {{ padding: 4px !important; }}
td:nth-child(7), td:nth-child(9) {{ max-width: 420px; }}
a {{ color: #1155cc; text-decoration: underline; }}
</style>
</head>
<body>
<h1>Invoice Extraction Dashboard</h1>
<p class="note">Interactive flat table view for filtering, sorting, and searching.</p>
{table_html}
<script>
$(document).ready(function() {{
    const table = $('#invoiceTable').DataTable({{
        orderCellsTop: true, fHeader: true, pageLength: 100, order: [], autoWidth: false
    }});
    $('#invoiceTable thead tr:eq(1) th').each(function(index) {{
        const headerCell = $(this), input = headerCell.find('input'), select = headerCell.find('select');
        if (select.length) {{
            const values = new Set();
            table.column(index).data().each(function(d) {{
                const text = $('<div>').html(d).text().trim();
                if (text) values.add(text);
            }});
            Array.from(values).sort().forEach(function(v) {{ select.append($('<option>', {{ value: v, text: v }})); }});
            select.on('change', function() {{
                const val = $.fn.dataTable.util.escapeRegex($(this).val());
                table.column(index).search(val ? '^' + val + '$' : '', true, false).draw();
            }});
        }}
        if (input.length) {{
            input.on('keyup change clear', function() {{
                if (table.column(index).search() !== this.value) {{ table.column(index).search(this.value).draw(); }}
            }});
        }}
    }});
}});
</script>
</body>
</html>
"""
    with open(output_path, "w", encoding="utf-8") as f: f.write(html_doc)
    print(f"HTML dashboard saved to {output_path}")