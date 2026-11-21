from html import escape
from report_common import render_exception_rows

def create_exceptions_report(dup_names, dup_invs, errs, missing, mismatches, output_path):
    duplicate_name_rows = [
        {
            "key": info["key"],
            "kept": info["kept"],
            "discarded": info["discarded"],
            "discard_count": len(info["discarded"]),
        }
        for _, info in sorted(dup_names.items(), key=lambda kv: kv[0].lower())
    ]

    duplicate_invoice_rows = [
        {
            "key": info["key"],
            "kept": info["kept"],
            "discarded": info["discarded"],
            "discard_count": len(info["discarded"]),
            "all_files": [info["kept"], *info["discarded"]],
        }
        for _, info in sorted(dup_invs.items(), key=lambda kv: kv[0].lower())
    ]

    # Calculate metrics for the Summary Box Card
    total_dup_names = sum(len(x["discarded"]) for x in duplicate_name_rows)
    total_dup_invoices = sum(len(x["discarded"]) for x in duplicate_invoice_rows)
    total_errors = len(errs)
    total_missing = len(missing)
    total_mismatches = len(mismatches)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Extraction Exceptions</title>
<style>
body {{
    font-family: Arial, sans-serif;
    margin: 20px;
    background: #f7f7f7;
    color: #222;
}}
.section {{
    background: #fff;
    padding: 16px;
    margin-bottom: 25px;
    border-radius: 4px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}}
h2 {{
    margin-top: 0;
    border-bottom: 2px solid #eee;
    padding-bottom: 6px;
    color: #333;
}}
table {{
    width: 100%;
    border-collapse: collapse;
    margin-top: 10px;
}}
th, td {{
    text-align: left;
    padding: 10px 12px;
    border-bottom: 1px solid #ddd;
    font-size: 13px;
    vertical-align: top;
    line-height: 1.5;
}}
th {{
    background: #f0f0f0;
    font-weight: bold;
    color: #444;
}}
.empty {{
    color: #777;
    font-style: italic;
}}
.summary-container {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 15px;
    margin-top: 12px;
}}
.summary-card {{
    background: #fcfcfc;
    border: 1px solid #e6e6e6;
    padding: 12px;
    border-radius: 4px;
}}
.summary-label {{
    font-size: 12px;
    color: #666;
    text-transform: uppercase;
    font-weight: bold;
}}
.summary-value {{
    font-size: 22px;
    font-weight: bold;
    color: #111;
    margin-top: 4px;
}}
</style>
</head>
<body>
<h1>Extraction Exceptions Tracker Report</h1>

<div class="section">
    <h2>Summary Metrics</h2>
    <div class="summary-container">
        <div class="summary-card">
            <div class="summary-label">Duplicate Filenames</div>
            <div class="summary-value">{total_dup_names}</div>
        </div>
        <div class="summary-card">
            <div class="summary-label">Duplicate Invoices</div>
            <div class="summary-value">{total_dup_invoices}</div>
        </div>
        <div class="summary-card">
            <div class="summary-label">System Errors</div>
            <div class="summary-value">{total_errors}</div>
        </div>
        <div class="summary-card">
            <div class="summary-label">Missing Core Data</div>
            <div class="summary-value">{total_missing}</div>
        </div>
        <div class="summary-card">
            <div class="summary-label">Amount Mismatches</div>
            <div class="summary-value">{total_mismatches}</div>
        </div>
    </div>
</div>

<div class='section'><h2>Duplicate Filenames</h2>
"""
    # IMPROVEMENT: Discarded files are escaped individually and separated cleanly by <br> tags
    html += render_exception_rows(
        duplicate_name_rows,
        ["Filename Key", "Kept Master File Location", "Discarded Duplicate Instances"],
        lambda x: (
            f"<tr>"
            f"<td>{escape(x['key'])}</td>"
            f"<td>{escape(x['kept'])}</td>"
            f"<td>{'<br>'.join(escape(d) for d in x['discarded'])}</td>"
            f"</tr>"
        ),
    )
    html += "</div>"

    html += "<div class='section'><h2>Duplicate Invoice Numbers & Amounts</h2>"
    # IMPROVEMENT: Matched files stacked vertically row-by-row on a new line
    html += render_exception_rows(
        duplicate_invoice_rows,
        ["Invoice Key Identity", "Retained Copy", "All Discovered Matches Across Folders"],
        lambda x: (
            f"<tr>"
            f"<td>{escape(x['key'])}</td>"
            f"<td>{escape(x['kept'])}</td>"
            f"<td>{'<br>'.join(escape(f) for f in x['all_files'])}</td>"
            f"</tr>"
        ),
    )
    html += "</div>"

    html += "<div class='section'><h2>Errors</h2>"
    html += render_exception_rows(
        errs,
        ["File", "Error"],
        lambda x: (f"<tr><td>{escape(x['pdf'])}</td><td>{escape(x['error'])}</td></tr>"),
    )
    html += "</div>"

    html += "<div class='section'><h2>Missing Data</h2>"
    html += render_exception_rows(
        missing,
        ["File", "Missing"],
        lambda x: (f"<tr><td>{escape(x['pdf'])}</td><td>{escape(', '.join(x['missing']))}</td></tr>"),
    )
    html += "</div>"

    html += "<div class='section'><h2>Mismatches</h2>"
    html += render_exception_rows(
        mismatches,
        ["File", "Invoice Amt", "Line Sum", "VAT Sum", "Calc Total", "Diff"],
        lambda x: (
            f"<tr>"
            f"<td>{escape(x['pdf'])}</td>"
            f"<td>{escape(x['invoice_amount'])}</td>"
            f"<td>{escape(x['line_amount_sum'])}</td>"
            f"<td>{escape(x['vat_amount_sum'])}</td>"
            f"<td>{escape(x['calculated_total'])}</td>"
            f"<td>{escape(x['difference'])}</td>"
            f"</tr>"
        ),
    )
    html += "</div>"
    html += "</body>\n</html>"

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Exceptions report sheet written out to {output_path}")