import pandas as pd

# ✅ Step 1: Load Excel file
file_path = "GetQuoteApi.xlsx"
df = pd.read_excel(file_path)

# ✅ Step 2: Clean – Drop empty columns
df = df.dropna(axis=1, how='all')

# ✅ Step 3: Truncate long text
def truncate_text(text, length=40):
    if pd.isna(text):
        return ''
    text = str(text)
    return text if len(text) <= length else text[:length] + '...'

# ✅ Step 4: Build HTML rows with tooltip + modal click
def create_html_row(row):
    html = "<tr>"
    for col in df.columns:
        value = str(row[col]) if pd.notna(row[col]) else ''
        short = truncate_text(value)
        safe_value = value.replace("'", "\\'").replace('"', "&quot;").replace("\n", "<br>")
        tooltip = value.replace('"', "&quot;").replace("\n", " ")
        html += f"""
        <td>
            <span class="truncated" title="{tooltip}" onclick="showModal('{safe_value}')">{short}</span>
        </td>"""
    html += "</tr>"
    return html

# ✅ Step 5: Modal CSS and JavaScript
css_js = """
<style>
    table.quote-summary {
        border-collapse: collapse;
        width: 100%;
        table-layout: fixed;
        font-family: Arial, sans-serif;
        font-size: 13px;
    }
    table.quote-summary th, table.quote-summary td {
        border: 1px solid #ccc;
        padding: 8px;
        word-wrap: break-word;
        vertical-align: top;
    }
    table.quote-summary th {
        background-color: #2e74b5;
        color: white;
        font-weight: bold;
    }
    table.quote-summary tr:nth-child(even) {
        background-color: #f9f9f9;
    }
    .truncated {
        cursor: pointer;
        color: #006699;
        display: inline-block;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: 100%;
    }
    .modal {
        display: none;
        position: fixed;
        z-index: 999;
        padding-top: 60px;
        left: 0; top: 0;
        width: 100%; height: 100%;
        overflow: auto;
        background-color: rgba(0,0,0,0.5);
    }
    .modal-content {
        background-color: white;
        margin: auto;
        padding: 20px;
        border: 1px solid #aaa;
        width: 60%;
        max-height: 70%;
        overflow-y: auto;
        font-size: 15px;
        font-family: Arial;
    }
    .close {
        color: #aaa;
        float: right;
        font-size: 24px;
        font-weight: bold;
    }
    .close:hover, .close:focus {
        color: black;
        text-decoration: none;
        cursor: pointer;
    }
</style>

<script>
function showModal(content) {
    var modal = document.getElementById("myModal");
    var modalBody = document.getElementById("modal-body");
    modalBody.innerHTML = content;
    modal.style.display = "block";
}
window.onclick = function(event) {
    var modal = document.getElementById("myModal");
    if (event.target == modal) {
        modal.style.display = "none";
    }
}
</script>

<!-- Modal Box -->
<div id="myModal" class="modal">
  <div class="modal-content">
    <span class="close" onclick="document.getElementById('myModal').style.display='none'">&times;</span>
    <div id="modal-body"></div>
  </div>
</div>
"""

# ✅ Step 6: Build full table HTML
table_html = "<table class='quote-summary'><tr>" + "".join(f"<th>{col}</th>" for col in df.columns) + "</tr>"
for _, row in df.iterrows():
    table_html += create_html_row(row)
table_html += "</table>"

# ✅ Step 7: Final HTML structure
final_html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Get Quote API - Summary</title>
{css_js}
</head>
<body>
<h2 style='font-family: Arial;'>Get Quote API – Test Summary Report</h2>
{table_html}
</body>
</html>
"""

# ✅ Step 8: Save the HTML file
output_filename = "GetQuoteAPI_Summary_Report_v2.html"
with open(output_filename, "w", encoding="utf-8") as f:
    f.write(final_html)

print(f"✅ Report saved as: {output_filename}")
