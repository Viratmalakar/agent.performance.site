from flask import Flask, request, send_file, render_template_string
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

HTML = """
<html>
<head>
<title>Agent Performance</title>
<style>
body{font-family:Arial;background:#f2f2f2;padding:30px;}
.box{background:white;padding:25px;border-radius:12px;max-width:1200px;margin:auto;}
table{border-collapse:collapse;width:100%;}
th,td{border:1px solid #ccc;padding:6px;text-align:center;font-size:13px;}
th{background:#ffe082;}
button{padding:8px 16px;background:#2196f3;color:white;border:none;border-radius:5px;cursor:pointer;}
</style>
</head>
<body>
<div class="box">
<h2>Upload 4 Excel Reports</h2>
<form method="post" action="/process" enctype="multipart/form-data">
<input type="file" name="files" multiple required><br><br>
<button>Process</button>
</form>
<br>
{{table}}
{{download}}
</div>
</body>
</html>
"""

@app.route("/")
def home():
    return render_template_string(HTML, table="", download="")

@app.route("/process", methods=["POST"])
def process():

    files = request.files.getlist("files")
    paths = []

    for f in files:
        path = os.path.join(UPLOAD_FOLDER, f.filename)
        f.save(path)
        paths.append(path)

    # Load files
    login = pd.read_excel(paths[0], header=2)
    cdr = pd.read_excel(paths[1], header=1)
    agent = pd.read_excel(paths[2], header=2)
    crm = pd.read_excel(paths[3])

    # ---------------- LOGIN ----------------
    # Auto detect date column
    date_col = None
    for c in login.columns:
        if "date" in c.lower() or "time" in c.lower():
            date_col = c
            break

    if not date_col:
        return "Date column not found in Login report"

    login[date_col] = pd.to_datetime(login[date_col], errors="coerce")
    first_login = login.groupby("UserName")[date_col].min().dt.time

    # ---------------- AGENT PERFORMANCE ----------------
    agent["Total Break"] = agent["LUNCHBREAK"] + agent["SHORTBREAK"] + agent["TEABREAK"]
    agent["Total Meeting"] = agent["MEETING"] + agent["SYSTEMDOWN"]
    agent["Total Net Login"] = agent["Total Login Time"] - agent["Total Break"]

    # ---------------- CDR ----------------
    cdr["Mature"] = cdr["Disposition"].isin(["CALLMATURED", "TRANSFER"])

    total_mature = cdr.groupby("Username")["Mature"].sum()
    transfer_call = cdr[cdr["Disposition"]=="TRANSFER"].groupby("Username").size()
    ib_mature = cdr[
        (cdr["Disposition"].isin(["CALLMATURED","TRANSFER"])) &
        (cdr["Campaign"]=="CSRINBOUND")
    ].groupby("Username").size()

    # ---------------- CRM ----------------
    total_tagging = crm.groupby("CreatedByID").size()

    # ---------------- FINAL MERGE ----------------
    result = pd.DataFrame()

    result["Employee ID"] = agent["Agent Name"]
    result["Agent Name"] = agent["Agent Name"]
    result["First Login Time"] = result["Employee ID"].map(first_login)

    result["Total Login"] = agent["Total Login Time"]
    result["Total Net Login"] = agent["Total Net Login"]
    result["Total Break"] = agent["Total Break"]
    result["Total Meeting"] = agent["Total Meeting"]
    result["Total Talk Time"] = agent["Total Talk Time"]

    result["Total Mature"] = result["Employee ID"].map(total_mature).fillna(0)
    result["IB Mature"] = result["Employee ID"].map(ib_mature).fillna(0)
    result["Transfer Call"] = result["Employee ID"].map(transfer_call).fillna(0)
    result["OB Mature"] = result["Total Mature"] - result["IB Mature"]

    result["Total Tagging"] = result["Employee ID"].map(total_tagging).fillna(0)

    # AHT
    result["AHT"] = result["Total Talk Time"] / result["Total Mature"].replace(0,1)

    # Save Excel
    output_file = "Final_Report.xlsx"
    result.to_excel(output_file, index=False)

    table_html = result.to_html(index=False)

    return render_template_string(
        HTML,
        table=table_html,
        download='<br><a href="/download"><button>Download Excel</button></a>'
    )

@app.route("/download")
def download():
    return send_file("Final_Report.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
