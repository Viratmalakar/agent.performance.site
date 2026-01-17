from flask import Flask, request, send_file, render_template_string
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

HTML = """
<html>
<head>
<title>Agent Performance</title>
<style>
body{font-family:Arial;background:linear-gradient(135deg,#ff9acb,#a9b9ff);padding:40px;}
.box{background:white;padding:30px;border-radius:15px;max-width:1200px;margin:auto;}
table{border-collapse:collapse;width:100%;}
th,td{border:1px solid #ccc;padding:8px;text-align:center;}
th{background:#ffd27d;}
button{padding:10px 20px;background:#2196f3;color:white;border:none;border-radius:6px;cursor:pointer;}
</style>
</head>
<body>
<div class="box">
<h2>Upload Reports</h2>
<form method="post" action="/process" enctype="multipart/form-data">
<input type="file" name="files" multiple required><br><br>
<button>Process</button>
</form>
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

    login = pd.read_excel(paths[0], header=2)
    cdr = pd.read_excel(paths[1], header=1)
    agent = pd.read_excel(paths[2], header=2)
    crm = pd.read_excel(paths[3])

    # -------- LOGIN --------
    login["Date"] = pd.to_datetime(login["Date"])
    first_login = login.groupby("UserName")["Date"].min().dt.time

    # -------- AGENT --------
    agent["Total Break"] = agent["LUNCHBREAK"] + agent["SHORTBREAK"] + agent["TEABREAK"]
    agent["Total Meeting"] = agent["MEETING"] + agent["SYSTEMDOWN"]
    agent["Total Net Login"] = agent["Total Login Time"] - agent["Total Break"]

    # -------- CDR --------
    cdr["Mature"] = cdr["Disposition"].isin(["CALLMATURED","TRANSFER"])
    total_mature = cdr.groupby("Username")["Mature"].sum()

    transfer = cdr[cdr["Disposition"]=="TRANSFER"].groupby("Username").size()
    ib = cdr[(cdr["Disposition"].isin(["CALLMATURED","TRANSFER"])) & (cdr["Campaign"]=="CSRINBOUND")].groupby("Username").size()

    # -------- CRM --------
    tagging = crm.groupby("CreatedByID").size()

    # -------- MERGE --------
    result = agent.copy()
    result["Employee ID"] = result["Agent Name"]
    result["First Login Time"] = result["Employee ID"].map(first_login)
    result["Total Mature"] = result["Employee ID"].map(total_mature).fillna(0)
    result["Transfer Call"] = result["Employee ID"].map(transfer).fillna(0)
    result["IB Mature"] = result["Employee ID"].map(ib).fillna(0)
    result["OB Mature"] = result["Total Mature"] - result["IB Mature"]
    result["Total Tagging"] = result["Employee ID"].map(tagging).fillna(0)

    result["AHT"] = result["Total Talk Time"] / result["Total Mature"].replace(0,1)

    final = result[[
        "Employee ID","Agent Name","First Login Time",
        "Total Login Time","Total Net Login","Total Break","Total Meeting",
        "Total Talk Time","AHT","Total Mature","IB Mature",
        "Transfer Call","OB Mature","Total Tagging"
    ]]

    final.to_excel("output.xlsx", index=False)

    table_html = final.to_html(index=False)

    return render_template_string(
        HTML,
        table=table_html,
        download='<br><a href="/download"><button>Download Excel</button></a>'
    )

@app.route("/download")
def download():
    return send_file("output.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
