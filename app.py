from flask import Flask, request, send_file, render_template_string
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

HTML = """
<h2>Upload 4 Excel Reports</h2>
<form method="post" action="/process" enctype="multipart/form-data">
<input type="file" name="files" multiple required><br><br>
<button>Process</button>
</form>
<br>
{{table}}
{{download}}
"""

def load_excel_safe(path, header_row):
    df = pd.read_excel(path, header=header_row, engine="openpyxl", dtype=str)
    df = df.fillna("")
    return df

@app.route("/")
def home():
    return render_template_string(HTML, table="", download="")

@app.route("/process", methods=["POST"])
def process():

    files = request.files.getlist("files")
    paths = []

    for f in files:
        p = os.path.join(UPLOAD_FOLDER, f.filename)
        f.save(p)
        paths.append(p)

    # -------- Load all as TEXT ----------
    login = load_excel_safe(paths[0], 2)
    cdr   = load_excel_safe(paths[1], 1)
    agent = load_excel_safe(paths[2], 2)
    crm   = load_excel_safe(paths[3], 0)

    # -------- LOGIN ----------
    date_col = None
    for c in login.columns:
        if "date" in c.lower():
            date_col = c
            break

    login[date_col] = pd.to_datetime(login[date_col], errors="coerce")
    first_login = login.groupby("UserName")[date_col].min().dt.time

    # -------- AGENT ----------
    for col in ["LUNCHBREAK","SHORTBREAK","TEABREAK","MEETING","SYSTEMDOWN",
                "Total Login Time","Total Talk Time"]:
        agent[col] = pd.to_timedelta(agent[col], errors="coerce").dt.total_seconds()

    agent["Total Break"] = agent["LUNCHBREAK"] + agent["SHORTBREAK"] + agent["TEABREAK"]
    agent["Total Meeting"] = agent["MEETING"] + agent["SYSTEMDOWN"]
    agent["Total Net Login"] = agent["Total Login Time"] - agent["Total Break"]

    # -------- CDR ----------
    cdr["Mature"] = cdr["Disposition"].isin(["CALLMATURED","TRANSFER"])

    total_mature = cdr.groupby("Username")["Mature"].sum()
    transfer_call = cdr[cdr["Disposition"]=="TRANSFER"].groupby("Username").size()
    ib_mature = cdr[
        (cdr["Disposition"].isin(["CALLMATURED","TRANSFER"])) &
        (cdr["Campaign"]=="CSRINBOUND")
    ].groupby("Username").size()

    # -------- CRM ----------
    total_tagging = crm.groupby("CreatedByID").size()

    # -------- FINAL ----------
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

    result["AHT"] = result["Total Talk Time"] / result["Total Mature"].replace(0,1)

    # Convert seconds back to HH:MM:SS
    for c in ["Total Login","Total Net Login","Total Break","Total Meeting","Total Talk Time","AHT"]:
        result[c] = pd.to_timedelta(result[c], unit="s")

    output = "Final_Report.xlsx"
    result.to_excel(output, index=False)

    return render_template_string(
        HTML,
        table=result.to_html(index=False),
        download='<br><a href="/download"><button>Download Excel</button></a>'
    )

@app.route("/download")
def download():
    return send_file("Final_Report.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
