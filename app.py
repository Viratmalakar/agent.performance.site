from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import tempfile

app = Flask(__name__)

# ------------------ PERMANENT COLUMN NORMALIZER ------------------
def normalize_columns(df):
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "")
        .str.replace("_", "")
    )
    return df

# ------------------ HOME ------------------
@app.route("/")
def index():
    return render_template("index.html")

# ------------------ PROCESS ------------------
@app.route("/process", methods=["POST"])
def process():

    files = request.files.getlist("files")

    if len(files) != 3:
        return "Please upload exactly 3 files (Agent, CDR, CRM)"

    temp_dir = tempfile.mkdtemp()
    paths = []

    for f in files:
        path = os.path.join(temp_dir, f.filename)
        f.save(path)
        paths.append(path)

    # -------- READ FILES --------
    agent = pd.read_excel(paths[0], header=0)
    cdr = pd.read_excel(paths[1], header=1)
    crm = pd.read_excel(paths[2], header=0)

    # -------- NORMALIZE --------
    agent = normalize_columns(agent)
    cdr = normalize_columns(cdr)
    crm = normalize_columns(crm)

    # -------- CREATE EMP ID --------
    agent["empid"] = agent["agentname"]
    cdr["empid"] = cdr["username"]
    crm["empid"] = crm["createdbyid"]

    # -------- SAFE NUMERIC CONVERT --------
    numeric_cols = agent.columns
    for c in numeric_cols:
        agent[c] = pd.to_numeric(agent[c], errors="ignore")

    # -------- TOTAL BREAK --------
    agent["totalbreak"] = (
        agent.get("lunchbreak", 0) +
        agent.get("shortbreak", 0) +
        agent.get("teabreak", 0)
    )

    # -------- TOTAL MEETING --------
    agent["totalmeeting"] = agent.get("meeting", 0)

    # -------- TAGGING --------
    tagging = crm.groupby("empid").size()

    # -------- MERGE --------
    final = agent.copy()
    final["totaltagging"] = final["empid"].map(tagging).fillna(0)

    # -------- RENAME OUTPUT --------
    final = final.rename(columns={
        "empid": "EMP ID",
        "agentname": "Agent Name",
        "totallogintime": "Total Login",
        "totalbreak": "Total Break",
        "totalmeeting": "Total Meeting",
        "totaltagging": "Total Tagging"
    })

    show_cols = [
        "EMP ID",
        "Agent Name",
        "Total Login",
        "Total Break",
        "Total Meeting",
        "Total Tagging"
    ]

    final = final[show_cols]

    # -------- SAVE FILE --------
    output_path = os.path.join(temp_dir, "Agent_Performance_Report.xlsx")
    final.to_excel(output_path, index=False)

    return send_file(output_path, as_attachment=True)

# ------------------ RUN ------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
