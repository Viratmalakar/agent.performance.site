from flask import Flask, render_template, request, send_file
import pandas as pd
import os, tempfile

app = Flask(__name__)

def clean(df):
    df.columns = df.columns.astype(str).str.strip().str.lower()
    return df

def find(df, keys):
    for c in df.columns:
        for k in keys:
            if k in c:
                return c
    return None

def fix_time(val):
    if str(val).strip() == "-":
        return "00:00:00"
    return val

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():

    files = request.files.getlist("files")
    if len(files) != 2:
        return "Upload exactly 2 files: Agent Performance + CDR"

    temp = tempfile.mkdtemp()

    agent_path = os.path.join(temp, files[0].filename)
    cdr_path   = os.path.join(temp, files[1].filename)

    files[0].save(agent_path)
    files[1].save(cdr_path)

    # ---- READ FILES ----
    agent = pd.read_excel(agent_path, header=2).iloc[:,0:31]   # A to AE
    cdr   = pd.read_excel(cdr_path, header=1).iloc[:,0:29]     # A to AC

    agent = clean(agent)
    cdr   = clean(cdr)

    # ---- FIX "-" TO 00:00:00 ----
    for col in agent.columns[1:31]:
        agent[col] = agent[col].apply(fix_time)

    # ---- COLUMN DETECT ----
    agent_col = find(agent, ["agent","name"])
    cdr_col   = find(cdr, ["username","user"])

    if not agent_col or not cdr_col:
        return f"Column missing. Agent:{agent_col}, CDR:{cdr_col}"

    agent["empid"] = agent[agent_col].astype(str)
    cdr["empid"]   = cdr[cdr_col].astype(str)

    # ---- TIME COLUMNS ----
    lunch = find(agent,["lunch"])
    short = find(agent,["short"])
    tea   = find(agent,["tea"])
    meet  = find(agent,["meeting"])
    login = find(agent,["login"])

    def safe(col):
        if col and col in agent.columns:
            return pd.to_timedelta(agent[col])
        return pd.to_timedelta("00:00:00")

    agent["totalbreak"] = safe(lunch) + safe(short) + safe(tea)
    agent["totallogin"] = safe(login)
    agent["totalmeeting"] = safe(meet)

    # ---- TAGGING FROM CDR ----
    tagging = cdr.groupby("empid").size()

    # ---- FINAL REPORT ----
    final = pd.DataFrame()
    final["EMP ID"] = agent["empid"]
    final["Agent Name"] = agent[agent_col]
    final["Total Login"] = agent["totallogin"].astype(str)
    final["Total Break"] = agent["totalbreak"].astype(str)
    final["Total Meeting"] = agent["totalmeeting"].astype(str)
    final["Total Calls"] = final["EMP ID"].map(tagging).fillna(0).astype(int)

    # ---- EXPORT ----
    out = os.path.join(temp, "Agent_Performance_Report.xlsx")

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        final.to_excel(writer,index=False,sheet_name="Report")

    return send_file(out, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
