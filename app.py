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

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():

    files = request.files.getlist("files")
    if len(files) != 3:
        return "Upload exactly 3 files in order: Agent, CDR, CRM"

    temp = tempfile.mkdtemp()
    paths = []

    for f in files:
        p = os.path.join(temp, f.filename)
        f.save(p)
        paths.append(p)

    # ---- READ FILES CORRECT ----
    agent = clean(pd.read_excel(paths[0], header=2))   # Agent Performance
    cdr   = clean(pd.read_excel(paths[1], header=1))   # CDR
    crm   = clean(pd.read_excel(paths[2], header=2))   # CRM

    print("AGENT:", agent.columns.tolist())
    print("CDR:", cdr.columns.tolist())
    print("CRM:", crm.columns.tolist())

    # ---- COLUMN MAP ----
    agent_id = find(agent, ["agent","name","emp","id"])
    crm_id   = find(crm, ["createdbyid","created","emp","id"])
    cdr_id   = find(cdr, ["username","user"])

    if not agent_id or not crm_id:
        return f"Column missing. Agent:{agent_id}, CRM:{crm_id}"

    agent["empid"] = agent[agent_id].astype(str)
    crm["empid"]   = crm[crm_id].astype(str)

    # ---- BREAK ----
    lunch = find(agent,["lunch"])
    short = find(agent,["short"])
    tea   = find(agent,["tea"])

    agent["totalbreak"] = (
        agent.get(lunch,0).fillna(0) +
        agent.get(short,0).fillna(0) +
        agent.get(tea,0).fillna(0)
    )

    # ---- LOGIN ----
    login = find(agent,["login"])
    agent["totallogin"] = agent.get(login,0).fillna(0)

    # ---- MEETING ----
    meet = find(agent,["meeting","system"])
    agent["totalmeeting"] = agent.get(meet,0).fillna(0)

    # ---- TAGGING ----
    tagging = crm.groupby("empid").size()

    # ---- FINAL REPORT ----
    final = pd.DataFrame()
    final["EMP ID"] = agent["empid"]
    final["Agent Name"] = agent[agent_id]
    final["Total Login"] = agent["totallogin"]
    final["Total Break"] = agent["totalbreak"]
    final["Total Meeting"] = agent["totalmeeting"]
    final["Total Tagging"] = final["EMP ID"].map(tagging).fillna(0).astype(int)

    # ---- EXPORT ----
    out = os.path.join(temp,"Agent_Report.xlsx")
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        final.to_excel(writer,index=False,sheet_name="Report")

    return send_file(out, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
