from flask import Flask, render_template, request, send_file
import pandas as pd
import os, tempfile

app = Flask(__name__)

# ---------------- SAFE COLUMN NORMALIZER ----------------
def clean_cols(df):
    df.columns = df.columns.astype(str).str.strip().str.lower()
    return df

def find_col(df, keywords):
    for c in df.columns:
        col = str(c).lower()
        for k in keywords:
            if k in col:
                return c
    return None

# ---------------- HOME ----------------
@app.route("/")
def index():
    return render_template("index.html")

# ---------------- PROCESS ----------------
@app.route("/process", methods=["POST"])
def process():

    files = request.files.getlist("files")
    if len(files) != 3:
        return "Upload exactly 3 files"

    temp = tempfile.mkdtemp()
    paths = []

    for f in files:
        p = os.path.join(temp, f.filename)
        f.save(p)
        paths.append(p)

    # ---------- READ ----------
    agent = pd.read_excel(paths[0])
    cdr = pd.read_excel(paths[1], header=1)
    crm = pd.read_excel(paths[2])

    agent = clean_cols(agent)
    cdr = clean_cols(cdr)
    crm = clean_cols(crm)

    # ---------- AUTO DETECT IDS ----------
    agent_name_col = find_col(agent, ["agent"])
    cdr_user_col = find_col(cdr, ["user"])
    crm_emp_col = find_col(crm, ["created"])

    if not agent_name_col or not cdr_user_col or not crm_emp_col:
        return f"Column missing. Agent:{agent_name_col}, CDR:{cdr_user_col}, CRM:{crm_emp_col}"

    agent["empid"] = agent[agent_name_col]
    cdr["empid"] = cdr[cdr_user_col]
    crm["empid"] = crm[crm_emp_col]

    # ---------- BREAK ----------
    lunch = find_col(agent, ["lunch"])
    short = find_col(agent, ["short"])
    tea = find_col(agent, ["tea"])

    agent["totalbreak"] = agent.get(lunch,0)+agent.get(short,0)+agent.get(tea,0)

    # ---------- MEETING ----------
    meet = find_col(agent, ["meeting"])
    agent["totalmeeting"] = agent.get(meet,0)

    # ---------- LOGIN ----------
    login = find_col(agent, ["login"])
    agent["totallogin"] = agent.get(login,0)

    # ---------- TAGGING ----------
    tagging = crm.groupby("empid").size()

    # ---------- FINAL ----------
    final = agent.copy()
    final["Total Tagging"] = final["empid"].map(tagging).fillna(0)

    final = final.rename(columns={
        "empid":"EMP ID",
        agent_name_col:"Agent Name",
        "totallogin":"Total Login",
        "totalbreak":"Total Break",
        "totalmeeting":"Total Meeting"
    })

    show = [
        "EMP ID",
        "Agent Name",
        "Total Login",
        "Total Break",
        "Total Meeting",
        "Total Tagging"
    ]

    final = final[show]

    out = os.path.join(temp,"Agent_Report.xlsx")
    final.to_excel(out,index=False)

    return send_file(out, as_attachment=True)

# ---------------- RUN ----------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)

