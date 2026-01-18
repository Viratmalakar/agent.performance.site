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

    agent_file = request.files.get("agent_file")
    cdr_file = request.files.get("cdr_file")

    if not agent_file or not cdr_file:
        return "Both files required"

    temp = tempfile.mkdtemp()

    agent_path = os.path.join(temp, agent_file.filename)
    cdr_path = os.path.join(temp, cdr_file.filename)

    agent_file.save(agent_path)
    cdr_file.save(cdr_path)

    # ---------- READ ----------
    agent = pd.read_excel(agent_path, dtype=str)
    cdr   = pd.read_excel(cdr_path, dtype=str)

    agent = clean(agent)
    cdr   = clean(cdr)

    # ---------- DASH FIX ----------
    agent.replace("-", "00:00:00", inplace=True)

    # ---------- FIND ----------
    agent_name_col = find(agent, ["agent name","agent full name","agent"])
    login_col = find(agent, ["total login"])
    break_col = find(agent, ["total break"])
    meeting_col = find(agent, ["meeting"])

    cdr_user_col = find(cdr, ["username","user name","user"])

    if not agent_name_col or not cdr_user_col:
        return f"Column missing. Agent:{agent_name_col}, CDR:{cdr_user_col}"

    # ---------- CDR COUNT ----------
    call_count = cdr.groupby(cdr_user_col).size()

    # ---------- FINAL ----------
    final = pd.DataFrame()
    final["Agent Name"] = agent[agent_name_col]
    final["Total Login"] = agent.get(login_col, "00:00:00")
    final["Total Break"] = agent.get(break_col, "00:00:00")
    final["Meeting"] = agent.get(meeting_col, "00:00:00")

    final["Total Calls"] = final["Agent Name"].map(call_count).fillna(0).astype(int)

    # ---------- EXPORT ----------
    out = os.path.join(temp,"Agent_Report.xlsx")
    final.to_excel(out,index=False)

    return send_file(out, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
