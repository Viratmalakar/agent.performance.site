from flask import Flask, render_template, request, send_file
import pandas as pd
import os, tempfile

app = Flask(__name__)

# ---------------- HELPERS ----------------

def clean_columns(df):
    df.columns = df.columns.astype(str).str.strip().str.lower()
    return df

def find_col(df, keys):
    for c in df.columns:
        for k in keys:
            if k in c:
                return c
    return None

def fix_time(val):
    if pd.isna(val) or str(val).strip() == "-":
        return "00:00:00"
    return str(val)

# ---------------- ROUTES ----------------

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():

    agent_file = request.files.get("agent_file")
    cdr_file = request.files.get("cdr_file")

    if not agent_file or not cdr_file:
        return "Please upload Agent Performance and CDR files both"

    temp = tempfile.mkdtemp()

    agent_path = os.path.join(temp, agent_file.filename)
    cdr_path   = os.path.join(temp, cdr_file.filename)

    agent_file.save(agent_path)
    cdr_file.save(cdr_path)

    # ---------------- READ FILES ----------------

    agent = pd.read_excel(agent_path, skiprows=2)
    cdr   = pd.read_excel(cdr_path, skiprows=1)

    agent = clean_columns(agent)
    cdr   = clean_columns(cdr)

    # Replace "-" with 00:00:00
    agent = agent.replace("-", "00:00:00")

    # ---------------- COLUMN MAP ----------------

    agent_id_col = find_col(agent, ["agent id","username","login"])
    agent_name_col = find_col(agent, ["agent full name","full name","agent name"])

    login_col = find_col(agent, ["total login"])
    net_login_col = find_col(agent, ["net login"])
    break_col = find_col(agent, ["total break"])
    meeting_col = find_col(agent, ["meeting"])
    talk_col = find_col(agent, ["talk"])
    aht_col = find_col(agent, ["aht"])

    cdr_user_col = find_col(cdr, ["username"])
    disposition_col = find_col(cdr, ["disposition"])

    # ---------------- SAFE CAST ----------------

    agent["agent_id"] = agent.get(agent_id_col,"").astype(str)
    agent["agent_name"] = agent.get(agent_name_col,"").astype(str)

    # ---------------- CDR CALC ----------------

    if cdr_user_col and disposition_col:
        mature = cdr.groupby(cdr_user_col)[disposition_col].count()
    else:
        mature = pd.Series(dtype=int)

    # ---------------- FINAL REPORT ----------------

    final = pd.DataFrame()

    final["Agent ID"] = agent["agent_id"]
    final["Agent Full Name"] = agent["agent_name"]

    final["Total Login"] = agent.get(login_col,"00:00:00").apply(fix_time)
    final["Total Net Login"] = agent.get(net_login_col,"00:00:00").apply(fix_time)
    final["Total Break"] = agent.get(break_col,"00:00:00").apply(fix_time)
    final["Total Meeting"] = agent.get(meeting_col,"00:00:00").apply(fix_time)
    final["Total Talk Time"] = agent.get(talk_col,"00:00:00").apply(fix_time)
    final["AHT"] = agent.get(aht_col,"00:00:00").apply(fix_time)

    final["Total Mature"] = final["Agent ID"].map(mature).fillna(0).astype(int)
    final["IB Mature + Transfer call"] = 0
    final["OB Mature"] = 0

    # ---------------- EXPORT ----------------

    out = os.path.join(temp,"Agent_Report.xlsx")

    final.to_excel(out,index=False,engine="openpyxl")

    return send_file(out, as_attachment=True)

# ---------------- RUN ----------------

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
