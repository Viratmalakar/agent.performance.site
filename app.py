from flask import Flask, render_template, request, send_file
import pandas as pd
import os, tempfile

app = Flask(__name__)

def clean_cols(df):
    df.columns = df.columns.astype(str).str.strip().str.lower()
    return df

def find_col(df, keys):
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
        return "Upload both Agent and CDR files"

    temp = tempfile.mkdtemp()

    agent_path = os.path.join(temp, agent_file.filename)
    cdr_path = os.path.join(temp, cdr_file.filename)

    agent_file.save(agent_path)
    cdr_file.save(cdr_path)

    # -------- SAFE READ --------
    agent = pd.read_excel(agent_path, header=None, dtype=str, engine="openpyxl")
    cdr = pd.read_excel(cdr_path, header=None, dtype=str, engine="openpyxl")

    # -------- ROW FIX --------
    agent = agent.iloc[2:].reset_index(drop=True)
    cdr = cdr.iloc[1:].reset_index(drop=True)

    # -------- COLUMN LIMIT --------
    agent = agent.iloc[:, :31]
    cdr = cdr.iloc[:, :29]

    # -------- SET HEADER --------
    agent.columns = agent.iloc[0]
    agent = agent[1:].reset_index(drop=True)

    cdr.columns = cdr.iloc[0]
    cdr = cdr[1:].reset_index(drop=True)

    # -------- CLEAN --------
    agent = clean_cols(agent)
    cdr = clean_cols(cdr)

    # -------- DASH FIX --------
    agent.replace("-", "00:00:00", inplace=True)

    # -------- COLUMN FIND --------
    agent_name = find_col(agent, ["agent name","agent full name","agent"])
    login = find_col(agent, ["total login"])
    break_col = find_col(agent, ["total break"])
    meeting = find_col(agent, ["meeting"])

    cdr_user = find_col(cdr, ["username","user name","user"])

    if not agent_name or not cdr_user:
        return f"Column missing. Agent:{agent_name}, CDR:{cdr_user}"

    # -------- CDR COUNT --------
    call_count = cdr.groupby(cdr_user).size()

    # -------- FINAL REPORT --------
    final = pd.DataFrame()
    final["Agent Name"] = agent[agent_name]
    final["Total Login"] = agent.get(login, "00:00:00")
    final["Total Break"] = agent.get(break_col, "00:00:00")
    final["Meeting"] = agent.get(meeting, "00:00:00")
    final["Total Calls"] = final["Agent Name"].map(call_count).fillna(0).astype(int)

    # -------- EXPORT --------
    out = os.path.join(temp, "Agent_Performance_Report.xlsx")
    final.to_excel(out, index=False, engine="openpyxl")

    return send_file(out, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
