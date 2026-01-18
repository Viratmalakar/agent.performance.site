from flask import Flask, render_template, request, send_file
import pandas as pd
import os, tempfile, warnings

warnings.filterwarnings("ignore", category=UserWarning)

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

def time_fix(col):
    col = col.replace("-", "00:00:00")
    return pd.to_timedelta(col, errors="coerce").fillna(pd.Timedelta(0))

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():

    files = request.files.getlist("files")
    if len(files) != 2:
        return "Upload exactly 2 files (Agent Performance + CDR)"

    temp = tempfile.mkdtemp()

    agent_path = os.path.join(temp, files[0].filename)
    cdr_path   = os.path.join(temp, files[1].filename)

    files[0].save(agent_path)
    files[1].save(cdr_path)

    # Agent Performance: header from 3rd row
    agent = clean(pd.read_excel(agent_path, header=2, engine="openpyxl"))
    # CDR: header from 2nd row
    cdr   = clean(pd.read_excel(cdr_path, header=1, engine="openpyxl"))

    # ---- COLUMN MAP ----
    agent_name = find(agent, ["agent"])
    login_col  = find(agent, ["total login"])
    talk_col   = find(agent, ["total talk"])

    lunch = find(agent, ["lunch"])
    short = find(agent, ["short"])
    tea   = find(agent, ["tea"])
    meet  = find(agent, ["meeting"])
    system= find(agent, ["system"])

    cdr_user = find(cdr, ["username"])
    disp_col = find(cdr, ["disposition"])
    camp_col = find(cdr, ["campaign"])

    if not agent_name or not cdr_user or not disp_col:
        return "Required columns not found in files."

    # ---- TIME FIX ----
    for c in [login_col,talk_col,lunch,short,tea,meet,system]:
        if c:
            agent[c] = time_fix(agent[c])

    # ---- CALCULATIONS ----
    agent["total break"] = agent.get(lunch,0) + agent.get(short,0) + agent.get(tea,0)
    agent["total meeting"] = agent.get(meet,0) + agent.get(system,0)

    agent["total net login"] = agent[login_col] - agent["total break"]

    # ---- CDR CALCULATIONS ----
    cdr["user"] = cdr[cdr_user].astype(str)

    matured = cdr[disp_col].str.contains("callmatured|transfer", case=False, na=False)
    total_mature = matured.groupby(cdr["user"]).sum()

    transfer = cdr[disp_col].str.contains("transfer", case=False, na=False)
    transfer_cnt = transfer.groupby(cdr["user"]).sum()

    inbound = cdr[camp_col].str.contains("csr", case=False, na=False)
    ib_mature = (matured & inbound).groupby(cdr["user"]).sum()

    # ---- FINAL REPORT ----
    final = pd.DataFrame()
    final["EMP ID"] = agent[agent_name]
    final["Agent Name"] = agent[agent_name]

    final["Total Login"] = agent[login_col]
    final["Total Net Login"] = agent["total net login"]
    final["Total Break"] = agent["total break"]
    final["Total Meeting"] = agent["total meeting"]
    final["Total Talk Time"] = agent[talk_col]

    final["Total Mature"] = final["EMP ID"].map(total_mature).fillna(0).astype(int)
    final["IB Mature"] = final["EMP ID"].map(ib_mature).fillna(0).astype(int)
    final["Transfer Call"] = final["EMP ID"].map(transfer_cnt).fillna(0).astype(int)
    final["OB Mature"] = final["Total Mature"] - final["IB Mature"]

    final["AHT"] = final["Total Talk Time"] / final["Total Mature"].replace(0,1)

    # ---- EXPORT ----
    out = os.path.join(temp,"Agent_Final_Report.xlsx")
    final.to_excel(out,index=False, engine="openpyxl")

    return send_file(out, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
