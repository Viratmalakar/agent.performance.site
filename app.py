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

def fix_time(x):
    if pd.isna(x) or str(x).strip() in ["-", "", "nan"]:
        return "00:00:00"
    return str(x)

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

    # ---- READ FILES ----
    agent = clean(pd.read_excel(agent_path, header=2))
    agent = agent.iloc[:, :31]

    cdr = clean(pd.read_excel(cdr_path, header=1))
    cdr = cdr.iloc[:, :29]

    # ---- COLUMNS ----
    emp_col = find(agent, ["agent id","employee"])
    name_col = find(agent, ["agent name"])
    fullname_col = find(agent, ["full"])

    total_login_col = find(agent, ["total login time"])
    net_login_col = find(agent, ["net login"])
    talk_col = find(agent, ["total talk"])

    aht_col = find(agent, ["average call"])

    lunch = find(agent,["lunch"])
    short = find(agent,["short"])
    tea = find(agent,["tea"])
    meeting = find(agent,["meeting"])
    system = find(agent,["system"])

    # ---- FIX DASHES ----
    agent.iloc[:,1:31] = agent.iloc[:,1:31].astype(str)

    # ---- SAFE SERIES CREATOR ----
    def safe(col):
        if col:
            return agent[col].apply(fix_time)
        else:
            return pd.Series(["00:00:00"]*len(agent))

    # ---- BREAK ----
    def b(col):
        if col:
            return pd.to_timedelta(agent[col].apply(fix_time))
        return pd.to_timedelta("00:00:00")

    total_break = b(lunch)+b(short)+b(tea)+b(system)

    # ---- CDR MATURE ----
    disp_col = find(cdr,["disposition"])
    calltype_col = find(cdr,["call type"])

    matured = cdr[cdr[disp_col].str.contains("mature",case=False,na=False)]

    ib = matured[matured[calltype_col].str.contains("inbound",case=False,na=False)]
    ob = matured[matured[calltype_col].str.contains("outbound",case=False,na=False)]
    transfer = matured[matured[calltype_col].str.contains("transfer",case=False,na=False)]

    # ---- FINAL REPORT ----
    final = pd.DataFrame()

    final["Agent ID"] = agent[emp_col] if emp_col else ""
    final["Agent Name"] = agent[name_col] if name_col else ""
    final["Agent Full Name"] = agent[fullname_col] if fullname_col else ""

    final["Total Login"] = safe(total_login_col)
    final["Total Net Login"] = safe(net_login_col)

    final["Total Break"] = total_break.astype(str)
    final["Total Meeting"] = safe(meeting)

    final["Total Talk Time"] = safe(talk_col)
    final["AHT"] = safe(aht_col)

    final["Total Mature"] = len(matured)
    final["IB Mature + Transfer call"] = len(ib)+len(transfer)
    final["OB Mature"] = len(ob)

    # ---- EXPORT ----
    out = os.path.join(temp,"Agent_Report.xlsx")
    final.to_excel(out,index=False)

    return send_file(out,as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
