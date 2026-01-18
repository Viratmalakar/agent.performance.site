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

def time_to_seconds(val):
    if pd.isna(val) or val == "-" or val == "":
        return 0
    try:
        if isinstance(val, str) and ":" in val:
            h, m, s = val.split(":")
            return int(h)*3600 + int(m)*60 + int(s)
        return float(val)
    except:
        return 0

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():

    files = request.files.getlist("files")

    if len(files) != 2:
        return "Upload only 2 files: Agent Performance & CDR"

    temp = tempfile.mkdtemp()
    paths = []

    for f in files:
        p = os.path.join(temp, f.filename)
        f.save(p)
        paths.append(p)

    # ===== READ FILES =====

    agent = pd.read_excel(paths[0], header=2)
    agent = agent.iloc[:, :31]   # B to AE zone safe
    agent = clean(agent)

    cdr = pd.read_excel(paths[1], header=1)
    cdr = cdr.iloc[:, :29]
    cdr = clean(cdr)

    # ===== CONVERT "-" TO 00:00:00 =====
    for col in agent.columns[1:31]:
        agent[col] = agent[col].apply(time_to_seconds)

    # ===== FIND COLUMNS =====

    agent_name_col = find(agent, ["agent", "name"])
    empid_col = find(agent, ["id", "emp", "createdby"])
    cdr_user_col = find(cdr, ["username", "user"])

    if not agent_name_col or not empid_col or not cdr_user_col:
        return f"""
        Column missing<br>
        Agent Name: {agent_name_col}<br>
        Emp ID: {empid_col}<br>
        CDR Username: {cdr_user_col}
        """

    agent["empid"] = agent[empid_col].astype(str)
    cdr["empid"] = cdr[cdr_user_col].astype(str)

    # ===== BREAK =====

    lunch = find(agent, ["lunch"])
    short = find(agent, ["short"])
    tea = find(agent, ["tea"])
    dinner = find(agent, ["dinner"])
    aux = find(agent, ["aux"])

    def safe(col):
        if col and col in agent:
            return agent[col]
        return 0

    agent["totalbreak"] = (
        safe(lunch) +
        safe(short) +
        safe(tea) +
        safe(dinner) +
        safe(aux)
    )

    # ===== LOGIN =====
    login = find(agent, ["login"])
    agent["totallogin"] = safe(login)

    # ===== MEETING =====
    meet = find(agent, ["meeting"])
    agent["totalmeeting"] = safe(meet)

    # ===== CALL COUNT =====
    callcount = cdr.groupby("empid").size()

    # ===== FINAL =====

    final = pd.DataFrame()
    final["EMP ID"] = agent["empid"]
    final["Agent Name"] = agent[agent_name_col]
    final["Total Login (sec)"] = agent["totallogin"]
    final["Total Break (sec)"] = agent["totalbreak"]
    final["Total Meeting (sec)"] = agent["totalmeeting"]
    final["Total Calls"] = final["EMP ID"].map(callcount).fillna(0).astype(int)

    # ===== EXPORT =====
    out = os.path.join(temp, "Agent_Performance_Report.xlsx")
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        final.to_excel(writer, index=False, sheet_name="Report")

    return send_file(out, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
