from flask import Flask, render_template, request, send_file
import pandas as pd
import os, tempfile

app = Flask(__name__)

def clean(df):
    df.columns = df.columns.astype(str).str.strip().str.lower()
    return df

def to_sec(x):
    try:
        return pd.to_timedelta(x).total_seconds()
    except:
        return 0

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():

    files = request.files.getlist("files")
    if len(files) != 2:
        return "Upload only 2 files: CRM Performance + CDR"

    temp = tempfile.mkdtemp()
    paths = []

    for f in files:
        p = os.path.join(temp, f.filename)
        f.save(p)
        paths.append(p)

    # ===== READ FILES SAFELY =====
    agent = clean(pd.read_excel(paths[0], header=0))
    cdr   = clean(pd.read_excel(paths[1], header=0))

    print("AGENT:", list(agent.columns))
    print("CDR:", list(cdr.columns))

    # ===== REQUIRED COLUMNS =====
    emp_col = "agent name"
    login_col = "total login time"
    talk_col = "total talk time"

    lunch_col = "lunchbreak"
    short_col = "shortbreak"
    tea_col   = "teabreak"

    meeting_col = "meeting"
    system_col  = "systemdown"

    # ===== CHECK =====
    for c in [emp_col, login_col, talk_col]:
        if c not in agent.columns:
            return f"Missing column in Agent file: {c}"

    # ===== PROCESS =====
    agent["empid"] = agent[emp_col].astype(str)

    agent["totallogin"] = agent[login_col].apply(to_sec)
    agent["totaltalk"] = agent[talk_col].apply(to_sec)

    agent["totalbreak"] = (
        agent[lunch_col].apply(to_sec) if lunch_col in agent else 0 +
        agent[short_col].apply(to_sec) if short_col in agent else 0 +
        agent[tea_col].apply(to_sec) if tea_col in agent else 0
    )

    agent["totalmeeting"] = (
        agent[meeting_col].apply(to_sec) if meeting_col in agent else 0 +
        agent[system_col].apply(to_sec) if system_col in agent else 0
    )

    # ===== CDR Mature =====
    if "username" not in cdr.columns:
        return "Missing column in CDR: username"

    if "disposition" not in cdr.columns:
        return "Missing column in CDR: disposition"

    cdr["username"] = cdr["username"].astype(str)

    mature = cdr[cdr["disposition"].str.contains("call", case=False, na=False)]
    mature_count = mature.groupby("username").size()

    # ===== FINAL REPORT =====
    final = pd.DataFrame()
    final["EMP ID"] = agent["empid"]
    final["Agent Name"] = agent[emp_col]
    final["Total Login (sec)"] = agent["totallogin"]
    final["Total Break (sec)"] = agent["totalbreak"]
    final["Total Meeting (sec)"] = agent["totalmeeting"]
    final["Total Talk Time (sec)"] = agent["totaltalk"]
    final["Total Mature"] = final["EMP ID"].map(mature_count).fillna(0).astype(int)

    # ===== EXPORT =====
    out = os.path.join(temp, "Agent_Report.xlsx")
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        final.to_excel(writer, index=False)

    return send_file(out, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
