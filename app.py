from flask import Flask, render_template, request, send_file
import pandas as pd
import os, tempfile

app = Flask(__name__)

def clean(df):
    df.columns = df.columns.astype(str).str.strip().str.lower()
    return df

def to_seconds(x):
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
        return "Upload exactly 2 files: Agent Performance and CDR"

    temp = tempfile.mkdtemp()
    paths = []

    for f in files:
        p = os.path.join(temp, f.filename)
        f.save(p)
        paths.append(p)

    # Read files
    agent = clean(pd.read_excel(paths[0], header=2))
    cdr   = clean(pd.read_excel(paths[1], header=1))

    # ---- REQUIRED COLUMNS (from your file) ----
    emp_col = "agent name"
    login_col = "total login time"
    talk_col = "total talk time"

    lunch_col = "lunchbreak"
    short_col = "shortbreak"
    tea_col = "teabreak"

    meeting_col = "meeting"
    system_col = "systemdown"

    # ---- Convert ----
    agent["empid"] = agent[emp_col].astype(str)

    agent["totallogin"] = agent[login_col].apply(to_seconds)
    agent["totaltalk"] = agent[talk_col].apply(to_seconds)

    agent["totalbreak"] = (
        agent[lunch_col].apply(to_seconds) +
        agent[short_col].apply(to_seconds) +
        agent[tea_col].apply(to_seconds)
    )

    agent["totalmeeting"] = (
        agent[meeting_col].apply(to_seconds) +
        agent[system_col].apply(to_seconds)
    )

    # ---- CDR Mature ----
    cdr["username"] = cdr["username"].astype(str)

    mature = cdr[cdr["disposition"].str.contains("call", case=False, na=False)]
    mature_count = mature.groupby("username").size()

    # ---- FINAL REPORT ----
    final = pd.DataFrame()

    final["EMP ID"] = agent["empid"]
    final["Agent Name"] = agent[emp_col]
    final["Total Login (sec)"] = agent["totallogin"]
    final["Total Break (sec)"] = agent["totalbreak"]
    final["Total Meeting (sec)"] = agent["totalmeeting"]
    final["Total Talk Time (sec)"] = agent["totaltalk"]
    final["Total Mature"] = final["EMP ID"].map(mature_count).fillna(0).astype(int)

    # ---- EXPORT ----
    out = os.path.join(temp,"Agent_Report.xlsx")
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        final.to_excel(writer,index=False,sheet_name="Report")

    return send_file(out, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
