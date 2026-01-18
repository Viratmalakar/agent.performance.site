from flask import Flask, render_template, request, send_file
import pandas as pd
import os, tempfile

app = Flask(__name__)

def clean(df):
    df.columns = df.columns.astype(str).str.strip().str.lower()
    return df

def fix_time_cells(df):
    df.replace("-", "00:00:00", inplace=True)
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
    if len(files) != 2:
        return "Upload exactly 2 files (Agent Performance + CDR)"

    temp = tempfile.mkdtemp()
    paths = []

    for f in files:
        p = os.path.join(temp, f.filename)
        f.save(p)
        paths.append(p)

    # ---------- LOAD FILES ----------
    agent = pd.read_excel(paths[0], header=2)
    cdr   = pd.read_excel(paths[1], header=1)

    agent = clean(agent)
    cdr   = clean(cdr)

    # ---------- FIX DASH ----------
    agent.iloc[:,1:31] = agent.iloc[:,1:31].astype(str)
    agent = fix_time_cells(agent)

    # ---------- DETECT COLUMNS ----------
    agent_name_col = find(agent, ["agent name","agent full","name"])
    login_col = find(agent, ["total login"])
    meeting_col = find(agent, ["meeting"])
    break_col = find(agent, ["total break"])

    cdr_user_col = find(cdr, ["username","user"])

    if not agent_name_col or not cdr_user_col:
        return f"Column missing. Agent:{agent_name_col}, CDR:{cdr_user_col}"

    # ---------- CDR COUNT ----------
    call_count = cdr.groupby(cdr_user_col).size()

    # ---------- FINAL REPORT ----------
    final = pd.DataFrame()
    final["Agent Name"] = agent[agent_name_col]

    if login_col:
        final["Total Login"] = agent[login_col]
    else:
        final["Total Login"] = "00:00:00"

    if break_col:
        final["Total Break"] = agent[break_col]
    else:
        final["Total Break"] = "00:00:00"

    if meeting_col:
        final["Total Meeting"] = agent[meeting_col]
    else:
        final["Total Meeting"] = "00:00:00"

    final["Total Calls"] = final["Agent Name"].map(call_count).fillna(0).astype(int)

    # ---------- EXPORT ----------
    out = os.path.join(temp,"Agent_Report.xlsx")
    final.to_excel(out,index=False)

    return send_file(out, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
