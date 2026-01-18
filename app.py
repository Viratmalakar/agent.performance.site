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

    files = request.files.getlist("files")
    if len(files) != 2:
        return "Upload exactly 2 files (Agent Performance + CDR)"

    temp = tempfile.mkdtemp()
    paths = []

    for f in files:
        p = os.path.join(temp, f.filename)
        f.save(p)
        paths.append(p)

    # ---------- LOAD ----------
    agent = pd.read_excel(paths[0])
    cdr   = pd.read_excel(paths[1])

    agent = clean(agent)
    cdr   = clean(cdr)

    # ---------- FIX DASH ----------
    agent.replace("-", "00:00:00", inplace=True)

    # ---------- DETECT ----------
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

    final["Total Login"] = agent[login_col] if login_col else "00:00:00"
    final["Total Break"] = agent[break_col] if break_col else "00:00:00"
    final["Total Meeting"] = agent[meeting_col] if meeting_col else "00:00:00"

    final["Total Calls"] = final["Agent Name"].map(call_count).fillna(0).astype(int)

    # ---------- EXPORT ----------
    out = os.path.join(temp,"Agent_Report.xlsx")
    final.to_excel(out,index=False)

    return send_file(out, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
