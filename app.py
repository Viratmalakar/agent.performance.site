from flask import Flask, request, render_template, send_file
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def to_seconds(t):
    try:
        return pd.to_timedelta(str(t)).total_seconds()
    except:
        return 0

def find_col(df, keyword):
    for c in df.columns:
        if keyword.lower() in str(c).lower():
            return c
    return None

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload")
def upload():
    return index()

@app.route("/process", methods=["POST"])
def process():

    files = request.files.getlist("files")

    if len(files) != 3:
        return "Please upload exactly 3 files"

    paths = []
    for f in files:
        path = os.path.join(UPLOAD_FOLDER, f.filename)
        f.save(path)
        paths.append(path)

    agent = pd.read_excel(paths[0], header=2)
    cdr   = pd.read_excel(paths[1], header=1)
    crm   = pd.read_excel(paths[2], header=0)

    agent.columns = agent.columns.astype(str)
    cdr.columns = cdr.columns.astype(str)
    crm.columns = crm.columns.astype(str)

    # ---- Find ID columns ----
    agent_id = find_col(agent, "agent")
    cdr_id = find_col(cdr, "user")
    crm_id = find_col(crm, "created")

    if not crm_id:
        return "CRM Employee ID column not found"

    agent.rename(columns={agent_id:"EMP ID"}, inplace=True)
    cdr.rename(columns={cdr_id:"EMP ID"}, inplace=True)
    crm.rename(columns={crm_id:"EMP ID"}, inplace=True)

    # ---- Time columns ----
    time_cols = ["Total Login Time","LUNCHBREAK","SHORTBREAK","TEABREAK",
                 "MEETING","SYSTEMDOWN","Total Talk Time"]

    for c in time_cols:
        if c in agent.columns:
            agent[c] = agent[c].apply(to_seconds)
        else:
            agent[c] = 0

    agent["Total Break"] = agent["LUNCHBREAK"] + agent["SHORTBREAK"] + agent["TEABREAK"]
    agent["Total Meeting"] = agent["MEETING"] + agent["SYSTEMDOWN"]
    agent["Total Net Login"] = agent["Total Login Time"] - agent["Total Break"]

    # ---- CDR ----
    cdr["Disposition"] = cdr["Disposition"].astype(str)

    mature = cdr[cdr["Disposition"].str.contains("CALLMATURED|TRANSFER", case=False, na=False)]

    total_mature = mature.groupby("EMP ID").size()
    ib_mature = mature[mature["Campaign"]=="CSRINBOUND"].groupby("EMP ID").size()
    transfer_call = cdr[cdr["Disposition"].str.contains("TRANSFER", case=False, na=False)].groupby("EMP ID").size()

    # ---- CRM ----
    tagging = crm.groupby("EMP ID").size()

    # ---- Final ----
    final = agent.copy()

    final["Total Mature"] = final["EMP ID"].map(total_mature).fillna(0).astype(int)
    final["IB Mature"] = final["EMP ID"].map(ib_mature).fillna(0).astype(int)
    final["Transfer Call"] = final["EMP ID"].map(transfer_call).fillna(0).astype(int)
    final["OB Mature"] = final["Total Mature"] - final["IB Mature"]
    final["Total Tagging"] = final["EMP ID"].map(tagging).fillna(0).astype(int)

    final["AHT"] = final["Total Talk Time"] / final["Total Mature"].replace(0,1)

    def sec_to_time(x):
        try:
            return str(pd.to_timedelta(int(x), unit="s"))
        except:
            return "00:00:00"

    for c in ["Total Login Time","Total Net Login","Total Break","Total Meeting","Total Talk Time","AHT"]:
        final[c] = final[c].apply(sec_to_time)

    final = final[[
        "EMP ID","Total Login Time","Total Net Login","Total Break",
        "Total Meeting","Total Talk Time","AHT","Total Mature",
        "IB Mature","Transfer Call","OB Mature","Total Tagging"
    ]]

    output = "final_report.xlsx"
    final.to_excel(output,index=False)

    return render_template("result.html",
        tables=final.to_html(index=False, classes="table table-bordered"),
        file=output
    )

@app.route("/download")
def download():
    return send_file("final_report.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
