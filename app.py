from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import io
from datetime import datetime

app = Flask(__name__)

def tsec(t):
    try:
        h,m,s = map(int,str(t).split(":"))
        return h*3600+m*60+s
    except:
        return 0

def stime(s):
    h=s//3600
    m=(s%3600)//60
    s=s%60
    return f"{h:02d}:{m:02d}:{s:02d}"

@app.route("/")
def upload():
    return render_template("upload.html")

@app.route("/dashboard")
def dashboard():
    return render_template("dashboard.html")

@app.route("/process", methods=["POST"])
def process():

    agent = pd.read_excel(request.files["agent"])
    cdr = pd.read_excel(request.files["cdr"])

    agent = agent.fillna("00:00:00")

    emp = agent.columns[1]
    full = agent.columns[2]
    login = agent.columns[3]
    talk = agent.columns[5]

    t = agent.columns[19]
    u = agent.columns[20]
    w = agent.columns[22]
    x = agent.columns[23]
    y = agent.columns[24]

    c_emp = cdr.columns[1]
    camp = cdr.columns[6]
    disp = cdr.columns[25]

    cdr[disp] = cdr[disp].astype(str)
    cdr[camp] = cdr[camp].astype(str)

    mature = cdr[cdr[disp].str.contains("callmature|transfer",case=False)]
    ib = mature[mature[camp].str.upper()=="CSRINBOUND"]

    mature_cnt = mature.groupby(c_emp).size()
    ib_cnt = ib.groupby(c_emp).size()

    final = pd.DataFrame()
    final["Agent Name"] = agent[emp]
    final["Agent Full Name"] = agent[full]
    final["Total Login Time"] = agent[login]
    final["Total Net Login"] = (agent[login].apply(tsec) -
                               (agent[t].apply(tsec)+agent[w].apply(tsec)+agent[y].apply(tsec))).apply(stime)

    final["Total Break"] = (agent[t].apply(tsec)+agent[w].apply(tsec)+agent[y].apply(tsec)).apply(stime)
    final["Total Meeting"] = (agent[u].apply(tsec)+agent[x].apply(tsec)).apply(stime)
    final["AHT"] = (agent[talk].apply(tsec)/mature_cnt.reindex(agent[emp]).fillna(1)).astype(int).apply(stime)

    final["Total Call"] = agent[emp].map(mature_cnt).fillna(0).astype(int)
    final["IB Mature"] = agent[emp].map(ib_cnt).fillna(0).astype(int)
    final["OB Mature"] = final["Total Call"] - final["IB Mature"]

    final = final[~final["Agent Name"].astype(str).str.lower().isin(["nan","agent name"])]
    final = final.reset_index(drop=True)

    # -------- Grand Totals ----------
    total_ivr = (cdr[camp].str.upper()=="CSRINBOUND").sum()
    total_mature = mature.shape[0]
    total_ib = ib.shape[0]
    total_ob = total_mature-total_ib
    total_talk = stime(agent[talk].apply(tsec).sum())
    avg_aht = stime(int(final["AHT"].apply(tsec).mean()))
    login_count = final["Agent Name"].nunique()

    return jsonify({
        "cards":{
            "ivr":int(total_ivr),
            "mature":int(total_mature),
            "ib":int(total_ib),
            "ob":int(total_ob),
            "talk":total_talk,
            "aht":avg_aht,
            "login":int(login_count)
        },
        "table":final.to_dict(orient="records")
    })

@app.route("/export", methods=["POST"])
def export():
    data = pd.DataFrame(request.json)
    out = io.BytesIO()
    data.to_excel(out,index=False)
    out.seek(0)

    fname = "Agent_Performance_Report_"+datetime.now().strftime("%d-%m-%Y_%H-%M-%S")+".xlsx"
    return send_file(out,download_name=fname,as_attachment=True)

if __name__=="__main__":
    app.run(debug=True)
