from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import io
from datetime import datetime

app = Flask(__name__)

def fix_time(x):
    if pd.isna(x) or str(x).strip() in ["-","nan","NaN",""]:
        return "00:00:00"
    return str(x)

def time_to_sec(t):
    try:
        h,m,s = map(int,str(t).split(":"))
        return h*3600+m*60+s
    except:
        return 0

def sec_to_time(sec):
    h=sec//3600
    m=(sec%3600)//60
    s=sec%60
    return f"{h:02d}:{m:02d}:{s:02d}"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process",methods=["POST"])
def process():

    agent = pd.read_excel(request.files["agent"])
    cdr = pd.read_excel(request.files["cdr"])

    agent.iloc[:,1:31] = agent.iloc[:,1:31].astype(str).replace("-", "00:00:00")

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

    cdr[disp]=cdr[disp].astype(str)
    cdr[camp]=cdr[camp].astype(str)

    # Total Mature
    mature = cdr[cdr[disp].str.contains("callmature|transfer",case=False,na=False)]

    # IB Mature only CSRINBOUND
    ib = mature[mature[camp].str.upper()=="CSRINBOUND"]

    mature_cnt = mature.groupby(c_emp).size()
    ib_cnt = ib.groupby(c_emp).size()

    final = pd.DataFrame()
    final["Agent Name"]=agent[emp]
    final["Agent Full Name"]=agent[full]
    final["Total Login Time"]=agent[login].apply(fix_time)

    final["Total Break"]=(
        agent[t].apply(time_to_sec)+agent[w].apply(time_to_sec)+agent[y].apply(time_to_sec)
    ).apply(sec_to_time)

    final["Total Meeting"]=(
        agent[u].apply(time_to_sec)+agent[x].apply(time_to_sec)
    ).apply(sec_to_time)

    final["Total Net Login"]=(
        agent[login].apply(time_to_sec)-final["Total Break"].apply(time_to_sec)
    ).apply(sec_to_time)

    final["Total Talk Time"]=agent[talk].apply(fix_time)

    final["Total Mature"]=final["Agent Name"].map(mature_cnt).fillna(0).astype(int)
    final["IB Mature"]=final["Agent Name"].map(ib_cnt).fillna(0).astype(int)
    final["OB Mature"]=final["Total Mature"]-final["IB Mature"]

    final["AHT"]=(
        final["Total Talk Time"].apply(time_to_sec)/
        final["Total Mature"].replace(0,1)
    ).astype(int).apply(sec_to_time)

    final=final.dropna(how="all")
    final=final[final["Agent Name"].astype(str).str.lower()!="nan"]

    return final.to_json(orient="records")

@app.route("/export",methods=["POST"])
def export():
    data=pd.read_json(request.data)
    out=io.BytesIO()
    data.to_excel(out,index=False)
    out.seek(0)

    now=datetime.now().strftime("%d-%m-%y_%H-%M-%S")
    name=f"Agent_Performance_Report_Chandan-Malakar_{now}.xlsx"

    return send_file(out,download_name=name,as_attachment=True)
