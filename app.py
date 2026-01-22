from flask import Flask, render_template, request, redirect, url_for, session, send_file
import pandas as pd
import io
from datetime import datetime

app = Flask(__name__)
app.secret_key = "agentdashboard"

def tsec(t):
    try:
        h,m,s=map(int,str(t).split(":"))
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

@app.route("/process",methods=["POST"])
def process():

    agent=pd.read_excel(request.files["agent"])
    cdr=pd.read_excel(request.files["cdr"])

    agent=agent.fillna("00:00:00").replace("-","00:00:00")

    emp=agent.columns[1]
    full=agent.columns[2]
    login=agent.columns[3]
    talk=agent.columns[5]
    t=agent.columns[19]; u=agent.columns[20]
    w=agent.columns[22]; x=agent.columns[23]; y=agent.columns[24]

    c_emp=cdr.columns[1]
    camp=cdr.columns[6]
    disp=cdr.columns[25]

    mature=cdr[cdr[disp].astype(str).str.contains("callmature|transfer",case=False,na=False)]
    ib=mature[mature[camp].astype(str).str.upper()=="CSRINBOUND"]

    mature_cnt=mature.groupby(c_emp).size()
    ib_cnt=ib.groupby(c_emp).size()

    mature_cnt.index=mature_cnt.index.astype(str).str.strip()
    ib_cnt.index=ib_cnt.index.astype(str).str.strip()

    final=pd.DataFrame()
    final["Agent Name"]=agent[emp].astype(str).str.strip()
    final["Agent Full Name"]=agent[full]
    final["Total Login Time"]=agent[login]

    final["Total Break"]=(agent[t].apply(tsec)+agent[w].apply(tsec)+agent[y].apply(tsec)).apply(stime)
    final["Total Meeting"]=(agent[u].apply(tsec)+agent[x].apply(tsec)).apply(stime)
    final["Total Net Login"]=(agent[login].apply(tsec)-final["Total Break"].apply(tsec)).apply(stime)
    final["AHT"]=agent[talk]

    final["Total Call"]=final["Agent Name"].map(mature_cnt).fillna(0).astype(int)
    final["IB Mature"]=final["Agent Name"].map(ib_cnt).fillna(0).astype(int)
    final["OB Mature"]=final["Total Call"]-final["IB Mature"]

    # Highlight rules
    final["__red_net"]=(final["Total Net Login"].apply(tsec)<8*3600) & (final["Total Login Time"].apply(tsec)>8*3600+15*60)
    final["__red_break"]=final["Total Break"].apply(tsec)>40*60
    final["__red_meet"]=final["Total Meeting"].apply(tsec)>35*60

    session["data"]=final.to_dict(orient="records")

    gt={}
    gt["TOTAL IVR HIT"]=len(cdr)
    gt["TOTAL MATURE"]=int(final["Total Call"].sum())
    gt["IB MATURE"]=int(final["IB Mature"].sum())
    gt["OB MATURE"]=int(final["OB Mature"].sum())
    gt["TOTAL TALK TIME"]=stime(agent[talk].apply(tsec).sum())
    gt["AHT"]=stime(int(agent[talk].apply(tsec).sum()/max(1,gt["TOTAL MATURE"])))
    gt["LOGIN COUNT"]=len(final)

    session["gt"]=gt

    return redirect(url_for("result"))

@app.route("/result")
def result():
    return render_template("result.html",data=session["data"],gt=session["gt"])
