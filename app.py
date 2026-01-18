from flask import Flask,render_template,request,send_file
import pandas as pd,io,json
from datetime import datetime

app=Flask(__name__)

def tsec(t):
    try:
        h,m,s=map(int,str(t).split(":"))
        return h*3600+m*60+s
    except: return 0

def stime(s):
    h=s//3600; m=(s%3600)//60; s=s%60
    return f"{h:02d}:{m:02d}:{s:02d}"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process",methods=["POST"])
def process():

    agent=pd.read_excel(request.files["agent"])
    cdr=pd.read_excel(request.files["cdr"])

    agent.iloc[:,1:31]=agent.iloc[:,1:31].astype(str).replace("-", "00:00:00")

    emp=agent.columns[1]
    full=agent.columns[2]
    login=agent.columns[3]
    talk=agent.columns[5]

    t=agent.columns[19]; u=agent.columns[20]
    w=agent.columns[22]; x=agent.columns[23]; y=agent.columns[24]

    c_emp=cdr.columns[1]
    camp=cdr.columns[6]
    disp=cdr.columns[25]

    cdr[disp]=cdr[disp].astype(str)
    cdr[camp]=cdr[camp].astype(str)

    mature=cdr[cdr[disp].str.contains("callmature|transfer",case=False,na=False)]
    ib=mature[mature[camp].str.upper()=="CSRINBOUND"]

    mature_cnt=mature.groupby(c_emp).size()
    ib_cnt=ib.groupby(c_emp).size()

    final=pd.DataFrame()
    final["Agent Name"]=agent[emp]
    final["Agent Full Name"]=agent[full]
    final["Total Login Time"]=agent[login]

    final["Total Break"]=(agent[t].apply(tsec)+agent[w].apply(tsec)+agent[y].apply(tsec)).apply(stime)
    final["Total Meeting"]=(agent[u].apply(tsec)+agent[x].apply(tsec)).apply(stime)

    final["Total Net Login"]=(agent[login].apply(tsec)-final["Total Break"].apply(tsec)).apply(stime)

    final["Total Call"]=final["Agent Name"].map(mature_cnt).fillna(0).astype(int)
    final["IB Mature"]=final["Agent Name"].map(ib_cnt).fillna(0).astype(int)
    final["OB Mature"]=final["Total Call"]-final["IB Mature"]

    final["AHT"]=(agent[talk].apply(tsec)/final["Total Call"].replace(0,1)).astype(int).apply(stime)

    final=final[~final["Agent Name"].astype(str).str.lower().isin(["nan","agent name"])]
    final=final.dropna(how="all")

    # ---- GRAND TOTAL ROW ----
    total={}
    total["Agent Name"]="TOTAL"
    total["Agent Full Name"]=""

    total["Total Login Time"]=stime(final["Total Login Time"].apply(tsec).sum())
    total["Total Net Login"]=stime(final["Total Net Login"].apply(tsec).sum())
    total["Total Break"]=stime(final["Total Break"].apply(tsec).sum())
    total["Total Meeting"]=stime(final["Total Meeting"].apply(tsec).sum())

    total["Total Call"]=int(final["Total Call"].sum())
    total["IB Mature"]=int(final["IB Mature"].sum())
    total["OB Mature"]=int(final["OB Mature"].sum())

    total["AHT"]=stime(int(final["AHT"].apply(tsec).mean()))

    ivr_hit=len(cdr[cdr[camp].str.upper()=="CSRINBOUND"])

    total["TOTAL IVR HIT"]=ivr_hit
    total["TOTAL MATURE"]=int(final["Total Call"].sum())
    total["TOTAL TALK TIME"]=stime(agent[talk].apply(tsec).sum())
    total["LOGIN COUNT"]=len(final)

    final=final[[
    "Agent Name","Agent Full Name","Total Login Time","Total Net Login",
    "Total Break","Total Meeting","AHT","Total Call","IB Mature","OB Mature"
    ]]

    final=pd.concat([final,pd.DataFrame([total])])

    return final.to_json(orient="records")

@app.route("/export",methods=["POST"])
def export():
    data=pd.DataFrame(json.loads(request.data))
    out=io.BytesIO()
    data.to_excel(out,index=False)
    out.seek(0)

    now=datetime.now().strftime("%d-%m-%y %H-%M-%S")
    return send_file(out,download_name=f"Agent_Performance_Report_Chandan-Malakar_{now}.xlsx",as_attachment=True)
