from flask import Flask, request, render_template, send_file
import pandas as pd
import io
import json

app = Flask(__name__)

def fix_time(val):
    if pd.isna(val) or str(val).strip() in ["-", ""]:
        return "00:00:00"
    return str(val)

def time_to_seconds(t):
    try:
        h,m,s = map(int,str(t).split(":"))
        return h*3600+m*60+s
    except:
        return 0

def seconds_to_time(sec):
    h=sec//3600
    m=(sec%3600)//60
    s=sec%60
    return f"{h:02}:{m:02}:{s:02}"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():

    agent = pd.read_excel(request.files["agent_file"], skiprows=2, usecols="A:AE")
    cdr = pd.read_excel(request.files["cdr_file"], skiprows=1, usecols="A:AC")

    agent.iloc[:,1:31] = agent.iloc[:,1:31].astype(str).replace("-", "00:00:00")

    emp_col = agent.columns[1]
    total_login_col = agent.columns[3]
    talk_time_col = agent.columns[5]

    t_col = agent.columns[19]
    u_col = agent.columns[20]
    w_col = agent.columns[22]
    x_col = agent.columns[23]
    y_col = agent.columns[24]

    for c in [total_login_col,talk_time_col,t_col,u_col,w_col,x_col,y_col]:
        agent[c]=agent[c].apply(fix_time)

    agent["BreakSec"]=agent[t_col].apply(time_to_seconds)+agent[w_col].apply(time_to_seconds)+agent[y_col].apply(time_to_seconds)
    agent["MeetingSec"]=agent[u_col].apply(time_to_seconds)+agent[x_col].apply(time_to_seconds)
    agent["LoginSec"]=agent[total_login_col].apply(time_to_seconds)
    agent["NetLoginSec"]=agent["LoginSec"]-agent["BreakSec"]
    agent["TalkSec"]=agent[talk_time_col].apply(time_to_seconds)

    cdr_emp=cdr.columns[1]
    disp_col=cdr.columns[25]
    camp_col=cdr.columns[5]

    mature=cdr[cdr[disp_col].astype(str).str.contains("Callmatured|Transfer",case=False,na=False)]
    ib=mature[mature[camp_col].astype(str).str.contains("csrinbound",case=False,na=False)]

    total_mature=mature.groupby(cdr_emp).size()
    ib_mature=ib.groupby(cdr_emp).size()

    final=pd.DataFrame()
    final["Agent Name"]=agent[emp_col]
    final["Total Login"]=agent[total_login_col]
    final["Total Net Login"]=agent["NetLoginSec"].apply(seconds_to_time)
    final["Total Break"]=agent["BreakSec"].apply(seconds_to_time)
    final["Total Meeting"]=agent["MeetingSec"].apply(seconds_to_time)
    final["Total Talk Time"]=agent["TalkSec"].apply(seconds_to_time)

    final["Total Mature"]=agent[emp_col].map(total_mature).fillna(0).astype(int)
    final["IB Mature"]=agent[emp_col].map(ib_mature).fillna(0).astype(int)
    final["OB Mature"]=final["Total Mature"]-final["IB Mature"]

    final["AHT"]=final.apply(lambda r: seconds_to_time(int(time_to_seconds(r["Total Talk Time"])/r["Total Mature"])) if r["Total Mature"]>0 else "00:00:00",axis=1)

    table_html=final.to_html(index=False)

    return render_template("index.html",
        tables=table_html,
        raw_data=final.to_json(orient="records")
    )

@app.route("/export", methods=["POST"])
def export():
    data=json.loads(request.form["data"])
    df=pd.DataFrame(data)

    output=io.BytesIO()
    with pd.ExcelWriter(output,engine="openpyxl") as writer:
        df.to_excel(writer,index=False)

    output.seek(0)
    return send_file(output,download_name="Final_Report.xlsx",as_attachment=True)

if __name__=="__main__":
    app.run(debug=True)
