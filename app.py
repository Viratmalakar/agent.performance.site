from flask import Flask, render_template, request, send_file
import pandas as pd
import io

app = Flask(__name__)

def fix_time(x):
    if pd.isna(x) or x == "-" or str(x).strip()=="":
        return "00:00:00"
    return str(x)

def to_seconds(t):
    try:
        h,m,s = map(int,str(t).split(":"))
        return h*3600+m*60+s
    except:
        return 0

def to_hms(sec):
    h=sec//3600
    m=(sec%3600)//60
    s=sec%60
    return f"{h:02}:{m:02}:{s:02}"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():

    agent_file = request.files["agent"]
    cdr_file = request.files["cdr"]

    agent = pd.read_excel(agent_file, header=2).iloc[:,0:31]
    cdr = pd.read_excel(cdr_file, header=0).iloc[:,0:29]

    agent.columns = agent.columns.str.strip().str.lower()
    cdr.columns = cdr.columns.str.strip().str.lower()

    agent.rename(columns={
        "agent name":"employee_id",
        "agent full name":"agent_full_name",
        "total login time":"total_login",
        "total aux time":"total_break",
        "meeting":"total_meeting",
        "total talk time":"total_talk",
        "average call handling time":"aht",
        "lunchbreak":"lunch",
        "teabreak":"tea",
        "shortbreak":"short",
        "dinnerbreak":"dinner",
        "systemdown":"systemdown"
    }, inplace=True)

    cdr.rename(columns={
        "username":"employee_id",
        "call type":"call_type",
        "disposition":"disposition"
    }, inplace=True)

    for c in ["total_login","total_break","total_meeting","total_talk","aht","lunch","tea","short","dinner","systemdown"]:
        if c in agent:
            agent[c] = agent[c].apply(fix_time)

    agent["total_net_login"] = agent.apply(lambda r:
        to_hms(
            to_seconds(r.get("total_login")) -
            (to_seconds(r.get("lunch")) +
             to_seconds(r.get("tea")) +
             to_seconds(r.get("short")) +
             to_seconds(r.get("dinner")) +
             to_seconds(r.get("systemdown")))
        ), axis=1)

    mature = cdr[cdr["disposition"].str.contains("mature",case=False,na=False)]
    ib = mature[mature["call_type"].str.contains("inbound",case=False,na=False)]
    ob = mature[mature["call_type"].str.contains("outbound",case=False,na=False)]

    final = agent[["employee_id","agent_full_name","total_login","total_net_login","total_break","total_meeting","total_talk","aht"]].copy()

    final["Total Mature"] = final["employee_id"].map(mature.groupby("employee_id").size()).fillna(0).astype(int)
    final["IB Mature + Transfer call"] = final["employee_id"].map(ib.groupby("employee_id").size()).fillna(0).astype(int)
    final["OB Mature"] = final["employee_id"].map(ob.groupby("employee_id").size()).fillna(0).astype(int)

    final.columns = [
        "Agent Name","Agent Full Name","Total Login","Total Net Login",
        "Total Break","Total Meeting","Total Talk time","AHT",
        "Total Mature","IB Mature +Transfer call","OB Mature"
    ]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final.to_excel(writer,index=False)

    output.seek(0)
    return send_file(output,download_name="Final_Agent_Report.xlsx",as_attachment=True)

if __name__=="__main__":
    app.run(debug=True)
