from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import io

app = Flask(__name__)

def fix_time(x):
    if pd.isna(x) or str(x).strip() == "-" or str(x).strip()=="":
        return "00:00:00"
    return str(x)

def time_to_sec(t):
    try:
        h,m,s = map(int,str(t).split(":"))
        return h*3600+m*60+s
    except:
        return 0

def sec_to_time(sec):
    h = sec//3600
    m = (sec%3600)//60
    s = sec%60
    return f"{h:02d}:{m:02d}:{s:02d}"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process",methods=["POST"])
def process():

    agent_file = request.files.get("agent")
    cdr_file = request.files.get("cdr")

    agent = pd.read_excel(agent_file,header=0)
    cdr = pd.read_excel(cdr_file,header=0)

    # ---------- AGENT PERFORMANCE CLEAN ----------
    agent.iloc[:,1:31] = agent.iloc[:,1:31].astype(str)
    agent.iloc[:,1:31] = agent.iloc[:,1:31].replace("-", "00:00:00")

    emp_col = agent.columns[1]        # Employee ID / Agent Name
    fullname_col = agent.columns[2]   # Agent Full Name

    login_col = agent.columns[3]      # D
    talk_col = agent.columns[5]       # F

    t_col = agent.columns[19]         # T
    u_col = agent.columns[20]         # U
    w_col = agent.columns[22]         # W
    x_col = agent.columns[23]         # X
    y_col = agent.columns[24]         # Y

    # ---------- CDR CLEAN ----------
    cdr_emp = cdr.columns[1]          # Username
    camp_col = cdr.columns[6]         # Campaign
    disp_col = cdr.columns[25]        # Z Disposition

    cdr[disp_col]=cdr[disp_col].astype(str)
    cdr[camp_col]=cdr[camp_col].astype(str)

    mature = cdr[cdr[disp_col].str.contains("callmatured|transfer",case=False,na=False)]
    ib = mature[mature[camp_col].str.contains("csr",case=False,na=False)]

    mature_cnt = mature.groupby(cdr_emp).size()
    ib_cnt = ib.groupby(cdr_emp).size()

    # ---------- FINAL REPORT ----------
    final = pd.DataFrame()
    final["Agent Name"] = agent[emp_col]
    final["Agent Full Name"] = agent[fullname_col]

    final["Total Login"] = agent[login_col].apply(fix_time)

    final["Total Break"] = (
        agent[t_col].apply(time_to_sec)+
        agent[w_col].apply(time_to_sec)+
        agent[y_col].apply(time_to_sec)
    ).apply(sec_to_time)

    final["Total Meeting"] = (
        agent[u_col].apply(time_to_sec)+
        agent[x_col].apply(time_to_sec)
    ).apply(sec_to_time)

    final["Total Net Login"] = (
        agent[login_col].apply(time_to_sec)
        - final["Total Break"].apply(time_to_sec)
    ).apply(sec_to_time)

    final["Total Talk Time"] = agent[talk_col].apply(fix_time)

    final["Total Mature"] = final["Agent Name"].map(mature_cnt).fillna(0).astype(int)
    final["IB Mature"] = final["Agent Name"].map(ib_cnt).fillna(0).astype(int)
    final["OB Mature"] = final["Total Mature"] - final["IB Mature"]

    final["AHT"] = (
        final["Total Talk Time"].apply(time_to_sec) /
        final["Total Mature"].replace(0,1)
    ).astype(int).apply(sec_to_time)

    final = final.fillna("00:00:00")

    return final.to_json(orient="records")

@app.route("/export",methods=["POST"])
def export():
    data = pd.read_json(request.data)
    out = io.BytesIO()
    data.to_excel(out,index=False)
    out.seek(0)
    return send_file(out,download_name="Final_Report.xlsx",as_attachment=True)
