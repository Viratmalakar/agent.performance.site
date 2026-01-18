from flask import Flask, render_template, request, send_file
import pandas as pd
import io

app = Flask(__name__)

def fix_time(x):
    if pd.isna(x) or str(x).strip() in ["", "-"]:
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

    # ===== READ FILES =====
    agent = pd.read_excel(agent_file, header=2).iloc[:,0:31]
    cdr = pd.read_excel(cdr_file, header=0).iloc[:,0:29]

    agent.columns = agent.columns.str.strip()
    cdr.columns = cdr.columns.str.strip()

    # ===== FIX TIMES =====
    time_cols = [
        "Total Login Time","Total Aux Time","Meeting",
        "Total Talk Time","Average Call Handling Time",
        "Lunch Break","Tea Break","Short Break","Dinner Break","System Down"
    ]

    for c in time_cols:
        if c in agent.columns:
            agent[c] = agent[c].apply(fix_time)

    # ===== NET LOGIN =====
    agent["Total Net Login"] = agent.apply(lambda r:
        to_hms(
            to_seconds(r["Total Login Time"]) -
            (to_seconds(r.get("Lunch Break","00:00:00")) +
             to_seconds(r.get("Tea Break","00:00:00")) +
             to_seconds(r.get("Short Break","00:00:00")) +
             to_seconds(r.get("Dinner Break","00:00:00")) +
             to_seconds(r.get("System Down","00:00:00")))
        ), axis=1)

    # ===== MATURE FILTER =====
    mature = cdr[cdr["Disposition"].astype(str).str.contains("mature",case=False,na=False)]
    ib = mature[mature["Call Type"].astype(str).str.contains("in",case=False,na=False)]
    ob = mature[mature["Call Type"].astype(str).str.contains("out",case=False,na=False)]

    # ===== FINAL REPORT =====
    final = pd.DataFrame()
    final["Agent Name"] = agent["Agent Name"]
    final["Agent Full Name"] = agent["Agent Full Name"]
    final["Total Login"] = agent["Total Login Time"]
    final["Total Net Login"] = agent["Total Net Login"]
    final["Total Break"] = agent["Total Aux Time"]
    final["Total Meeting"] = agent["Meeting"]
    final["Total Talk time"] = agent["Total Talk Time"]
    final["AHT"] = agent["Average Call Handling Time"]

    final["Total Mature"] = final["Agent Name"].map(mature.groupby("Username").size()).fillna(0).astype(int)
    final["IB Mature +Transfer call"] = final["Agent Name"].map(ib.groupby("Username").size()).fillna(0).astype(int)
    final["OB Mature"] = final["Agent Name"].map(ob.groupby("Username").size()).fillna(0).astype(int)

    # ===== EXPORT =====
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final.to_excel(writer,index=False)

    output.seek(0)
    return send_file(output,download_name="Final_Agent_Report.xlsx",as_attachment=True)

if __name__=="__main__":
    app.run()
