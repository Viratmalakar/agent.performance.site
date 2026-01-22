from flask import Flask, render_template, request, redirect, url_for, session, send_file
import pandas as pd
import io, os
from datetime import datetime

app = Flask(__name__)
app.secret_key = os.urandom(24)

# ---------- TIME HELPERS ----------
def fix_time(x):
    if pd.isna(x) or str(x).strip().lower() in ["-","nan",""]:
        return "00:00:00"
    return str(x)

def tsec(t):
    try:
        h,m,s=map(int,str(t).split(":"))
        return h*3600+m*60+s
    except:
        return 0

def stime(s):
    s=int(max(0,s))
    h=s//3600
    m=(s%3600)//60
    s=s%60
    return f"{h:02d}:{m:02d}:{s:02d}"

def is_red_netlogin(total_login, net_login):
    return tsec(total_login) > tsec("08:15:00") and tsec(net_login) < tsec("08:00:00")

# ---------- ROUTES ----------
@app.route("/")
def upload():
    return render_template("upload.html")

@app.route("/process",methods=["POST"])
def process():

    agent=pd.read_excel(request.files["agent"],engine="openpyxl")
    cdr=pd.read_excel(request.files["cdr"],engine="openpyxl")

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
    final["Total Login Time"]=agent[login].apply(fix_time)

    final["Total Break"]=(agent[t].apply(tsec)+agent[w].apply(tsec)+agent[y].apply(tsec)).apply(stime)
    final["Total Meeting"]=(agent[u].apply(tsec)+agent[x].apply(tsec)).apply(stime)
    final["Total Net Login"]=(agent[login].apply(tsec)-final["Total Break"].apply(tsec)).clip(lower=0).apply(stime)

    final["Total Talk Time"]=agent[talk].apply(fix_time)

    final["Total Call"]=final["Agent Name"].map(mature_cnt).fillna(0).astype(int)
    final["IB Mature"]=final["Agent Name"].map(ib_cnt).fillna(0).astype(int)
    final["OB Mature"]=final["Total Call"]-final["IB Mature"]

    final["AHT"]=(final["Total Talk Time"].apply(tsec)/final["Total Call"].replace(0,1)).astype(int).apply(stime)

    final = final[final["Agent Name"].notna()]

    # 🔴 HIGHLIGHT FLAGS
    final["__red_net"]=final.apply(lambda r: is_red_netlogin(r["Total Login Time"],r["Total Net Login"]),axis=1)
    final["__red_break"]=final["Total Break"].apply(lambda x: tsec(x)>tsec("00:40:00"))
    final["__red_meet"]=final["Total Meeting"].apply(lambda x: tsec(x)>tsec("00:30:00"))

    total_talk_sec = final["Total Talk Time"].apply(tsec).sum()
    total_call = int(final["Total Call"].sum())

    gt={
        "TOTAL IVR HIT":int(cdr[cdr[camp].str.upper()=="CSRINBOUND"].shape[0]),
        "TOTAL MATURE":total_call,
        "IB MATURE":int(final["IB Mature"].sum()),
        "OB MATURE":int(final["OB Mature"].sum()),
        "TOTAL TALK TIME":stime(total_talk_sec),
        "AHT":stime(int(total_talk_sec/max(1,total_call))),
        "LOGIN COUNT":int(final["Agent Name"].count())
    }

    final = final[[
    "Agent Name","Agent Full Name","Total Login Time","Total Net Login",
    "Total Break","Total Meeting","AHT","Total Call","IB Mature","OB Mature",
    "__red_net","__red_break","__red_meet"
    ]]

    session["data"]=final.to_dict(orient="records")
    session["gt"]=gt

    return redirect(url_for("result"))

@app.route("/result")
def result():
    return render_template("result.html",data=session["data"],gt=session["gt"])

@app.route("/export")
def export():

    data=pd.DataFrame(session["data"])
    flags=["__red_net","__red_break","__red_meet"]
    excel=data.drop(columns=flags)

    out=io.BytesIO()
    with pd.ExcelWriter(out,engine="xlsxwriter") as writer:
        excel.to_excel(writer,index=False,startrow=1,sheet_name="Report")

        wb=writer.book
        ws=writer.sheets["Report"]

        header=wb.add_format({"bold":True,"bg_color":"#1fa463","color":"white","border":1,"align":"center"})
        cell=wb.add_format({"border":1,"align":"center"})
        red=wb.add_format({"border":1,"align":"center","font_color":"red","bold":True})

        for c in range(len(excel.columns)):
            ws.write(0,c,excel.columns[c],header)
            ws.set_column(c,c,22)

        for r in range(len(excel)):
            for c in range(len(excel.columns)):
                fmt=cell
                if c==3 and data.iloc[r]["__red_net"]: fmt=red
                if c==4 and data.iloc[r]["__red_break"]: fmt=red
                if c==5 and data.iloc[r]["__red_meet"]: fmt=red
                ws.write(r+1,c,excel.iloc[r,c],fmt)

    out.seek(0)
    fname=f"Agent_Performance_Report_{datetime.now().strftime('%d-%m-%y %H-%M-%S')}.xlsx"
    return send_file(out,download_name=fname,as_attachment=True)

if __name__=="__main__":
    app.run()
