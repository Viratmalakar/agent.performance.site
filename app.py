from flask import Flask, render_template, request, send_file
import pandas as pd
import io, json
from datetime import datetime

app = Flask(__name__)

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
    h=s//3600
    m=(s%3600)//60
    s=s%60
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

    ivr_total = (cdr[cdr[camp].str.upper()=="CSRINBOUND"].shape[0])

    final=pd.DataFrame()
    final["Agent Name"]=agent[emp]
    final["Agent Full Name"]=agent[full]
    final["Total Login"]=agent[login].apply(fix_time)

    final["Total Break"]=(agent[t].apply(tsec)+agent[w].apply(tsec)+agent[y].apply(tsec)).apply(stime)
    final["Total Meeting"]=(agent[u].apply(tsec)+agent[x].apply(tsec)).apply(stime)

    final["Total Net Login"]=(agent[login].apply(tsec)-final["Total Break"].apply(tsec)).apply(stime)

    final["Total Talk Time"]=agent[talk].apply(fix_time)

    final["Total Mature"]=final["Agent Name"].map(mature_cnt).fillna(0).astype(int)
    final["IB Mature"]=final["Agent Name"].map(ib_cnt).fillna(0).astype(int)
    final["OB Mature"]=final["Total Mature"]-final["IB Mature"]

    final["Total IVR Hit"]=final["Agent Name"].map(
        cdr[cdr[camp].str.upper()=="CSRINBOUND"].groupby(c_emp).size()
    ).fillna(0).astype(int)

    final["AHT"]=(final["Total Talk Time"].apply(tsec)/final["Total Mature"].replace(0,1)).astype(int).apply(stime)

    final=final[~final["Agent Name"].astype(str).str.lower().isin(["nan","agent name"])]
    final=final.dropna(how="all")

    # -------- GRAND TOTAL ----------
    gt={}
    gt["Agent Name"]="GRAND TOTAL"
    gt["Agent Full Name"]=""
    gt["Total Login"]=""   # removed as per instruction
    gt["Total Break"]=""
    gt["Total Meeting"]=""
    gt["Total Net Login"]=""
    gt["Total Talk Time"]=stime(final["Total Talk Time"].apply(tsec).sum())

    total_mature = mature.shape[0]
    total_ib = ib.shape[0]
    total_ob = total_mature-total_ib

    gt["Total Mature"]=total_mature
    gt["IB Mature"]=total_ib
    gt["OB Mature"]=total_ob
    gt["Total IVR Hit"]=ivr_total

    if total_mature>0:
        gt["AHT"]=stime(final["Total Talk Time"].apply(tsec).sum()//total_mature)
    else:
        gt["AHT"]="00:00:00"

    final=pd.concat([final,pd.DataFrame([gt])],ignore_index=True)

    return final.to_json(orient="records")

@app.route("/export",methods=["POST"])
def export():
    data=pd.DataFrame(json.loads(request.data.decode()))

    out=io.BytesIO()

    with pd.ExcelWriter(out,engine="openpyxl") as writer:
        data.to_excel(writer,index=False,sheet_name="Report")
        ws=writer.book["Report"]

        # style
        from openpyxl.styles import Font, PatternFill, Border, Side

        header_fill=PatternFill("solid",fgColor="1FA463")
        header_font=Font(color="FFFFFF",bold=True)
        border=Border(left=Side(style="thin"),right=Side(style="thin"),
                      top=Side(style="thin"),bottom=Side(style="thin"))

        for cell in ws[1]:
            cell.fill=header_fill
            cell.font=header_font
            cell.border=border

        for row in ws.iter_rows():
            for c in row:
                c.border=border

    out.seek(0)

    now=datetime.now().strftime("%d-%m-%y %H-%M-%S")
    fname=f"Agent_Performance_Report_Chandan-Malakar & {now}.xlsx"

    return send_file(out,download_name=fname,as_attachment=True)
