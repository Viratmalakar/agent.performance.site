from flask import Flask, request, send_file
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os

app = Flask(__name__)

EXCEL_FILE = "Final_Agent_Report.xlsx"
PDF_FILE = "Final_Agent_Report.pdf"

THEME = """
<style>
body{
margin:0;font-family:Arial;
background:linear-gradient(135deg,#7f5cff,#ff9acb,#6ec6ff);
background-size:400% 400%;
animation:bg 12s infinite;
}
@keyframes bg{
0%{background-position:0% 50%}
50%{background-position:100% 50%}
100%{background-position:0% 50%}
}
.box{
width:90%;margin:40px auto;
background:white;padding:25px;border-radius:15px;text-align:center;
}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{border:1px solid #999;padding:6px}
th{background:#ffd27f}
.btn{padding:10px 20px;background:#2196f3;color:white;border-radius:6px;border:none;cursor:pointer}
</style>
"""

LAST_HTML = ""

@app.route("/")
def home():
    return THEME + """
<div class='box'>
<h2>Upload 4 Excel Reports</h2>
<form action="/process" method="post" enctype="multipart/form-data">
<input type="file" name="files" multiple required><br><br>
<button class='btn'>Upload & Calculate</button>
</form>
</div>
"""

@app.route("/process", methods=["POST"])
def process():
    global LAST_HTML

    files=request.files.getlist("files")

    login=cdr=agent=crm=None

    for f in files:
        n=f.filename.lower()
        df=pd.read_excel(f)
        if "login" in n: login=df
        elif "cdr" in n: cdr=df
        elif "agent" in n: agent=df
        elif "crm" in n: crm=df

    if any(x is None for x in [login,cdr,agent,crm]):
        return "All 4 reports not detected"

    login["Duration"]=pd.to_timedelta(login["Duration"],errors="coerce").dt.total_seconds()
    total_login=login.groupby("UserName")["Duration"].sum()

    break_mask=login["Activity"].str.contains("lunch|short|tea",case=False,na=False)
    total_break=login[break_mask].groupby("UserName")["Duration"].sum()

    meeting_mask=login["Activity"].str.contains("meeting|system",case=False,na=False)
    total_meeting=login[meeting_mask].groupby("UserName")["Duration"].sum()

    net_login=total_login-total_break

    agent["Talk Time"]=pd.to_timedelta(agent["Talk Time"],errors="coerce").dt.total_seconds()
    total_talk=agent.groupby("Agent Name")["Talk Time"].sum()

    mature_mask=cdr["Disposition"].str.contains("callmatured|transfer",case=False,na=False)
    total_mature=cdr[mature_mask].groupby("Username").size()

    ib_mask=mature_mask & cdr["Campaign"].str.contains("csr",case=False,na=False)
    ib_mature=cdr[ib_mask].groupby("Username").size()

    transfer_call=cdr[cdr["Disposition"].str.contains("transfer",case=False,na=False)].groupby("Username").size()

    ob_mature=total_mature-ib_mature

    total_tagging=crm.groupby("CreatedByID").size()

    final=pd.DataFrame({
        "EMP ID":total_login.index,
        "Agent Name":total_login.index,
        "Total Login":total_login.values,
        "Total Net Login":net_login.reindex(total_login.index).fillna(0).values,
        "Total Break":total_break.reindex(total_login.index).fillna(0).values,
        "Total Meeting":total_meeting.reindex(total_login.index).fillna(0).values,
        "Total Talk time":total_talk.reindex(total_login.index).fillna(0).values,
        "Total Mature":total_mature.reindex(total_login.index).fillna(0).values,
        "IB Mature":ib_mature.reindex(total_login.index).fillna(0).values,
        "Transfer Call":transfer_call.reindex(total_login.index).fillna(0).values,
        "OB Mature":ob_mature.reindex(total_login.index).fillna(0).values,
        "Total tagging":total_tagging.reindex(total_login.index).fillna(0).values
    })

    final["AHT"]=(final["Total Talk time"]/final["Total Mature"]).fillna(0).round(2)

    for c in ["Total Login","Total Net Login","Total Break","Total Meeting","Total Talk time"]:
        final[c]=pd.to_timedelta(final[c],unit="s")

    final.to_excel(EXCEL_FILE,index=False)
    create_pdf(final)

    LAST_HTML = THEME + f"""
<div class='box'>
<h2>Final Agent Report</h2>
{final.to_html(index=False)}
<br><br>
<a href='/download_excel'>Download Excel</a> |
<a href='/download_pdf'>Download PDF</a>
</div>
"""

    return LAST_HTML

@app.route("/download_excel")
def download_excel():
    return send_file(EXCEL_FILE,as_attachment=True)

@app.route("/download_pdf")
def download_pdf():
    return send_file(PDF_FILE,as_attachment=True)

def create_pdf(df):
    c=canvas.Canvas(PDF_FILE,pagesize=A4)
    w,h=A4
    c.setFont("Helvetica-Bold",10)
    c.drawRightString(w-40,h-40,"Chandan Malakar")
    y=h-80
    c.setFont("Helvetica",8)
    for _,r in df.iterrows():
        line=" | ".join([str(x) for x in r])
        c.drawString(40,y,line)
        y-=12
        if y<50:
            c.showPage()
            c.drawRightString(w-40,h-40,"Chandan Malakar")
            y=h-80
    c.save()

if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0",port=port)
