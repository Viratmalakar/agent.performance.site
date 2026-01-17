from flask import Flask, request, send_file
import pandas as pd
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

app = Flask(__name__)

RESULT_EXCEL = "Final_Agent_Report.xlsx"
RESULT_PDF = "Final_Agent_Report.pdf"

# ---------------- HOME ----------------
@app.route("/")
def home():
    return """
    <h2>Upload 4 Reports</h2>
    <form action="/process" method="post" enctype="multipart/form-data">
        <input type="file" name="files" multiple required><br><br>
        <button>Upload & Calculate</button>
    </form>
    """

# ---------------- PROCESS ----------------
@app.route("/process", methods=["POST"])
def process():
    files = request.files.getlist("files")

    login=cdr=agent=crm=None

    for f in files:
        name=f.filename.lower()
        df=pd.read_excel(f)

        if "login" in name:
            login=df
        elif "cdr" in name:
            cdr=df
        elif "agent" in name or "performance" in name:
            agent=df
        elif "crm" in name or "detail" in name:
            crm=df

    if any(x is None for x in [login,cdr,agent,crm]):
        return "All 4 reports not detected"

    # ---------------- LOGIN REPORT ----------------
    login["Duration"] = pd.to_timedelta(login["Duration"], errors="coerce").dt.total_seconds()

    total_login = login.groupby("UserName")["Duration"].sum()

    break_mask = login["Activity"].str.contains("lunch|short|tea",case=False,na=False)
    total_break = login[break_mask].groupby("UserName")["Duration"].sum()

    meeting_mask = login["Activity"].str.contains("meeting|system",case=False,na=False)
    total_meeting = login[meeting_mask].groupby("UserName")["Duration"].sum()

    net_login = total_login - total_break

    # ---------------- AGENT PERFORMANCE ----------------
    agent["Talk Time"] = pd.to_timedelta(agent["Talk Time"], errors="coerce").dt.total_seconds()
    total_talk = agent.groupby("Agent Name")["Talk Time"].sum()

    # ---------------- CDR ----------------
    mature_mask = cdr["Disposition"].str.contains("callmatured|transfer",case=False,na=False)
    total_mature = cdr[mature_mask].groupby("Username").size()

    ib_mask = mature_mask & cdr["Campaign"].str.contains("csr",case=False,na=False)
    ib_mature = cdr[ib_mask].groupby("Username").size()

    transfer_call = cdr[cdr["Disposition"].str.contains("transfer",case=False,na=False)].groupby("Username").size()

    ob_mature = total_mature - ib_mature

    # ---------------- CRM ----------------
    total_tagging = crm.groupby("CreatedByID").size()

    # ---------------- FINAL MERGE ----------------
    final = pd.DataFrame({
        "EMP ID": total_login.index,
        "Agent Name": total_login.index,
        "Total Login": total_login.values,
        "Total Net Login": net_login.reindex(total_login.index).fillna(0).values,
        "Total Break": total_break.reindex(total_login.index).fillna(0).values,
        "Total Meeting": total_meeting.reindex(total_login.index).fillna(0).values,
        "Total Talk time": total_talk.reindex(total_login.index).fillna(0).values,
        "Total Mature": total_mature.reindex(total_login.index).fillna(0).values,
        "IB Mature": ib_mature.reindex(total_login.index).fillna(0).values,
        "Transfer Call": transfer_call.reindex(total_login.index).fillna(0).values,
        "OB Mature": ob_mature.reindex(total_login.index).fillna(0).values,
        "Total tagging": total_tagging.reindex(total_login.index).fillna(0).values
    })

    final["AHT"] = (final["Total Talk time"] / final["Total Mature"]).fillna(0).round(2)

    # seconds → HH:MM:SS
    for col in ["Total Login","Total Net Login","Total Break","Total Meeting","Total Talk time"]:
        final[col] = pd.to_timedelta(final[col], unit="s")

    final.to_excel(RESULT_EXCEL,index=False)

    create_pdf(final)

    return final.to_html(index=False) + """
    <br><br>
    <a href='/download_excel'>Download Excel</a> |
    <a href='/download_pdf'>Download PDF</a>
    """

# ---------------- PDF ----------------
def create_pdf(df):
    c=canvas.Canvas(RESULT_PDF,pagesize=A4)
    w,h=A4
    c.setFont("Helvetica-Bold",10)
    c.drawRightString(w-40,h-40,"Chandan Malakar")

    y=h-80
    c.setFont("Helvetica",8)

    for col in df.columns:
        c.drawString(40,y,str(col)); y-=12

    y-=10

    for _,row in df.iterrows():
        for val in row:
            c.drawString(40,y,str(val)); y-=12
        y-=10
        if y<50:
            c.showPage()
            c.drawRightString(w-40,h-40,"Chandan Malakar")
            y=h-80

    c.save()

# ---------------- DOWNLOAD ----------------
@app.route("/download_excel")
def download_excel():
    return send_file(RESULT_EXCEL,as_attachment=True)

@app.route("/download_pdf")
def download_pdf():
    return send_file(RESULT_PDF,as_attachment=True)

# ---------------- RUN ----------------
if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0",port=port)
