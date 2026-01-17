from flask import Flask, request, send_file
import pandas as pd
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

app = Flask(__name__)

RESULT_EXCEL = "Final_Agent_Report.xlsx"
RESULT_PDF = "Final_Agent_Report.pdf"

THEME = """
<style>
body{margin:0;font-family:Arial;background:linear-gradient(135deg,#7f5cff,#ff9acb,#6ec6ff);background-size:400% 400%;animation:bg 12s infinite;}
@keyframes bg{0%{background-position:0% 50%}50%{background-position:100% 50%}100%{background-position:0% 50%}}
.box{width:90%;margin:50px auto;background:white;padding:25px;border-radius:15px;text-align:center;}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{border:1px solid #999;padding:6px}
th{background:#ffd27f}
.btn{display:inline-block;padding:10px 20px;background:#2196f3;color:white;border-radius:6px;margin:10px;text-decoration:none;font-weight:bold}
</style>
"""

@app.route("/")
def home():
    return THEME + """
    <div class='box'>
    <h2>Upload 4 Excel Reports</h2>
    <form action='/process' method='post' enctype='multipart/form-data'>
        <input type='file' name='files' multiple required><br><br>
        <button class='btn'>Upload & Calculate</button>
    </form>
    </div>
    """

@app.route("/process", methods=["POST"])
def process():
    files = request.files.getlist("files")

    login = cdr = agent = crm = None

    for f in files:
        df = pd.read_excel(f)
        cols = " ".join(df.columns.astype(str)).lower()

        if "login" in cols and "logout" in cols:
            login = df
        elif "disposition" in cols:
            cdr = df
        elif "talk" in cols:
            agent = df
        elif "created" in cols:
            crm = df

    if any(x is None for x in [login, cdr, agent, crm]):
        return THEME + "<div class='box'><h3>❌ Could not detect all 4 reports automatically.</h3></div>"

    # ---------------- LOGIN REPORT ----------------
    emp_col = login.columns[0]
    dur_col = login.columns[1]
    status_col = login.columns[2]

    login[dur_col] = pd.to_timedelta(login[dur_col], errors='coerce').dt.total_seconds()

    total_login = login.groupby(emp_col)[dur_col].sum()

    break_mask = login[status_col].str.contains("lunch|short|tea", case=False, na=False)
    total_break = login[break_mask].groupby(emp_col)[dur_col].sum()

    meeting_mask = login[status_col].str.contains("meeting|system", case=False, na=False)
    total_meeting = login[meeting_mask].groupby(emp_col)[dur_col].sum()

    net_login = total_login - total_break

    # ---------------- AGENT PERFORMANCE ----------------
    emp2 = agent.columns[0]
    talk_col = agent.columns[1]
    agent[talk_col] = pd.to_timedelta(agent[talk_col], errors='coerce').dt.total_seconds()
    total_talk = agent.groupby(emp2)[talk_col].sum()

    # ---------------- CDR REPORT ----------------
    emp3 = cdr.columns[0]
    disp_col = cdr.columns[2]
    camp_col = cdr.columns[3]

    mature_mask = cdr[disp_col].str.contains("mature|transfer", case=False, na=False)
    total_mature = cdr[mature_mask].groupby(emp3).size()

    ib_mask = mature_mask & cdr[camp_col].str.contains("csr", case=False, na=False)
    ib_mature = cdr[ib_mask].groupby(emp3).size()

    ob_mature = total_mature - ib_mature

    # ---------------- CRM REPORT ----------------
    emp4 = crm.columns[0]
    total_tagging = crm.groupby(emp4).size()

    # ---------------- FINAL MERGE ----------------
    final = pd.DataFrame({
        "EMP ID": total_login.index,
        "Total Login": total_login.values,
        "Total Net Login": net_login.reindex(total_login.index).fillna(0).values,
        "Total Break": total_break.reindex(total_login.index).fillna(0).values,
        "Total Meeting": total_meeting.reindex(total_login.index).fillna(0).values,
        "Total Talk Time": total_talk.reindex(total_login.index).fillna(0).values,
        "Total Mature": total_mature.reindex(total_login.index).fillna(0).values,
        "IB Mature": ib_mature.reindex(total_login.index).fillna(0).values,
        "OB Mature": ob_mature.reindex(total_login.index).fillna(0).values,
        "Total Tagging": total_tagging.reindex(total_login.index).fillna(0).values
    })

    final["AHT"] = (final["Total Talk Time"] / final["Total Mature"]).fillna(0).round(2)

    # Convert seconds to HH:MM:SS
    for col in ["Total Login","Total Net Login","Total Break","Total Meeting","Total Talk Time"]:
        final[col] = pd.to_timedelta(final[col], unit='s')

    final.to_excel(RESULT_EXCEL, index=False)

    create_pdf(final)

    table_html = final.to_html(index=False)

    return THEME + f"""
    <div class='box'>
    <h2>Final Agent Performance Report ✅</h2>
    {table_html}
    <br>
    <a class='btn' href='/download_excel'>⬇ Download Excel</a>
    <a class='btn' href='/download_pdf'>⬇ Download PDF</a>
    </div>
    """


def create_pdf(df):
    c = canvas.Canvas(RESULT_PDF, pagesize=A4)
    w,h = A4
    c.setFont("Helvetica-Bold",10)
    c.drawRightString(w-40, h-40, "Chandan Malakar")

    y = h-80
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
            c.drawRightString(w-40, h-40, "Chandan Malakar")
            y=h-80

    c.save()

@app.route("/download_excel")
def download_excel():
    return send_file(RESULT_EXCEL, as_attachment=True)

@app.route("/download_pdf")
def download_pdf():
    return send_file(RESULT_PDF, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
