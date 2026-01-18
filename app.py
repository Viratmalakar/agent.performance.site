from flask import Flask, render_template, request, redirect, url_for, session, send_file
import pandas as pd
import io
from datetime import datetime

app = Flask(__name__)
app.secret_key = "agentdashboard"


# ------------------ TIME FUNCTIONS ------------------

def fix_time(x):
    if pd.isna(x) or str(x).strip().lower() in ["-","nan",""]:
        return "00:00:00"
    return str(x)

def tsec(t):
    try:
        h,m,s = map(int,str(t).split(":"))
        return h*3600+m*60+s
    except:
        return 0

def stime(s):
    h=s//3600
    m=(s%3600)//60
    s=s%60
    return f"{h:02d}:{m:02d}:{s:02d}"


# ------------------ ROUTES ------------------

@app.route("/")
def upload():
    return render_template("upload.html")


@app.route("/process",methods=["POST"])
def process():

    agent = pd.read_excel(request.files["agent"],engine="openpyxl")
    cdr   = pd.read_excel(request.files["cdr"],engine="openpyxl")

    # clean time columns
    agent.iloc[:,1:31] = agent.iloc[:,1:31].astype(str).replace("-", "00:00:00")

    emp   = agent.columns[1]
    full  = agent.columns[2]
    login = agent.columns[3]
    talk  = agent.columns[5]

    t = agent.columns[19]
    u = agent.columns[20]
    w = agent.columns[22]
    x = agent.columns[23]
    y = agent.columns[24]

    c_emp = cdr.columns[1]
    camp  = cdr.columns[6]
    disp  = cdr.columns[25]

    cdr[disp] = cdr[disp].astype(str)
    cdr[camp] = cdr[camp].astype(str)

    # Mature calls
    mature = cdr[cdr[disp].str.contains("callmature|transfer",case=False,na=False)]
    ib     = mature[mature[camp].str.upper()=="CSRINBOUND"]

    mature_cnt = mature.groupby(c_emp).size()
    ib_cnt     = ib.groupby(c_emp).size()

    final = pd.DataFrame()

    final["Agent Name"]      = agent[emp]
    final["Agent Full Name"] = agent[full]

    final["Total Login Time"] = agent[login].apply(fix_time)

    final["Total Break"] = (agent[t].apply(tsec) + agent[w].apply(tsec) + agent[y].apply(tsec)).apply(stime)
    final["Total Meeting"] = (agent[u].apply(tsec) + agent[x].apply(tsec)).apply(stime)

    final["Total Net Login"] = (agent[login].apply(tsec) - final["Total Break"].apply(tsec)).apply(stime)

    final["Total Talk Time"] = agent[talk].apply(fix_time)

    final["Total Mature"] = final["Agent Name"].map(mature_cnt).fillna(0).astype(int)
    final["IB Mature"]    = final["Agent Name"].map(ib_cnt).fillna(0).astype(int)
    final["OB Mature"]    = final["Total Mature"] - final["IB Mature"]

    final["AHT"] = (
        final["Total Talk Time"].apply(tsec) /
        final["Total Mature"].replace(0,1)
    ).astype(int).apply(stime)

    # remove bad rows
    final = final[~final["Agent Name"].astype(str).str.lower().isin(["nan","agent name"])]
    final = final.dropna(subset=["Agent Name"])

    # ------------------ GRAND TOTAL ------------------

    gt = {}

    total_talk_sec = final["Total Talk Time"].apply(tsec).sum()
    total_mature   = int(final["Total Mature"].sum())

    gt["Total IVR Hit"] = int(cdr[cdr[camp].str.upper()=="CSRINBOUND"].shape[0])
    gt["Total Mature"] = total_mature
    gt["IB Mature"]    = int(final["IB Mature"].sum())
    gt["OB Mature"]    = int(final["OB Mature"].sum())
    gt["Total Talk Time"] = stime(total_talk_sec)
    gt["AHT"] = stime(int(total_talk_sec / max(1,total_mature)))

    session["data"] = final.to_dict(orient="records")
    session["gt"]   = gt

    return redirect(url_for("result"))


@app.route("/result")
def result():
    return render_template(
        "result.html",
        data=session.get("data",[]),
        gt=session.get("gt",{})
    )


# ------------------ EXPORT ------------------

@app.route("/export")
def export():

    data = pd.DataFrame(session.get("data",[]))

    out = io.BytesIO()

    with pd.ExcelWriter(out,engine="openpyxl") as writer:
        data.to_excel(writer,index=False,sheet_name="Report")
        ws = writer.sheets["Report"]

        from openpyxl.styles import PatternFill,Font,Border,Side

        header_fill = PatternFill(start_color="1FA463",end_color="1FA463",fill_type="solid")
        header_font = Font(color="FFFFFF",bold=True)
        border = Border(left=Side(style="thin"),right=Side(style="thin"),
                        top=Side(style="thin"),bottom=Side(style="thin"))

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = border

        ws.freeze_panes = "A2"

    out.seek(0)

    now = datetime.now().strftime("%d-%m-%y %H-%M-%S")
    fname = f"Agent_Performance_Report_Chandan-Malakar & {now}.xlsx"

    return send_file(out,download_name=fname,as_attachment=True)


# ------------------ RUN ------------------

if __name__ == "__main__":
    app.run(debug=True)
