from flask import Flask, render_template, request, redirect, url_for, session, send_file
import pandas as pd
import io

app = Flask(__name__)
app.secret_key = "agentdashboard"

# ---------- TIME HELPERS ----------
def tsec(t):
    try:
        h, m, s = map(int, str(t).split(":"))
        return h * 3600 + m * 60 + s
    except:
        return 0

def stime(sec):
    sec = int(max(0, sec))
    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60
    return f"{h:02d}:{m:02d}:{s:02d}"

# ---------- ROUTES ----------
@app.route("/")
def upload():
    return render_template("upload.html")

@app.route("/process", methods=["POST"])
def process():

    agent = pd.read_excel(request.files["agent"], engine="openpyxl")
    cdr   = pd.read_excel(request.files["cdr"], engine="openpyxl")

    agent = agent.fillna("00:00:00").replace("-", "00:00:00")

    # ---- AGENT COLUMNS ----
    emp   = agent.columns[1]
    full  = agent.columns[2]
    login = agent.columns[3]
    talk  = agent.columns[5]

    t = agent.columns[19]
    u = agent.columns[20]
    w = agent.columns[22]
    x = agent.columns[23]
    y = agent.columns[24]

    # ---- CDR COLUMNS ----
    c_emp = cdr.columns[1]
    camp  = cdr.columns[6]
    disp  = cdr.columns[25]

    cdr[camp] = cdr[camp].astype(str)
    cdr[disp] = cdr[disp].astype(str)

    # ---- IVR HIT ----
    ivr_hit = cdr[cdr[camp].str.upper() == "CSRINBOUND"].shape[0]

    # ---- MATURE CALLS ----
    mature = cdr[cdr[disp].str.contains("callmature|transfer", case=False, na=False)]
    ib     = mature[mature[camp].str.upper() == "CSRINBOUND"]

    mature_cnt = mature.groupby(c_emp).size()
    ib_cnt     = ib.groupby(c_emp).size()

    mature_cnt.index = mature_cnt.index.astype(str).str.strip()
    ib_cnt.index     = ib_cnt.index.astype(str).str.strip()

    # ---- FINAL TABLE ----
    final = pd.DataFrame()
    final["Agent Name"] = agent[emp].astype(str).str.strip()
    final["Agent Full Name"] = agent[full]
    final["Total Login Time"] = agent[login]

    # ---- FAST TIME CALC ----
    break_sec = agent[[t, w, y]].apply(lambda col: col.map(tsec)).sum(axis=1)
    meet_sec  = agent[[u, x]].apply(lambda col: col.map(tsec)).sum(axis=1)
    login_sec = agent[login].map(tsec)
    talk_sec  = agent[talk].map(tsec)

    final["Total Break"] = break_sec.map(stime)
    final["Total Meeting"] = meet_sec.map(stime)
    final["Total Net Login"] = (login_sec - break_sec).map(stime)

    # ---- CALL COUNTS ----
    final["Total Call"] = final["Agent Name"].map(mature_cnt).fillna(0).astype(int)
    final["IB Mature"] = final["Agent Name"].map(ib_cnt).fillna(0).astype(int)
    final["OB Mature"] = final["Total Call"] - final["IB Mature"]

    # ---- ✅ AHT FIX (TOTAL TALK TIME / TOTAL CALLS) ----
    call_cnt = final["Total Call"]
    final["AHT"] = [
        stime(t / c) if c > 0 else "00:00:00"
        for t, c in zip(talk_sec, call_cnt)
    ]

    # ---- REMOVE BAD ROWS ----
    final = final[final["Agent Name"].notna()]
    final = final[~final["Agent Name"].str.lower().isin(
        ["agent name", "emp id", "nan", "none", "0"]
    )]
    final = final[~(
        (final["Total Login Time"] == "00:00:00") &
        (final["Total Net Login"] == "00:00:00") &
        (final["Total Break"] == "00:00:00") &
        (final["Total Meeting"] == "00:00:00")
    )]

    # ---- HIGHLIGHT FLAGS ----
    final["__red_net"]   = (login_sec > (8*3600 + 15*60)) & ((login_sec - break_sec) < 8*3600)
    final["__red_break"] = break_sec > 35*60
    final["__red_meet"]  = meet_sec  > 35*60
    final["__red_net"] = (net_login_sec >= 8*3600) & (final["Total Call"] < 100)

    # ---- GRAND TOTAL ----
    gt = {
        "TOTAL IVR HIT": int(ivr_hit),
        "TOTAL MATURE": int(final["Total Call"].sum()),
        "IB MATURE": int(final["IB Mature"].sum()),
        "OB MATURE": int(final["OB Mature"].sum()),
        "TOTAL TALK TIME": stime(talk_sec.sum()),
        "AHT": stime(int(talk_sec.sum() / max(1, final["Total Call"].sum()))),
        "LOGIN COUNT": int(len(final))
    }

    session["data"] = final.to_dict(orient="records")
    session["gt"] = gt

    return redirect(url_for("result"))

@app.route("/result")
def result():
    if "data" not in session or "gt" not in session:
        return redirect(url_for("upload"))
    return render_template("result.html", data=session["data"], gt=session["gt"])

@app.route("/export")
def export():

    if "data" not in session:
        return redirect(url_for("upload"))

    df = pd.DataFrame(session["data"])

    df = df[[
        "Agent Name","Agent Full Name","Total Login Time","Total Net Login",
        "Total Break","Total Meeting","AHT",
        "Total Call","IB Mature","OB Mature"
    ]]

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Dashboard")

        wb = writer.book
        ws = writer.sheets["Dashboard"]

        header = wb.add_format({
            "bold": True, "bg_color": "#065f46",
            "color": "white", "border": 1, "align": "center"
        })
        cell = wb.add_format({"border": 1, "align": "center"})

        for col in range(len(df.columns)):
            ws.write(0, col, df.columns[col], header)
            ws.set_column(col, col, 22)

        for r in range(1, len(df) + 1):
            for c in range(len(df.columns)):
                ws.write(r, c, df.iloc[r-1, c], cell)

    out.seek(0)
    return send_file(out, as_attachment=True,
                     download_name="Agent_Performance_Dashboard.xlsx")


