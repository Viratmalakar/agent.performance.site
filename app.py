from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

# ================= HELPERS =================
def to_sec(t):
    try:
        h, m, s = map(int, str(t).split(":"))
        return h * 3600 + m * 60 + s
    except:
        return 0

def to_time(sec):
    sec = int(sec)
    return f"{sec//3600:02}:{(sec%3600)//60:02}:{sec%60:02}"

# ================= ROUTES =================
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():

    agent_files = request.files.getlist("agent_files")
    cdr_files = request.files.getlist("cdr_files")

    # ================= AGENT REPORT =================
    frames = []
    for f in agent_files:
        df = pd.read_excel(f).iloc[2:].reset_index(drop=True)
        df = df.replace("-", 0).infer_objects(copy=False)

        frames.append(pd.DataFrame({
            "Employee ID": df.iloc[:, 1],
            "Agent Name": df.iloc[:, 2],
            "Login": df.iloc[:, 3].apply(to_sec),
            "Talk": df.iloc[:, 5].apply(to_sec),
            "ACW": df.iloc[:, 7].apply(to_sec),
            "Lunch": df.iloc[:, 19].apply(to_sec),
            "Meeting": df.iloc[:, 20].apply(to_sec) + df.iloc[:, 23].apply(to_sec),
            "Short": df.iloc[:, 22].apply(to_sec),
            "Tea": df.iloc[:, 24].apply(to_sec),
        }))

    ap = pd.concat(frames).groupby("Employee ID", as_index=False).sum()
    ap["Total Break"] = ap["Lunch"] + ap["Short"] + ap["Tea"]
    ap["Net Login"] = (ap["Login"] - ap["Total Break"]).clip(lower=0)

    # ================= CDR =================
    cdr_frames = []
    for f in cdr_files:
        df = pd.read_excel(f).replace("-", 0)
        cdr_frames.append(pd.DataFrame({
            "Employee ID": df.iloc[:, 1],
            "Campaign": df.iloc[:, 6],
            "Status": df.iloc[:, 25]
        }))

    cdr = pd.concat(cdr_frames)
    cdr = cdr[cdr["Status"].isin(["CALLMATURED", "TRANSFERRED"])]

    ib = cdr[cdr["Campaign"] == "CSRINBOUND"].groupby("Employee ID").size()
    total = cdr.groupby("Employee ID").size()

    ap["IB Calls"] = ap["Employee ID"].map(ib).fillna(0).astype(int)
    ap["Total Calls"] = ap["Employee ID"].map(total).fillna(0).astype(int)
    ap["OB Calls"] = ap["Total Calls"] - ap["IB Calls"]

    # ================= AHT =================
    ap["AHT"] = ap.apply(
        lambda r: to_time(r["Talk"]/r["Total Calls"])
        if r["Total Calls"] > 0 else "00:00:00", axis=1
    )

    # ================= RED FLAGS =================
    ap["Break_Red"] = ap["Total Break"] > 2100
    ap["Meeting_Red"] = ap["Meeting"] > 2100
    ap["Net_Red"] = ap["Net Login"] < (ap["Login"] - 2100)

    # ================= SUMMARY =================
    total_calls = ap["Total Calls"].sum()
    total_talk = ap["Talk"].sum()

    summary = {
        "ivr": int(total_calls),
        "mature": int(total_calls),
        "ib": int(ap["IB Calls"].sum()),
        "ob": int(ap["OB Calls"].sum()),
        "talk": to_time(total_talk),
        "login": to_time(ap["Login"].sum()),
        "aht": to_time(total_talk/total_calls) if total_calls else "00:00:00"
    }

    # ================= FINAL FORMAT =================
    for c in ["Login","Net Login","Talk","ACW","Total Break","Meeting"]:
        ap[c] = ap[c].apply(to_time)

    return render_template(
        "result.html",
        data=ap.to_dict("records"),
        summary=summary
    )

if __name__ == "__main__":
    app.run(port=10000, host="0.0.0.0")
