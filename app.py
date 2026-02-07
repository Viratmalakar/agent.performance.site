from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

# ---------- helpers ----------
def to_sec(t):
    try:
        h, m, s = map(int, str(t).split(":"))
        return h * 3600 + m * 60 + s
    except:
        return 0


def to_time(sec):
    try:
        sec = int(sec)
        h = sec // 3600
        m = (sec % 3600) // 60
        s = sec % 60
        return f"{h:02}:{m:02}:{s:02}"
    except:
        return "00:00:00"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():

    agent_files = request.files.getlist("agent_files")
    cdr_files = request.files.getlist("cdr_files")

    # ================= AGENT PERFORMANCE =================
    agent_frames = []

    for file in agent_files:
        if not file.filename:
            continue

        df = pd.read_excel(file)
        df = df.iloc[2:].reset_index(drop=True)
        df = df.replace("-", 0).infer_objects(copy=False)

        # Column mapping (LETTER BASED)
        emp_id   = df.iloc[:, 1]          # B
        name     = df.iloc[:, 2]          # C
        login    = df.iloc[:, 3].apply(to_sec)   # D
        talk     = df.iloc[:, 5].apply(to_sec)   # F
        acw      = df.iloc[:, 7].apply(to_sec)   # H
        lunch    = df.iloc[:, 19].apply(to_sec)  # T
        meeting  = df.iloc[:, 20].apply(to_sec)  # U
        shortb   = df.iloc[:, 22].apply(to_sec)  # W
        sysdown  = df.iloc[:, 23].apply(to_sec)  # X
        tea      = df.iloc[:, 24].apply(to_sec)  # Y

        temp = pd.DataFrame({
            "Employee ID": emp_id,
            "Agent Name": name,
            "Login": login,
            "Talk": talk,
            "ACW": acw,
            "Lunch": lunch,
            "Short": shortb,
            "Tea": tea,
            "Meeting": meeting + sysdown
        })

        agent_frames.append(temp)

    ap = pd.concat(agent_frames, ignore_index=True)

    ap = ap.groupby("Employee ID", as_index=False).agg({
        "Agent Name": "first",
        "Login": "sum",
        "Talk": "sum",
        "ACW": "sum",
        "Lunch": "sum",
        "Short": "sum",
        "Tea": "sum",
        "Meeting": "sum"
    })

    ap["Total Break"] = ap["Lunch"] + ap["Short"] + ap["Tea"]
    ap["Net Login"] = ap["Login"] - ap["Total Break"]

    # ================= CDR =================
    cdr_frames = []

    for file in cdr_files:
        if not file.filename:
            continue

        cdr = pd.read_excel(file)
        cdr = cdr.replace("-", 0).infer_objects(copy=False)

        emp = cdr.iloc[:, 1]   # B
        camp = cdr.iloc[:, 6]  # G
        calls = cdr.iloc[:, 25]  # Z

        temp = pd.DataFrame({
            "Employee ID": emp,
            "Campaign": camp,
            "Calls": calls
        })

        cdr_frames.append(temp)

    cdr = pd.concat(cdr_frames, ignore_index=True)

    ib = cdr[cdr["Campaign"] == "CSRIBOUND"].groupby("Employee ID")["Calls"].sum()
    ob = cdr[cdr["Campaign"] != "CSRIBOUND"].groupby("Employee ID")["Calls"].sum()

    ap["IB Calls"] = ap["Employee ID"].map(ib).fillna(0).astype(int)
    ap["OB Calls"] = ap["Employee ID"].map(ob).fillna(0).astype(int)
    ap["Total Calls"] = ap["IB Calls"] + ap["OB Calls"]

    ap["AHT"] = ap.apply(
        lambda r: to_time(r["Talk"] / r["Total Calls"]) if r["Total Calls"] > 0 else "00:00:00",
        axis=1
    )

    # Convert time back
    for c in ["Login", "Net Login", "Talk", "ACW", "Total Break", "Meeting"]:
        ap[c] = ap[c].apply(to_time)

    final_cols = [
        "Employee ID", "Agent Name", "Login", "Net Login",
        "Talk", "ACW", "Total Break", "Meeting",
        "IB Calls", "OB Calls", "Total Calls", "AHT"
    ]

    ap = ap[final_cols]

    return render_template(
        "result.html",
        columns=ap.columns.tolist(),
        data=ap.to_dict(orient="records")
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
