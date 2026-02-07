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
    try:
        sec = int(sec)
        h = sec // 3600
        m = (sec % 3600) // 60
        s = sec % 60
        return f"{h:02}:{m:02}:{s:02}"
    except:
        return "00:00:00"


# ================= ROUTES =================
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():

    agent_files = request.files.getlist("agent_files")
    cdr_files = request.files.getlist("cdr_files")

    # ======================================================
    # AGENT PERFORMANCE REPORT (AGENT WISE)
    # ======================================================
    agent_frames = []

    for file in agent_files:
        if not file.filename:
            continue

        df = pd.read_excel(file)
        df = df.iloc[2:].reset_index(drop=True)

        # '-' means 0
        df = df.replace("-", 0).infer_objects(copy=False)

        temp = pd.DataFrame({
            "Employee ID": df.iloc[:, 1],                 # B
            "Agent Name": df.iloc[:, 2],                  # C
            "Login": df.iloc[:, 3].apply(to_sec),         # D
            "Talk": df.iloc[:, 5].apply(to_sec),          # F
            "ACW": df.iloc[:, 7].apply(to_sec),           # H
            "Lunch": df.iloc[:, 19].apply(to_sec),        # T
            "Meeting": (
                df.iloc[:, 20].apply(to_sec) +            # U
                df.iloc[:, 23].apply(to_sec)              # X
            ),
            "Short": df.iloc[:, 22].apply(to_sec),        # W
            "Tea": df.iloc[:, 24].apply(to_sec)           # Y
        })

        agent_frames.append(temp)

    ap = pd.concat(agent_frames, ignore_index=True)

    # Agent-wise SUM
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

    # ======================================================
    # CDR REPORT (CALL COUNT – TEXT SAFE)
    # ======================================================
    cdr_frames = []

    for file in cdr_files:
        if not file.filename:
            continue

        cdr = pd.read_excel(file)
        cdr = cdr.replace("-", 0).infer_objects(copy=False)

        temp = pd.DataFrame({
            "Employee ID": cdr.iloc[:, 1],    # B
            "Campaign": cdr.iloc[:, 6],       # G
            "Call Status": cdr.iloc[:, 25]    # Z
        })

        cdr_frames.append(temp)

    cdr = pd.concat(cdr_frames, ignore_index=True)

    # Only matured calls
    cdr = cdr[cdr["Call Status"] == "CALLMATURED"]

    # IB / OB COUNT (ROWS COUNT, NOT SUM)
    ib = (
        cdr[cdr["Campaign"] == "CSRIBOUND"]
        .groupby("Employee ID")
        .size()
    )

    ob = (
        cdr[cdr["Campaign"] != "CSRIBOUND"]
        .groupby("Employee ID")
        .size()
    )

    ap["IB Calls"] = ap["Employee ID"].map(ib).fillna(0).astype(int)
    ap["OB Calls"] = ap["Employee ID"].map(ob).fillna(0).astype(int)
    ap["Total Calls"] = ap["IB Calls"] + ap["OB Calls"]

    # ======================================================
    # AHT CALCULATION
    # ======================================================
    ap["AHT"] = ap.apply(
        lambda r: to_time(r["Talk"] / r["Total Calls"])
        if r["Total Calls"] > 0 else "00:00:00",
        axis=1
    )

    # Convert seconds back to hh:mm:ss
    for col in ["Login", "Net Login", "Talk", "ACW", "Total Break", "Meeting"]:
        ap[col] = ap[col].apply(to_time)

    # Final column order
    ap = ap[[
        "Employee ID",
        "Agent Name",
        "Login",
        "Net Login",
        "Talk",
        "ACW",
        "Total Break",
        "Meeting",
        "IB Calls",
        "OB Calls",
        "Total Calls",
        "AHT"
    ]]

    return render_template(
        "result.html",
        columns=ap.columns.tolist(),
        data=ap.to_dict(orient="records")
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
