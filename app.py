from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


# ---------- Time helpers ----------
def to_sec(t):
    try:
        h, m, s = map(int, str(t).split(":"))
        return h * 3600 + m * 60 + s
    except:
        return 0


def to_time(sec):
    try:
        h = sec // 3600
        m = (sec % 3600) // 60
        s = sec % 60
        return f"{h:02}:{m:02}:{s:02}"
    except:
        return "00:00:00"


@app.route("/process", methods=["POST"])
def process():

    agent_files = request.files.getlist("agent_files")
    cdr_files   = request.files.getlist("cdr_files")  # future use

    frames = []

    for file in agent_files:
        if file.filename == "":
            continue

        # ---- Read Excel ----
        df = pd.read_excel(file)

        # Remove top 2 junk rows
        df = df.iloc[2:].reset_index(drop=True)

        # Replace '-' safely
        df = df.replace("-", 0).infer_objects(copy=False)

        # ==================================================
        # 🔴 FIXED COLUMN MAPPING (EXCEL LETTER → INDEX)
        # ==================================================
        # A=0, B=1, C=2 ...

        total_login = df.iloc[:, 3].apply(to_sec)    # D:D
        lunch       = df.iloc[:, 19].apply(to_sec)   # T:T  (LUNCHBREAK)
        tea         = df.iloc[:, 22].apply(to_sec)   # W:W
        shortb      = df.iloc[:, 24].apply(to_sec)   # Y:Y
        meeting     = df.iloc[:, 20].apply(to_sec)   # U:U
        sysdown     = df.iloc[:, 23].apply(to_sec)   # X:X

        # ---------- Calculations ----------
        df["Total Break"]   = lunch + tea + shortb
        df["Total Meeting"] = meeting + sysdown
        df["Net Login"]     = total_login - df["Total Break"]

        # ---------- Convert back to hh:mm:ss ----------
        for c in ["Total Break", "Total Meeting", "Net Login"]:
            df[c] = df[c].apply(to_time)

        frames.append(df)

    final_df = (
        pd.concat(frames, ignore_index=True)
        if frames else pd.DataFrame()
    )

    return render_template(
        "result.html",
        columns=final_df.columns.tolist(),
        data=final_df.fillna("").to_dict(orient="records")
    )


if __name__ == "__main__":
    app.run(debug=True)
