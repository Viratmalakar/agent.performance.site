from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    agent_files = request.files.getlist("agent_files")
    cdr_files   = request.files.getlist("cdr_files")

    agent_data = []
    cdr_data = []

    # ---------- Agent Performance ----------
    for file in agent_files:
        if file.filename == "":
            continue

        df = pd.read_excel(file)
        df = df.iloc[2:].reset_index(drop=True)
        df.replace("-", 0, inplace=True)

        def to_sec(t):
            try:
                h, m, s = map(int, str(t).split(":"))
                return h * 3600 + m * 60 + s
            except:
                return 0

        def to_time(sec):
            h = sec // 3600
            m = (sec % 3600) // 60
            s = sec % 60
            return f"{h:02}:{m:02}:{s:02}"

        cols = [
            "Total Login",
            "Lunch Break",
            "Tea Break",
            "Short Break",
            "Meeting",
            "System Down"
        ]

        for c in cols:
            if c in df.columns:
                df[c] = df[c].apply(to_sec)

        df["Total Break"] = df["Lunch Break"] + df["Tea Break"] + df["Short Break"]
        df["Total Meeting"] = df["Meeting"] + df["System Down"]
        df["Net Login"] = df["Total Login"] - df["Total Break"]

        for c in ["Total Break", "Total Meeting", "Net Login"]:
            df[c] = df[c].apply(to_time)

        agent_data.append(df)

    # ---------- CDR ----------
    for file in cdr_files:
        if file.filename == "":
            continue
        cdr_df = pd.read_excel(file)
        cdr_data.append(cdr_df)

    final_agent_df = pd.concat(agent_data, ignore_index=True) if agent_data else pd.DataFrame()
    final_cdr_df   = pd.concat(cdr_data, ignore_index=True) if cdr_data else pd.DataFrame()

    # For now showing Agent Performance
    return render_template(
        "result.html",
        columns=final_agent_df.columns.tolist(),
        data=final_agent_df.fillna("").to_dict(orient="records"),
        cdr_rows=len(final_cdr_df)
    )


if __name__ == "__main__":
    app.run(debug=True)
