from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


# ---------- Time conversion ----------
def time_to_seconds(t):
    try:
        if pd.isna(t):
            return 0
        h, m, s = map(int, str(t).split(":"))
        return h * 3600 + m * 60 + s
    except:
        return 0


def seconds_to_time(sec):
    try:
        h = sec // 3600
        m = (sec % 3600) // 60
        s = sec % 60
        return f"{h:02}:{m:02}:{s:02}"
    except:
        return "00:00:00"


# ---------- Routes ----------
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    file = request.files.get("file")

    if not file:
        return "No file uploaded"

    filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(filepath)

    # ---------- Read Excel ----------
    df = pd.read_excel(filepath)

    # ---------- Clean data ----------
    df = df.iloc[2:].reset_index(drop=True)
    df.replace("-", 0, inplace=True)

    time_cols = [
        "Total Login",
        "Lunch Break",
        "Tea Break",
        "Short Break",
        "Meeting",
        "System Down"
    ]

    for col in time_cols:
        if col in df.columns:
            df[col] = df[col].apply(time_to_seconds)

    # ---------- Calculations ----------
    if set(time_cols).issubset(df.columns):
        df["Total Break"] = (
            df["Lunch Break"] +
            df["Tea Break"] +
            df["Short Break"]
        )

        df["Total Meeting"] = (
            df["Meeting"] +
            df["System Down"]
        )

        df["Net Login"] = (
            df["Total Login"] - df["Total Break"]
        )

    # ---------- Convert back to time ----------
    convert_cols = ["Total Break", "Total Meeting", "Net Login"]

    for col in convert_cols:
        if col in df.columns:
            df[col] = df[col].apply(seconds_to_time)

    # ---------- Send to UI ----------
    table_data = df.fillna("").to_dict(orient="records")
    columns = df.columns.tolist()

    return render_template(
        "dashboard.html",
        columns=columns,
        data=table_data
    )


if __name__ == "__main__":
    app.run(debug=True)
