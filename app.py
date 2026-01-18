from flask import Flask, render_template, request, jsonify
import pandas as pd

app = Flask(__name__)

@app.route("/")
def upload():
    return render_template("upload.html")

@app.route("/dashboard")
def dashboard():
    return render_template("dashboard.html")

@app.route("/process", methods=["POST"])
def process():

    if "agent" not in request.files or "cdr" not in request.files:
        return jsonify({"error":"Files missing"}),400

    agent = pd.read_excel(request.files["agent"])
    cdr   = pd.read_excel(request.files["cdr"])

    # ----- SIMPLE TEST DATA (confirm pipeline works) -----
    data = [
        {
            "Agent Name":"160250",
            "Agent Full Name":"Arti Vishwakarma",
            "Total Login Time":"06:04:11",
            "Total Net Login":"01:16:16",
            "Total Break":"00:46:39",
            "Total Meeting":"00:34:43",
            "AHT":"00:02:50",
            "Total Call":149,
            "IB Mature":48,
            "OB Mature":60
        }
    ]

    return jsonify(data)
