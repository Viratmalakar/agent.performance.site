from flask import Flask, render_template, request, redirect, session, send_file
from users import USERS
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = "agent_secret_key"

FINAL_DF = None

def score_cols(cols, keywords):
    score = 0
    for k in keywords:
        if k in cols:
            score += 1
    return score

def find_col(df, keys):
    for c in df.columns:
        for k in keys:
            if k in c.lower():
                return c
    return None

def time_to_sec(t):
    try:
        return pd.to_timedelta(t).total_seconds()
    except:
        return 0

def sec_to_time(s):
    return str(pd.to_timedelta(int(s), unit="s"))

@app.route("/", methods=["GET","POST"])
def login():
    if request.method=="POST":
        u=request.form["username"]
        p=request.form["password"]
        if u in USERS and USERS[u]==p:
            session["user"]=u
            return redirect("/dashboard")
        return render_template("login.html", error="Invalid login")
    return render_template("login.html")

@app.route("/dashboard")
def dashboard():
    return render_template("index.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

@app.route("/upload")
def upload():
    return render_template("upload.html")

@app.route("/process", methods=["POST"])
def process():
    global FINAL_DF

    files = request.files.getlist("files")
    dfs = []

    for f in files:
        dfs.append(pd.read_excel(f))

    login_df = cdr_df = agent_df = crm_df = None

    for df in dfs:
        cols = " ".join(df.columns.astype(str)).lower()

        login_score = score_cols(cols, ["login","logout","break","lunch","tea"])
        cdr_score   = score_cols(cols, ["disposition","campaign","cdr"])
        agent_score = score_cols(cols, ["talk","agent","aht"])
        crm_score   = score_cols(cols, ["createdby","tag","crm","detail"])

        scores = {
            "login": login_score,
            "cdr": cdr_score,
            "agent": agent_score,
            "crm": crm_score
        }

        best = max(scores, key=scores.get)

        if best=="login" and login_df is None:
            login_df=df
        elif best=="cdr" and cdr_df is None:
            cdr_df=df
        elif best=="agent" and agent_df is None:
            agent_df=df
        elif best=="crm" and crm_df is None:
            crm_df=df

    if not all([login_df is not None, cdr_df is not None, agent_df is not None, crm_df is not None]):
        return "❌ System still could not map all reports. Please share screenshot of Excel columns."

    emp_col = find_col(login_df,["user","agent","emp"])
    login_col = find_col(login_df,["login"])
    lunch_col = find_col(login_df,["lunch"])
    short_col = find_col(login_df,["short"])
    tea_col = find_col(login_df,["tea"])
    meet_col = find_col(login_df,["meeting"])
    sys_col = find_col(login_df,["system"])

    talk_col = find_col(agent_df,["talk"])
    agent_name_col = find_col(agent_df,["agent"])

    disp_col = find_col(cdr_df,["disposition"])
    camp_col = find_col(cdr_df,["campaign"])
    cdr_agent_col = find_col(cdr_df,["agent","user"])

    crm_agent_col = find_col(crm_df,["createdby"])

    login_df["LoginSec"] = login_df[login_col].apply(time_to_sec)
    login_df["BreakSec"] = (
        login_df[lunch_col].apply(time_to_sec) +
        login_df[short_col].apply(time_to_sec) +
        login_df[tea_col].apply(time_to_sec)
    )
    login_df["MeetingSec"] = (
        login_df[meet_col].apply(time_to_sec) +
        login_df[sys_col].apply(time_to_sec)
    )

    login_sum = login_df.groupby(emp_col).sum().reset_index()

    agent_df["TalkSec"] = agent_df[talk_col].apply(time_to_sec)
    talk_sum = agent_df.groupby(agent_name_col)["TalkSec"].sum().reset_index()

    mature = cdr_df[cdr_df[disp_col].str.contains("mature|transfer",case=False,na=False)]
    total_mature = mature.groupby(cdr_agent_col).size().reset_index(name="Total Mature")

    inbound = mature[mature[camp_col].str.contains("inbound",case=False,na=False)]
    ib_mature = inbound.groupby(cdr_agent_col).size().reset_index(name="IB Mature")

    crm_count = crm_df.groupby(crm_agent_col).size().reset_index(name="Total Tagging")

    final = login_sum.merge(talk_sum,left_on=emp_col,right_on=agent_name_col,how="left")
    final = final.merge(total_mature,left_on=emp_col,right_on=cdr_agent_col,how="left")
    final = final.merge(ib_mature,left_on=emp_col,right_on=cdr_agent_col,how="left")
    final = final.merge(crm_count,left_on=emp_col,right_on=crm_agent_col,how="left")

    final.fillna(0,inplace=True)

    final["Total Login"] = final["LoginSec"].apply(sec_to_time)
    final["Total Break"] = final["BreakSec"].apply(sec_to_time)
    final["Total Meeting"] = final["MeetingSec"].apply(sec_to_time)
    final["Total Talk Time"] = final["TalkSec"].apply(sec_to_time)
    final["OB Mature"] = final["Total Mature"] - final["IB Mature"]
    final["AHT"] = (final["TalkSec"]/final["Total Mature"]).fillna(0).round(2)

    FINAL_DF = final[[emp_col,
        "Total Login","Total Break","Total Meeting",
        "Total Talk Time","Total Mature","IB Mature",
        "OB Mature","Total Tagging","AHT"]]

    FINAL_DF.columns = [
        "EMP ID","Total Login","Total Break","Total Meeting",
        "Total Talk Time","Total Mature","IB Mature",
        "OB Mature","Total Tagging","AHT"
    ]

    return render_template("result.html",
        tables=[FINAL_DF.to_html(classes="table",index=False)])

@app.route("/download")
def download():
    FINAL_DF.to_excel("Final_Agent_Performance.xlsx",index=False)
    return send_file("Final_Agent_Performance.xlsx",as_attachment=True)

if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0",port=port)
