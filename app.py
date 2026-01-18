from flask import Flask, render_template, request, jsonify
import pandas as pd
from datetime import datetime

app = Flask(__name__)

def fix_time(x):
    if pd.isna(x) or x=="-" or x=="nan":
        return "00:00:00"
    return str(x)

def time_to_sec(t):
    try:
        h,m,s = t.split(":")
        return int(h)*3600+int(m)*60+int(s)
    except:
        return 0

def sec_to_time(s):
    h=s//3600
    m=(s%3600)//60
    sec=s%60
    return f"{h:02}:{m:02}:{sec:02}"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process",methods=["POST"])
def process():

    agent = pd.read_excel(request.files["agent"],dtype=str)
    cdr   = pd.read_excel(request.files["cdr"],dtype=str)

    # ---------- CLEAN ----------
    agent = agent.iloc[2:]     # remove top 2 rows
    cdr   = cdr.iloc[1:]       # remove top 1 row

    agent.fillna("00:00:00",inplace=True)
    cdr.fillna("",inplace=True)

    agent.iloc[:,1:31] = agent.iloc[:,1:31].replace("-", "00:00:00")

    # Column Mapping (as per your logic)
    EMP_COL = agent.columns[1]      # B column
    FULLNAME_COL = agent.columns[2]

    D = agent.columns[3]
    F = agent.columns[5]

    T = agent.columns[19]
    W = agent.columns[22]
    Y = agent.columns[24]

    U = agent.columns[20]
    X = agent.columns[23]

    DISP_COL = cdr.columns[25]      # Z
    CAMP_COL = cdr.columns[6]       # G
    USER_COL = cdr.columns[2]       # Username

    final = []

    # Speed optimization
    agent_groups = dict(tuple(agent.groupby(EMP_COL)))
    cdr_groups = dict(tuple(cdr.groupby(USER_COL)))

    grand_total_mature = 0
    grand_total_ivr = 0
    grand_talk_sec = 0

    for emp in agent_groups:

        a = agent_groups[emp]
        c = cdr_groups.get(emp,pd.DataFrame())

        fullname = a[FULLNAME_COL].iloc[0]

        total_login = sum(a[D].apply(lambda x: time_to_sec(fix_time(x))))
        total_break = sum(a[T].apply(lambda x: time_to_sec(fix_time(x)))) + \
                      sum(a[W].apply(lambda x: time_to_sec(fix_time(x)))) + \
                      sum(a[Y].apply(lambda x: time_to_sec(fix_time(x))))

        total_meeting = sum(a[U].apply(lambda x: time_to_sec(fix_time(x)))) + \
                        sum(a[X].apply(lambda x: time_to_sec(fix_time(x))))

        net_login = total_login - total_break

        talk_sec = sum(a[F].apply(lambda x: time_to_sec(fix_time(x))))

        # -------- Mature Logic ----------
        mature = c[c[DISP_COL].str.contains("callmature|transfer",case=False,na=False)]
        total_mature = len(mature)

        ib = mature[mature[CAMP_COL].str.contains("CSRINBOUND",case=False,na=False)]
        ib_mature = len(ib)

        ob_mature = total_mature - ib_mature

        # -------- IVR HIT ----------
        ivr_hit = len(c[c[CAMP_COL].str.contains("CSRINBOUND",case=False,na=False)])

        # -------- AHT ----------
        if total_mature>0:
            aht = sec_to_time(int(talk_sec/total_mature))
        else:
            aht = "00:00:00"

        grand_total_mature += total_mature
        grand_total_ivr += ivr_hit
        grand_talk_sec += talk_sec

        final.append([
            emp, fullname,
            sec_to_time(total_login),
            sec_to_time(net_login),
            sec_to_time(total_break),
            sec_to_time(total_meeting),
            sec_to_time(talk_sec),
            aht,
            total_mature,
            ib_mature,
            ob_mature,
            ivr_hit
        ])

    # -------- GRAND TOTAL ROW ----------
    if grand_total_mature>0:
        grand_aht = sec_to_time(int(grand_talk_sec/grand_total_mature))
    else:
        grand_aht="00:00:00"

    final.append([
        "GRAND TOTAL","",
        "","","","",
        sec_to_time(grand_talk_sec),
        grand_aht,
        grand_total_mature,
        "",
        "",
        grand_total_ivr
    ])

    return jsonify(final)
