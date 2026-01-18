from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
from datetime import datetime
import io

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

    agent = agent.iloc[2:]
    cdr   = cdr.iloc[1:]

    agent.fillna("00:00:00",inplace=True)
    cdr.fillna("",inplace=True)

    agent.iloc[:,1:31] = agent.iloc[:,1:31].replace("-", "00:00:00")

    # AGENT COLUMN MAP
    EMP_COL = agent.columns[1]
    FULLNAME_COL = agent.columns[2]
    D = agent.columns[3]
    F = agent.columns[5]
    T = agent.columns[19]
    W = agent.columns[22]
    Y = agent.columns[24]
    U = agent.columns[20]
    X = agent.columns[23]

    # CDR COLUMN MAP (BY NAME)
    USER_COL = "Username"
    CAMP_COL = "Campaign"
    DISP_COL = "Disposition"
    DISP_TYPE_COL = "DispositionType"

    for col in [USER_COL, CAMP_COL, DISP_COL, DISP_TYPE_COL]:
        if col not in cdr.columns:
            return jsonify({"error":f"Missing column in CDR: {col}"})

    agent_groups = dict(tuple(agent.groupby(EMP_COL)))
    cdr_groups = dict(tuple(cdr.groupby(USER_COL)))

    final = []
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

        mature = c[c[DISP_TYPE_COL].str.contains("callmature|transfer",case=False,na=False)]
        total_mature = len(mature)

        ib = mature[mature[CAMP_COL].str.contains("CSRINBOUND",case=False,na=False)]
        ib_mature = len(ib)
        ob_mature = total_mature - ib_mature

        ivr_hit = len(c[c[CAMP_COL].str.contains("CSRINBOUND",case=False,na=False)])

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

    columns = [
        "Employee ID","Full Name",
        "Total Login","Net Login","Total Break","Total Meeting",
        "Total Talk Time","AHT",
        "Total Mature","IB Mature","OB Mature","Total IVR Hit"
    ]

    output=[]
    for row in final:
        output.append(dict(zip(columns,row)))

    return jsonify(output)

@app.route("/export",methods=["POST"])
def export():
    data = request.get_json()
    df = pd.DataFrame(data)

    output = io.BytesIO()
    df.to_excel(output,index=False)
    output.seek(0)

    filename = "Agent_Report_"+datetime.now().strftime("%Y-%m-%d_%H-%M-%S")+".xlsx"

    return send_file(output,as_attachment=True,download_name=filename,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(debug=True)
