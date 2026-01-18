from flask import Flask,render_template,request,jsonify,send_file
import pandas as pd
import io,datetime

app = Flask(__name__)

def fix_time(x):
    if pd.isna(x) or x=="-" or str(x).strip()=="":
        return "00:00:00"
    return str(x)

def time_to_sec(t):
    h,m,s = map(int,t.split(":"))
    return h*3600+m*60+s

def sec_to_time(sec):
    h=sec//3600
    m=(sec%3600)//60
    s=sec%60
    return f"{h:02d}:{m:02d}:{s:02d}"

def divide_time(t,count):
    if count==0: return "00:00:00"
    return sec_to_time(time_to_sec(t)//count)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/process",methods=["POST"])
def process():

    agent = pd.read_excel(request.files["agent"])
    cdr   = pd.read_excel(request.files["cdr"])

    agent = agent.fillna("")

    # "-" replace
    agent.iloc[:,1:31] = agent.iloc[:,1:31].replace("-","00:00:00")

    agent["EMP"]=agent.iloc[:,1].astype(str)

    # CDR cols
    cdr["EMP"]=cdr["Username"].astype(str)
    cdr["Campaign"]=cdr.iloc[:,6].astype(str)
    cdr["Disposition"]=cdr.iloc[:,25].astype(str)

    final=[]

    for emp,grp in agent.groupby("EMP"):

        row={}
        row["Agent Name"]=grp.iloc[0,1]
        row["Agent Full Name"]=grp.iloc[0,2]

        login=fix_time(grp.iloc[0,3])
        talk=fix_time(grp.iloc[0,5])

        break_t = fix_time(grp.iloc[0,19])
        break_w = fix_time(grp.iloc[0,22])
        break_y = fix_time(grp.iloc[0,24])

        meeting_u = fix_time(grp.iloc[0,20])
        meeting_x = fix_time(grp.iloc[0,23])

        total_break = sec_to_time(time_to_sec(break_t)+time_to_sec(break_w)+time_to_sec(break_y))
        total_meet  = sec_to_time(time_to_sec(meeting_u)+time_to_sec(meeting_x))
        net_login   = sec_to_time(time_to_sec(login)-time_to_sec(total_break))

        sub = cdr[cdr["EMP"]==emp]

        mature=sub[sub["Disposition"].str.contains("callmatured|transfer",case=False,na=False)]

        ib=mature[mature["Campaign"].str.upper()=="CSRINBOUND"]
        ob=len(mature)-len(ib)

        aht=divide_time(talk,len(mature))

        row.update({
            "Total Login":login,
            "Total Net Login":net_login,
            "Total Break":total_break,
            "Total Meeting":total_meet,
            "Total Talk Time":talk,
            "AHT":aht,
            "Total Mature":len(mature),
            "IB Mature":len(ib),
            "OB Mature":ob
        })

        final.append(row)

    df=pd.DataFrame(final)

    # GRAND TOTAL
    total_ivr = cdr[cdr["Campaign"].str.upper()=="CSRINBOUND"].shape[0]

    total_mature = cdr[cdr["Disposition"].str.contains("callmatured|transfer",case=False,na=False)].shape[0]

    ib_mature = cdr[
        (cdr["Disposition"].str.contains("callmatured|transfer",case=False,na=False)) &
        (cdr["Campaign"].str.upper()=="CSRINBOUND")
    ].shape[0]

    ob_mature = total_mature - ib_mature

    total_talk = sec_to_time(df["Total Talk Time"].apply(time_to_sec).sum())
    aht = divide_time(total_talk,total_mature)

    grand={
        "Agent Name":"GRAND TOTAL",
        "Agent Full Name":"",
        "Total Login":"",
        "Total Net Login":"",
        "Total Break":"",
        "Total Meeting":"",
        "Total Talk Time":total_talk,
        "AHT":aht,
        "Total Mature":total_mature,
        "IB Mature":ib_mature,
        "OB Mature":ob_mature,
        "Total IVR Hit":total_ivr
    }

    return jsonify({"table":df.to_dict(orient="records"),"grand":grand})

@app.route("/export",methods=["POST"])
def export():

    data=pd.DataFrame(request.json)

    out=io.BytesIO()

    name=datetime.datetime.now().strftime("%d-%m-%y_%H-%M-%S")
    filename=f"Agent_Performance_Report_Chandan-Malakar_{name}.xlsx"

    with pd.ExcelWriter(out,engine="xlsxwriter") as writer:
        data.to_excel(writer,index=False,sheet_name="Report")

        wb=writer.book
        ws=writer.sheets["Report"]

        header=wb.add_format({"bold":True,"border":1,"bg_color":"#1fa463","color":"white","align":"center"})
        cell=wb.add_format({"border":1,"align":"center"})

        for col in range(len(data.columns)):
            ws.write(0,col,data.columns[col],header)

        for r in range(1,len(data)+1):
            for c in range(len(data.columns)):
                ws.write(r,c,data.iloc[r-1,c],cell)

        ws.autofilter(0,0,len(data),len(data.columns)-1)
        ws.freeze_panes(1,0)

    out.seek(0)
    return send_file(out,download_name=filename,as_attachment=True)

if __name__=="__main__":
    app.run()
