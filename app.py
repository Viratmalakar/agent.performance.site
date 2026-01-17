from flask import Flask, request, render_template, send_file
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def hhmmss_to_seconds(t):
    if pd.isna(t): return 0
    h,m,s = map(int,str(t).split(":"))
    return h*3600+m*60+s

def seconds_to_hhmmss(sec):
    sec=int(sec)
    return f"{sec//3600:02}:{(sec%3600)//60:02}:{sec%60:02}"

@app.route("/")
def home():
    return render_template("upload.html")

@app.route("/process", methods=["POST"])
def process():

    files = request.files.getlist("files")
    paths = []

    for f in files:
        path=os.path.join(UPLOAD_FOLDER,f.filename)
        f.save(path)
        paths.append(path)

    login = pd.read_excel(paths[0],header=2)
    cdr = pd.read_excel(paths[1],header=1)
    agent = pd.read_excel(paths[2],header=2)
    crm = pd.read_excel(paths[3])

    # First Login
    login['Date']=pd.to_datetime(login['Date'])
    first_login = login.groupby("UserName")['Date'].min().dt.time

    # CRM Tagging
    tagging = crm.groupby("CreatedByID").size()

    # CDR Mature
    cdr_mature=cdr[cdr['Disposition'].isin(["CALLMATURED","TRANSFER"])]
    total_mature=cdr_mature.groupby("Username").size()
    transfer_call=cdr[cdr['Disposition']=="TRANSFER"].groupby("Username").size()
    ib_mature=cdr_mature[cdr_mature['Campaign']=="CSRINBOUND"].groupby("Username").size()
    ob_mature=total_mature-ib_mature

    # Agent Calculation
    agent['Total Break']=agent['LUNCHBREAK']+agent['SHORTBREAK']+agent['TEABREAK']
    agent['Total Meeting']=agent['MEETING']+agent['SYSTEMDOWN']
    agent['Net Login']=agent['Total Login Time']-agent['Total Break']

    agent['Talk_sec']=agent['Total Talk Time'].apply(hhmmss_to_seconds)
    agent['Total Mature']=agent['Agent Name'].map(total_mature).fillna(0)
    agent['AHT']=agent['Talk_sec']/agent['Total Mature']

    final=pd.DataFrame()
    final['EMP ID']=agent['Agent Name']
    final['Agent Name']=agent['Agent Full Name']
    final['First Login Time']=agent['Agent Name'].map(first_login)
    final['Total Login']=agent['Total Login Time']
    final['Total Net Login']=agent['Net Login']
    final['Total Break']=agent['Total Break']
    final['Total Meeting']=agent['Total Meeting']
    final['Total Talk Time']=agent['Total Talk Time']
    final['AHT']=final['AHT']=agent['AHT'].apply(seconds_to_hhmmss)
    final['Total Mature']=agent['Agent Name'].map(total_mature)
    final['IB Mature']=agent['Agent Name'].map(ib_mature)
    final['Transfer Call']=agent['Agent Name'].map(transfer_call)
    final['OB Mature']=final['Total Mature']-final['IB Mature']
    final['Total Tagging']=agent['Agent Name'].map(tagging)

    output="Final_Report.xlsx"
    final.to_excel(output,index=False)

    return render_template("result.html", tables=final.to_html(index=False), file=output)

@app.route("/download")
def download():
    return send_file("Final_Report.xlsx",as_attachment=True)

if __name__=="__main__":
    app.run(host="0.0.0.0",port=int(os.environ.get("PORT",5000)))
