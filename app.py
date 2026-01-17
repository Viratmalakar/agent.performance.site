from flask import Flask, request, send_file
import pandas as pd
import os, time

app = Flask(__name__)

THEME = """
<style>
body{margin:0;font-family:Arial;background:linear-gradient(135deg,#7f5cff,#ff9acb,#6ec6ff);background-size:400% 400%;animation:bg 12s infinite;}
@keyframes bg{0%{background-position:0% 50%}50%{background-position:100% 50%}100%{background-position:0% 50%}}
.box{width:500px;margin:80px auto;background:white;padding:25px;border-radius:15px;text-align:center;}
.progress{width:100%;background:#ddd;height:10px;border-radius:5px;overflow:hidden}
.bar{height:10px;width:0%;background:#2196f3}
</style>
"""

@app.route("/")
def home():
    return THEME + """
    <div class='box'>
    <h2>Agent Performance Upload</h2>
    <input type='file' id='files' multiple><br><br>
    <div id='list'></div>
    <button onclick='upload()'>Start Upload</button>
    <div id='msg'></div>
    </div>

<script>
function upload(){
 let f=document.getElementById("files").files;
 let fd=new FormData();
 for(let i=0;i<f.length;i++) fd.append("files",f[i]);

 let xhr=new XMLHttpRequest();
 xhr.open("POST","/process",true);

 xhr.upload.onprogress=function(e){
  if(e.lengthComputable){
   let p=(e.loaded/e.total)*100;
   document.getElementById("msg").innerHTML="Uploading... "+p.toFixed(1)+"%";
  }
 }

 xhr.onload=function(){
  document.getElementById("msg").innerHTML="⚙️ Work in progress... Calculations running...";
  setTimeout(()=>window.location='/download',4000);
 }

 xhr.send(fd);
}
</script>
"""

@app.route("/process", methods=["POST"])
def process():
    files=request.files.getlist("files")

    login=cdr=agent=crm=None

    for f in files:
        df=pd.read_excel(f)
        cols=" ".join(df.columns.astype(str)).lower()

        if "login" in cols: login=df
        elif "disposition" in cols: cdr=df
        elif "talk" in cols: agent=df
        elif "createdby" in cols: crm=df

    if any(x is None for x in [login,cdr,agent,crm]):
        return "File detection failed"

    # ===== YOUR FORMULAS =====
    emp="User Name" if "User Name" in login.columns else login.columns[0]

    login["LoginSec"]=pd.to_timedelta(login.iloc[:,1]).dt.total_seconds()
    login_sum=login.groupby(emp)["LoginSec"].sum().reset_index()

    agent["TalkSec"]=pd.to_timedelta(agent.iloc[:,1]).dt.total_seconds()
    talk_sum=agent.groupby(agent.columns[0])["TalkSec"].sum().reset_index()

    mature=cdr[cdr.iloc[:,2].str.contains("mature|transfer",case=False,na=False)]
    mature_cnt=mature.groupby(cdr.columns[0]).size().reset_index(name="Total Mature")

    crm_cnt=crm.groupby(crm.columns[0]).size().reset_index(name="Total Tagging")

    final=login_sum.merge(talk_sum,left_on=emp,right_on=agent.columns[0],how="left")
    final=final.merge(mature_cnt,left_on=emp,right_on=cdr.columns[0],how="left")
    final=final.merge(crm_cnt,left_on=emp,right_on=crm.columns[0],how="left")

    final.fillna(0,inplace=True)

    final["AHT"]=(final["TalkSec"]/final["Total Mature"]).fillna(0).round(2)

    final_out=final[[emp,"LoginSec","TalkSec","Total Mature","Total Tagging","AHT"]]
    final_out.columns=["EMP ID","Total Login(sec)","Total Talk(sec)","Total Mature","Total Tagging","AHT"]

    final_out.to_excel("Final_Agent_Performance.xlsx",index=False)

    return "done"

@app.route("/download")
def download():
    return send_file("Final_Agent_Performance.xlsx",as_attachment=True)

if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0",port=port)
