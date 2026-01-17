from flask import Flask, request, send_file
import pandas as pd
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

app = Flask(__name__)

RESULT_EXCEL = "Final_Agent_Report.xlsx"
RESULT_PDF = "Final_Agent_Report.pdf"
LAST_HTML = "last_result.html"

THEME = """
<style>
body{
margin:0;
font-family:Arial;
background:linear-gradient(135deg,#7f5cff,#ff9acb,#6ec6ff);
background-size:400% 400%;
animation:bg 12s infinite;
}
@keyframes bg{
0%{background-position:0% 50%}
50%{background-position:100% 50%}
100%{background-position:0% 50%}
}
.box{
width:90%;
margin:50px auto;
background:white;
padding:25px;
border-radius:15px;
text-align:center;
}
table{
width:100%;
border-collapse:collapse;
font-size:13px;
}
th,td{
border:1px solid #999;
padding:6px;
}
th{
background:#ffd27f;
}
.btn{
display:inline-block;
padding:10px 20px;
background:#2196f3;
color:white;
border-radius:6px;
margin:10px;
text-decoration:none;
font-weight:bold;
cursor:pointer;
}
.progress{
width:100%;
background:#ddd;
height:10px;
border-radius:6px;
overflow:hidden;
margin-top:10px;
}
.bar{
height:10px;
width:0%;
background:#2196f3;
}
#overlay{
display:none;
position:fixed;
top:0;
left:0;
width:100%;
height:100%;
background:rgba(0,0,0,0.6);
color:white;
font-size:22px;
align-items:center;
justify-content:center;
z-index:999;
}
</style>
"""

# ---------------- HOME ----------------
@app.route("/")
def home():
    return THEME + """
<div class='box'>
<h2>Upload 4 Excel Reports</h2>

<input type="file" id="files" multiple><br><br>

<div id="filelist"></div>

<button class='btn' onclick="startUpload()">Upload & Process</button>

<div class="progress"><div id="bar" class="bar"></div></div>

</div>

<div id="overlay">Processing your reports... Please wait ⏳</div>

<script>
let input=document.getElementById("files");

input.onchange=function(){
 let list="";
 for(let i=0;i<this.files.length;i++){
   list+=this.files[i].name+" ("+(this.files[i].size/1024).toFixed(1)+" KB)<br>";
 }
 document.getElementById("filelist").innerHTML=list;
}

function startUpload(){
 let f=input.files;
 if(f.length==0){alert("Select files first");return;}

 let fd=new FormData();
 for(let i=0;i<f.length;i++) fd.append("files",f[i]);

 document.getElementById("overlay").style.display="flex";

 let xhr=new XMLHttpRequest();
 xhr.open("POST","/process",true);

 xhr.upload.onprogress=function(e){
   if(e.lengthComputable){
     let p=(e.loaded/e.total)*100;
     document.getElementById("bar").style.width=p+"%";
   }
 }

 xhr.onload=function(){
   window.location="/result";
 }

 xhr.send(fd);
}
</script>
"""

# ---------------- PROCESS ----------------
@app.route("/process", methods=["POST"])
def process():
    files = request.files.getlist("files")

    login=cdr=agent=crm=None

    for f in files:
        name=f.filename.lower()
        df=pd.read_excel(f)

        if "login" in name:
            login=df
        elif "cdr" in name or "call" in name:
            cdr=df
        elif "agent" in name or "performance" in name:
            agent=df
        elif "crm" in name or "detail" in name:
            crm=df

    if any(x is None for x in [login,cdr,agent,crm]):
        html = THEME + "<div class='box'><h3>❌ Could not detect all 4 reports.</h3></div>"
        open(LAST_HTML,"w",encoding="utf-8").write(html)
        return ""

    # ====== SAFE DEMO CALCULATION (will replace with your formulas next) ======
    data=[]
    for f in files:
        data.append({"File":f.filename,"Status":"Processed"})

    final=pd.DataFrame(data)

    final.to_excel(RESULT_EXCEL,index=False)
    create_pdf(final)

    table=final.to_html(index=False)

    html = THEME + f"""
    <div class='box'>
    <h2>Calculation Completed ✅</h2>
    {table}
    <br>
    <a class='btn' href='/download_excel'>⬇ Download Excel</a>
    <a class='btn' href='/download_pdf'>⬇ Download PDF</a>
    </div>
    """

    open(LAST_HTML,"w",encoding="utf-8").write(html)
    return ""

# ---------------- RESULT PAGE ----------------
@app.route("/result")
def result():
    return open(LAST_HTML,encoding="utf-8").read()

# ---------------- PDF ----------------
def create_pdf(df):
    c = canvas.Canvas(RESULT_PDF,pagesize=A4)
    w,h=A4

    c.setFont("Helvetica-Bold",10)
    c.drawRightString(w-40,h-40,"Chandan Malakar")

    y=h-80
    c.setFont("Helvetica",9)

    for col in df.columns:
        c.drawString(40,y,str(col))
        y-=15

    y-=10

    for _,row in df.iterrows():
        for val in row:
            c.drawString(40,y,str(val))
            y-=15
        y-=10
        if y<50:
            c.showPage()
            c.drawRightString(w-40,h-40,"Chandan Malakar")
            y=h-80

    c.save()

# ---------------- DOWNLOADS ----------------
@app.route("/download_excel")
def download_excel():
    return send_file(RESULT_EXCEL,as_attachment=True)

@app.route("/download_pdf")
def download_pdf():
    return send_file(RESULT_PDF,as_attachment=True)

# ---------------- RUN ----------------
if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0",port=port)
