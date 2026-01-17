from flask import Flask, request, jsonify
import os
import time

app = Flask(__name__)
app.secret_key = "agent_secret_key"

THEME = """
<style>
body{
margin:0;
font-family:Arial;
background: linear-gradient(135deg,#7f5cff,#ff9acb,#6ec6ff);
background-size:400% 400%;
animation:bg 12s infinite;
}
@keyframes bg{
0%{background-position:0% 50%}
50%{background-position:100% 50%}
100%{background-position:0% 50%}
}
.box{
width:480px;
margin:80px auto;
background:rgba(255,255,255,0.95);
padding:25px;
border-radius:15px;
box-shadow:0 0 20px rgba(0,0,0,0.3);
text-align:center;
}
.drop{
border:2px dashed #777;
padding:30px;
border-radius:10px;
margin-bottom:15px;
cursor:pointer;
background:#fafafa;
}
.progress{
width:100%;
background:#ddd;
border-radius:6px;
overflow:hidden;
height:10px;
margin-top:6px;
}
.bar{
height:10px;
width:0%;
background:#2196f3;
}
.msg{
font-weight:bold;
margin-top:10px;
color:#333;
}
</style>
"""

# ---------------- HOME ----------------
@app.route("/")
def home():
    return THEME + """
    <div class='box'>
    <h2>Agent Performance Upload</h2>

    <div class='drop' id='drop'>
        Drag & Drop Excel Files Here<br>
        or Click to Select
        <input type='file' id='files' multiple style='display:none'>
    </div>

    <div id='fileList'></div>

    <button onclick='startUpload()'>Start Upload</button>

    <div id='status' class='msg'></div>
    </div>

<script>
let drop=document.getElementById("drop");
let input=document.getElementById("files");

drop.onclick=()=>input.click();

drop.ondragover=e=>{
e.preventDefault();
drop.style.background="#e0f0ff";
};

drop.ondragleave=e=>{
drop.style.background="#fafafa";
};

drop.ondrop=e=>{
e.preventDefault();
drop.style.background="#fafafa";
input.files=e.dataTransfer.files;
showFiles();
};

input.onchange=showFiles;

function kb(x){return (x/1024).toFixed(1)+" KB";}

function showFiles(){
let list=document.getElementById("fileList");
list.innerHTML="";
for(let i=0;i<input.files.length;i++){
let f=input.files[i];
list.innerHTML+=`
<div>
${f.name} (${kb(f.size)})
<div class='progress'><div class='bar' id='bar${i}'></div></div>
</div>`;
}
}

function startUpload(){
let files=input.files;
if(files.length==0){
alert("Please select files");
return;
}

let formData=new FormData();
for(let i=0;i<files.length;i++){
formData.append("files",files[i]);
}

let xhr=new XMLHttpRequest();
xhr.open("POST","/process",true);

xhr.upload.onprogress=function(e){
if(e.lengthComputable){
let p=(e.loaded/e.total)*100;
for(let i=0;i<files.length;i++){
document.getElementById("bar"+i).style.width=p+"%";
}
}
};

xhr.onload=function(){
document.getElementById("status").innerHTML="⚙️ Work in progress... Calculations running...";
};

xhr.send(formData);
}
</script>
    """

# ---------------- PROCESS ----------------
@app.route("/process", methods=["POST"])
def process():
    files=request.files.getlist("files")

    # simulate heavy calculation time
    time.sleep(5)

    return jsonify({"status":"done"})
    
# ---------------- RUN ----------------
if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0",port=port)
