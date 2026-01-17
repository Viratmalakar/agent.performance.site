from flask import Flask, request, jsonify
import os
import time
from werkzeug.utils import secure_filename
import pandas as pd  # Excel processing के लिए

app = Flask(__name__)
app.secret_key = "agent_secret_key"
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

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
transition: width 0.3s ease;
}
.msg{
font-weight:bold;
margin-top:10px;
color:#333;
}
.file-item {
margin:10px 0;
text-align:left;
}
</style>
"""

# ---------------- HOME ----------------
@app.route("/")
def home():
    return THEME + """
    <div class='box'>
    <h2>Agent Performance Upload</h2>
    <p>Excel/CSV files drag-drop करें - Auto process होगा (UNIQUE, FILTER जैसे calculations)</p>

    <div class='drop' id='drop'>
        Drag & Drop Excel Files Here<br>
        or Click to Select
        <input type='file' id='files' multiple accept='.xlsx,.xls,.csv' style='display:none'>
    </div>

    <div id='fileList'></div>

    <button id='uploadBtn' onclick='startUpload()' disabled>Start Upload</button>

    <div id='status' class='msg'></div>
    </div>

<script>
let drop=document.getElementById("drop");
let input=document.getElementById("files");
let uploadBtn=document.getElementById("uploadBtn");

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
let files=input.files;
if(files.length===0){
    uploadBtn.disabled=true;
    return;
}
list.innerHTML="";
for(let i=0;i<files.length;i++){
let f=files[i];
list.innerHTML+=`
<div class='file-item'>
<strong>${f.name}</strong> (${kb(f.size)})
<div class='progress'><div class='bar' id='bar${i}'></div></div>
<span id='status${i}'></span>
</div>`;
}
uploadBtn.disabled=false;
uploadBtn.textContent=`Start Upload (${files.length} files)`;
}

async function startUpload(){
let files=input.files;
if(files.length==0){
    alert("Please select files");
    return;
}

let formData=new FormData();
for(let i=0;i<files.length;i++){
formData.append("files",files[i]);
}

uploadBtn.disabled=true;
uploadBtn.textContent="Uploading...";

let xhr=new XMLHttpRequest();
xhr.open("POST","/process",true);

xhr.upload.onprogress=function(e){
if(e.lengthComputable){
let p=Math.round((e.loaded/e.total)*100);
document.getElementById("status").innerHTML=`Upload Progress: ${p}%`;
for(let i=0;i<files.length;i++){
document.getElementById("bar"+i).style.width=p+"%";
}
}
};

xhr.onload=function(){
if(xhr.status===200){
document.getElementById("status").innerHTML="✅ Upload Success! Processing Excel files...";
setTimeout(()=>location.reload(), 3000);  // Reload after process
} else {
document.getElementById("status").innerHTML="❌ Error: " + xhr.responseText;
}
};

xhr.onerror=function(){
document.getElementById("status").innerHTML="❌ Network Error";
};

xhr.send(formData);
}
</script>
    """

# ---------------- PROCESS ----------------
@app.route("/process", methods=["POST"])
def process():
    try:
        files = request.files.getlist("files")
        if not files or len(files) == 0:
            return jsonify({"error": "No files received"}), 400

        results = []
        for file in files:
            if file.filename == '':
                continue
            if not allowed_file(file.filename):
                return jsonify({"error": f"Invalid file type: {file.filename}"}), 400

            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            # Excel Processing Example (आपके UNIQUE/FILTER needs के लिए)
            try:
                if filename.endswith('.csv'):
                    df = pd.read_csv(filepath)
                else:
                    df = pd.read_excel(filepath)
                
                unique_agents = df['Agent'].nunique() if 'Agent' in df.columns else 0
                total_records = len(df)
                
                results.append({
                    "filename": filename,
                    "records": total_records,
                    "unique_agents": unique_agents,
                    "status": "processed"
                })
                
                # Simulate heavy calculation
                time.sleep(2)
                
            except Exception as e:
                results.append({"filename": filename, "error": str(e)})

        return jsonify({
            "status": "done",
            "processed": len([r for r in results if r.get("status") == "processed"]),
            "results": results
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ---------------- RESULTS (optional) ----------------
@app.route("/results")
def results():
    files = [f for f in os.listdir(app.config['UPLOAD_FOLDER']) if allowed_file(f)]
    return jsonify({"files": files})

# ---------------- RUN ----------------
if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0", port=port, debug=True)
