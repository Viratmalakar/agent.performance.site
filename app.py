from flask import Flask, request, redirect, session, jsonify
import os

app = Flask(__name__)
app.secret_key = "agent_secret_key"

USERS = {"admin":"12345"}

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
width:420px;
margin:120px auto;
background:rgba(255,255,255,0.9);
padding:25px;
border-radius:15px;
box-shadow:0 0 20px rgba(0,0,0,0.3);
text-align:center;
}
input,button{
width:90%;
padding:10px;
margin:8px;
border-radius:6px;
border:1px solid #ccc;
}
button{
background:#ffb84d;
font-weight:bold;
cursor:pointer;
}
.file-item{
margin:10px 0;
font-size:13px;
text-align:left;
}
.progress{
width:100%;
background:#ddd;
border-radius:5px;
overflow:hidden;
height:8px;
margin-top:4px;
}
.bar{
height:8px;
width:0%;
background:#2196f3;
}
.success{
color:green;
font-weight:bold;
}
</style>
"""

# ---------- LOGIN ----------
@app.route("/", methods=["GET","POST"])
def login():
    if request.method=="POST":
        u=request.form["username"]
        p=request.form["password"]
        if u in USERS and USERS[u]==p:
            session["user"]=u
            return redirect("/upload")
        return THEME + "<div class='box'>❌ Invalid Login</div>"

    return THEME + """
    <div class='box'>
    <h2>Login</h2>
    <form method='post'>
    <input name='username' placeholder='Username'>
    <input type='password' name='password' placeholder='Password'>
    <button>Login</button>
    </form>
    </div>
    """

# ---------- UPLOAD ----------
@app.route("/upload")
def upload():
    if "user" not in session:
        return redirect("/")

    return THEME + """
    <div class='box'>
    <h2>Upload Excel Files</h2>

    <input type='file' id='files' multiple><br>
    <div id='list'></div>
    <button onclick='uploadFiles()'>Upload Files</button>

    <br><br><a href='/logout'>Logout</a>
    </div>

<script>
function formatKB(size){
    return (size/1024).toFixed(1)+" KB";
}

function uploadFiles(){
    let files=document.getElementById("files").files;
    let list=document.getElementById("list");
    list.innerHTML="";

    let formData=new FormData();

    for(let i=0;i<files.length;i++){
        let f=files[i];
        formData.append("files",f);

        let div=document.createElement("div");
        div.className="file-item";
        div.id="file"+i;

        div.innerHTML=f.name+" ("+formatKB(f.size)+")"+
        "<div class='progress'><div class='bar' id='bar"+i+"'></div></div>"+
        "<span id='text"+i+"'></span>";

        list.appendChild(div);
    }

    let xhr=new XMLHttpRequest();
    xhr.open("POST","/process",true);

    xhr.upload.onprogress=function(e){
        if(e.lengthComputable){
            let percent=(e.loaded/e.total)*100;
            for(let i=0;i<files.length;i++){
                document.getElementById("bar"+i).style.width=percent+"%";
                document.getElementById("text"+i).innerHTML=
                formatKB(e.loaded)+" / "+formatKB(e.total);
            }
        }
    };

    xhr.onload=function(){
        if(xhr.status==200){
            for(let i=0;i<files.length;i++){
                document.getElementById("text"+i).innerHTML=
                "<span class='success'>Uploaded ✔</span>";
            }
        }
    };

    xhr.send(formData);
}
</script>
    """

# ---------- PROCESS ----------
@app.route("/process", methods=["POST"])
def process():
    files=request.files.getlist("files")
    return jsonify({"status":"success","files":[f.filename for f in files]})

# ---------- LOGOUT ----------
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

# ---------- RUN ----------
if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0",port=port)
