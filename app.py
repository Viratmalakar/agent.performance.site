from flask import Flask, request, redirect, session
import os

app = Flask(__name__)
app.secret_key = "agent_secret_key"

USERS = {"admin": "12345"}

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
width:360px;
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
ul{list-style:none;padding:0;}
li{font-weight:bold;}
a{text-decoration:none;font-weight:bold;}
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
    <h2>Upload Files</h2>
    <form action='/process' method='post' enctype='multipart/form-data'>
    <input type='file' name='files' multiple onchange='showFiles(this)'><br>
    <ul id='fileList'></ul>
    <button>Submit</button>
    </form>
    <br><a href='/logout'>Logout</a>
    </div>

    <script>
    function showFiles(input){
        let list=document.getElementById("fileList");
        list.innerHTML="";
        for(let i=0;i<input.files.length;i++){
            let li=document.createElement("li");
            li.innerHTML=input.files[i].name+" ✔";
            list.appendChild(li);
        }
    }
    </script>
    """

# ---------- PROCESS ----------
@app.route("/process", methods=["POST"])
def process():
    files=request.files.getlist("files")

    html = THEME + "<div class='box'><h2>Uploaded Files</h2><ul>"
    for f in files:
        html += f"<li>{f.filename} ✔</li>"
    html += "</ul><br><a href='/upload'>Upload More</a></div>"

    return html

# ---------- LOGOUT ----------
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

# ---------- RUN ----------
if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0",port=port)
