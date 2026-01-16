from flask import Flask, request, redirect, session
import os

app = Flask(__name__)
app.secret_key = "agent_secret_key"

USERS = {
    "admin": "12345"
}

# ---------------- LOGIN ----------------
@app.route("/", methods=["GET","POST"])
def login():
    if request.method == "POST":
        u = request.form["username"]
        p = request.form["password"]
        if u in USERS and USERS[u] == p:
            session["user"] = u
            return redirect("/upload")
        return "Invalid login"

    return '''
    <h2>Login</h2>
    <form method="post">
        <input name="username" placeholder="Username"><br><br>
        <input type="password" name="password" placeholder="Password"><br><br>
        <button type="submit">Login</button>
    </form>
    '''

# ---------------- UPLOAD PAGE ----------------
@app.route("/upload")
def upload():
    if "user" not in session:
        return redirect("/")
    return '''
    <h2>Upload Files</h2>

    <form action="/process" method="post" enctype="multipart/form-data">
        <input type="file" name="files" multiple onchange="showFiles(this)"><br><br>
        <ul id="fileList"></ul>
        <button type="submit">Submit</button>
    </form>

    <br><a href="/logout">Logout</a>

    <script>
    function showFiles(input){
        let list = document.getElementById("fileList");
        list.innerHTML = "";
        for(let i=0;i<input.files.length;i++){
            let li = document.createElement("li");
            li.innerHTML = input.files[i].name + " ✔";
            list.appendChild(li);
        }
    }
    </script>
    '''

# ---------------- PROCESS ----------------
@app.route("/process", methods=["POST"])
def process():
    files = request.files.getlist("files")

    html = "<h2>Uploaded Files Preview</h2><ul>"

    for f in files:
        html += f"<li>{f.filename} ✔</li>"

    html += "</ul>"
    html += "<br><a href='/upload'>Upload More</a>"
    html += "<br><a href='/logout'>Logout</a>"

    return html

# ---------------- LOGOUT ----------------
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

# ---------------- RUN ----------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
