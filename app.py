from flask import Flask, render_template, request, redirect, session
from users import USERS
import os

app = Flask(__name__)
app.secret_key = "agent_secret_key"

# ---------------- LOGIN ----------------
@app.route("/", methods=["GET","POST"])
def login():
    if request.method=="POST":
        u = request.form["username"]
        p = request.form["password"]
        if u in USERS and USERS[u] == p:
            session["user"] = u
            return redirect("/dashboard")
        return render_template("login.html", error="Invalid Username or Password")
    return render_template("login.html")

# ---------------- DASHBOARD ----------------
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect("/")
    return render_template("index.html")

# ---------------- LOGOUT ----------------
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

# ---------------- CHANGE PASSWORD ----------------
@app.route("/change_password")
def change_password():
    if session.get("user")!="admin":
        return redirect("/")
    return render_template("change_password.html")

@app.route("/update_password", methods=["POST"])
def update_password():
    if session.get("user")!="admin":
        return redirect("/")
    old = request.form["old_password"]
    new = request.form["new_password"]

    if USERS["admin"] != old:
        return render_template("change_password.html", error="Old password wrong")

    USERS["admin"] = new
    return render_template("login.html", success="Password changed successfully")

# ---------------- UPLOAD PAGE ----------------
@app.route("/upload")
def upload():
    if "user" not in session:
        return redirect("/")
    return render_template("upload.html")

# ---------------- PROCESS EXCEL ----------------
@app.route("/process", methods=["POST"])
def process():
    login = request.files['login_report']
    cdr = request.files['cdr_report']
    agent = request.files['agent_report']
    crm = request.files['crm_report']

    return "Files uploaded successfully"

# ---------------- RUN SERVER ----------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
