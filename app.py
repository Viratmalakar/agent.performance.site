from flask import Flask, render_template, request, redirect, session, send_file
from users import USERS
import os
import pandas as pd

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

# ---------------- PROCESS FILES ----------------
@app.route("/process", methods=["POST"])
def process():

    files = request.files.getlist("files")

    login_df = None
    cdr_df = None
    agent_df = None
    crm_df = None

    for f in files:
        name = f.filename.lower()

        if "login" in name:
            login_df = pd.read_excel(f)
        elif "cdr" in name:
            cdr_df = pd.read_excel(f)
        elif "agent" in name:
            agent_df = pd.read_excel(f)
        elif "crm" in name:
            crm_df = pd.read_excel(f)

    if any(df is None for df in [login_df, cdr_df, agent_df, crm_df]):
        return "Please upload all 4 correct Excel files (login, cdr, agent, crm)"

    # ---------- TEMP TEST OUTPUT ----------
    result_df = pd.DataFrame({
        "Status": ["Login file OK","CDR file OK","Agent file OK","CRM file OK"]
    })

    output_file = "Test_Output.xlsx"
    result_df.to_excel(output_file, index=False)

    return send_file(output_file, as_attachment=True)

# ---------------- RUN SERVER ----------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
