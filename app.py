from flask import Flask, request
import pandas as pd
import os
import traceback

app = Flask(__name__)

@app.route("/")
def home():
    return """
    <h2>Upload 4 Excel Reports</h2>
    <form action="/process" method="post" enctype="multipart/form-data">
        <input type="file" name="files" multiple required>
        <br><br>
        <button>Upload</button>
    </form>
    """

@app.route("/process", methods=["POST"])
def process():
    try:
        files = request.files.getlist("files")

        login = cdr = agent = crm = None

        for f in files:
            name = f.filename.lower()

            if "login" in name:
                login = pd.read_excel(f, header=2)

            elif "cdr" in name:
                cdr = pd.read_excel(f, header=1)

            elif "agent" in name:
                agent = pd.read_excel(f, header=2)

            elif "crm" in name or "detail" in name:
                crm = pd.read_excel(f, header=0)

        if any(x is None for x in [login, cdr, agent, crm]):
            return "❌ All 4 reports not detected properly."

        return f"""
        ✅ Files Loaded Successfully<br><br>

        <b>Login Columns:</b><br>{list(login.columns)}<br><br>
        <b>CDR Columns:</b><br>{list(cdr.columns)}<br><br>
        <b>Agent Columns:</b><br>{list(agent.columns)}<br><br>
        <b>CRM Columns:</b><br>{list(crm.columns)}
        """

    except Exception:
        return "<pre>" + traceback.format_exc() + "</pre>"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
