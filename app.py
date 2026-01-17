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
            df = pd.read_excel(f)

            if "login" in name:
                login = df
            elif "cdr" in name:
                cdr = df
            elif "agent" in name:
                agent = df
            elif "crm" in name or "detail" in name:
                crm = df

        if any(x is None for x in [login, cdr, agent, crm]):
            return "❌ Could not detect all 4 reports."

        # Test only column existence
        return f"""
        Login columns: {list(login.columns)}<br><br>
        CDR columns: {list(cdr.columns)}<br><br>
        Agent columns: {list(agent.columns)}<br><br>
        CRM columns: {list(crm.columns)}
        """

    except Exception as e:
        return "<pre>" + traceback.format_exc() + "</pre>"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
