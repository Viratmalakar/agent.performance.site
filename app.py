from flask import Flask, request
import pandas as pd
import os

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
        return "❌ Could not detect all 4 reports. Check filenames."

    # ---------------- LOGIN ----------------
    login["Duration"] = pd.to_timedelta(login["Duration"], errors="coerce").dt.total_seconds()

    total_login = login.groupby("UserName")["Duration"].sum()

    break_mask = login["Activity"].str.contains("lunch|short|tea", case=False, na=False)
    total_break = login[break_mask].groupby("UserName")["Duration"].sum()

    meeting_mask = login["Activity"].str.contains("meeting|system", case=False, na=False)
    total_meeting = login[meeting_mask].groupby("UserName")["Duration"].sum()

    net_login = total_login - total_break

    # ---------------- AGENT PERFORMANCE ----------------
    agent["Talk Time"] = pd.to_timedelta(agent["Talk Time"], errors="coerce").dt.total_seconds()
    total_talk = agent.groupby("Agent Name")["Talk Time"].sum()

    # ---------------- CDR ----------------
    mature_mask = cdr["Disposition"].str.contains("callmatured|transfer", case=False, na=False)
    total_mature = cdr[mature_mask].groupby("Username").size()

    ib_mask = mature_mask & cdr["Campaign"].str.contains("csr", case=False, na=False)
    ib_mature = cdr[ib_mask].groupby("Username").size()

    transfer_call = cdr[cdr["Disposition"].str.contains("transfer", case=False, na=False)].groupby("Username").size()

    ob_mature = total_mature - ib_mature

    # ---------------- CRM ----------------
    total_tagging = crm.groupby("CreatedByID").size()

    # ---------------- FINAL ----------------
    final = pd.DataFrame({
        "EMP ID": total_login.index,
        "Agent Name": total_login.index,
        "Total Login": total_login,
        "Total Net Login": net_login,
        "Total Break": total_break,
        "Total Meeting": total_meeting,
        "Total Talk time": total_talk,
        "Total Mature": total_mature,
        "IB Mature": ib_mature,
        "Transfer Call": transfer_call,
        "OB Mature": ob_mature,
        "Total tagging": total_tagging
    }).fillna(0)

    final["AHT"] = (final["Total Talk time"] / final["Total Mature"]).fillna(0).round(2)

    # seconds to HH:MM:SS
    for c in ["Total Login","Total Net Login","Total Break","Total Meeting","Total Talk time"]:
        final[c] = pd.to_timedelta(final[c], unit="s")

    return final.to_html(index=False)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
