from flask import Flask, request
import pandas as pd
import os

app = Flask(__name__)

@app.route("/")
def home():
    return """
    <h2>Upload 4 Excel Reports</h2>
    <form action="/process" method="post" enctype="multipart/form-data">
        <input type="file" name="files" multiple required><br><br>
        <button>Upload & Calculate</button>
    </form>
    """

@app.route("/process", methods=["POST"])
def process():

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

    # ---------------- AGENT PERFORMANCE ----------------
    agent["Total Login Time"] = pd.to_timedelta(agent["Total Login Time"]).dt.total_seconds()

    agent["Total Break"] = (
        pd.to_timedelta(agent["LUNCHBREAK"]).dt.total_seconds() +
        pd.to_timedelta(agent["SHORTBREAK"]).dt.total_seconds() +
        pd.to_timedelta(agent["TEABREAK"]).dt.total_seconds()
    )

    agent["Total Net Login"] = agent["Total Login Time"] - agent["Total Break"]

    agent["Total Meeting"] = (
        pd.to_timedelta(agent["MEETING"]).dt.total_seconds() +
        pd.to_timedelta(agent["SYSTEMDOWN"]).dt.total_seconds()
    )

    agent["Total Talk Time"] = pd.to_timedelta(agent["Total Talk Time"]).dt.total_seconds()

    # ---------------- CDR ----------------
    total_mature = cdr[cdr["Disposition"].str.contains("callmatured|transfer",case=False,na=False)]\
                    .groupby("Username").size()

    ib_mature = cdr[
        cdr["Disposition"].str.contains("callmatured|transfer",case=False,na=False) &
        cdr["Campaign"].str.contains("csr",case=False,na=False)
    ].groupby("Username").size()

    transfer_call = cdr[cdr["Disposition"].str.contains("transfer",case=False,na=False)]\
                        .groupby("Username").size()

    ob_mature = total_mature - ib_mature

    # ---------------- CRM ----------------
    total_tagging = crm.groupby("CreatedByID").size()

    # ---------------- FINAL MERGE ----------------
    final = pd.DataFrame({
        "EMP ID": agent["Agent Name"],
        "Agent Name": agent["Agent Name"],
        "Total Login": agent["Total Login Time"],
        "Total Net Login": agent["Total Net Login"],
        "Total Break": agent["Total Break"],
        "Total Meeting": agent["Total Meeting"],
        "Total Talk time": agent["Total Talk Time"],
        "Total Mature": total_mature.reindex(agent["Agent Name"]).fillna(0).values,
        "IB Mature": ib_mature.reindex(agent["Agent Name"]).fillna(0).values,
        "Transfer Call": transfer_call.reindex(agent["Agent Name"]).fillna(0).values,
        "OB Mature": ob_mature.reindex(agent["Agent Name"]).fillna(0).values,
        "Total tagging": total_tagging.reindex(agent["Agent Name"]).fillna(0).values
    })

    final["AHT"] = (final["Total Talk time"] / final["Total Mature"]).fillna(0).round(2)

    for c in ["Total Login","Total Net Login","Total Break","Total Meeting","Total Talk time"]:
        final[c] = pd.to_timedelta(final[c], unit="s")

    return final.to_html(index=False)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
