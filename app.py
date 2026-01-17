# app.py (Flask Version)
from flask import Flask, request, send_file, render_template_string
import pandas as pd
import io
import uuid

app = Flask(__name__)

HTML_PAGE = """
<html>
<head>
<title>Agent Performance Processor</title>
</head>
<body style="font-family:Arial; max-width:800px; margin:auto;">
<h2>Final Agent Performance Generator</h2>
<form action="/process" method="post" enctype="multipart/form-data">
Login Report: <input type="file" name="login"><br><br>
Agent Performance Report: <input type="file" name="agent"><br><br>
CDR Report: <input type="file" name="cdr"><br><br>
CRM Report: <input type="file" name="crm"><br><br><br>
<button type="submit">Start Processing</button>
</form>
</body>
</html>
"""

@app.route("/")
def home():
    return render_template_string(HTML_PAGE)

# ---------- Helpers ----------
def to_seconds(x):
    if pd.isna(x):
        return 0
    return int(pd.to_timedelta(x).total_seconds())

def to_hms(sec):
    return str(pd.to_timedelta(int(sec), unit='s'))

# ---------- Processing ----------
@app.route("/process", methods=["POST"])
def process_files():

    login = request.files['login']
    agent = request.files['agent']
    cdr = request.files['cdr']
    crm = request.files['crm']

    login_df = pd.read_excel(login)
    agent_df = pd.read_excel(agent)
    cdr_df = pd.read_excel(cdr)
    crm_df = pd.read_excel(crm)

    # EMP ID mapping
    login_df['EMP ID'] = login_df['UserName']
    agent_df['EMP ID'] = agent_df['Agent Name']
    cdr_df['EMP ID'] = cdr_df['Username']
    crm_df['EMP ID'] = crm_df['CreatedByID']

    # Total Login
    login_df['LOGIN_SEC'] = login_df['Total Login Time'].apply(to_seconds)
    total_login = login_df.groupby('EMP ID')['LOGIN_SEC'].sum().reset_index()

    # Total Break
    login_df['BREAK_SEC'] = (
        login_df['LUNCHBREAK'].apply(to_seconds) +
        login_df['SHORTBREAK'].apply(to_seconds) +
        login_df['TEABREAK'].apply(to_seconds)
    )
    total_break = login_df.groupby('EMP ID')['BREAK_SEC'].sum().reset_index()

    # Total Meeting
    login_df['MEET_SEC'] = login_df['MEETING'].apply(to_seconds) + login_df['SYSTEMDOWN'].apply(to_seconds)
    total_meeting = login_df.groupby('EMP ID')['MEET_SEC'].sum().reset_index()

    # Total Talk Time
    agent_df['TALK_SEC'] = agent_df['Total Talk Time'].apply(to_seconds)
    talk_time = agent_df.groupby('EMP ID')['TALK_SEC'].sum().reset_index()

    # Total Mature
    matured = cdr_df[cdr_df['Disposition'].isin(['Callmatured','Transfer'])]
    total_mature = matured.groupby('EMP ID').size().reset_index(name='Total Mature')

    # IB Mature
    ib = matured[matured['Campaign'] == 'CSRINBOUND']
    ib_mature = ib.groupby('EMP ID').size().reset_index(name='IB Mature')

    # Total Tagging
    tagging = crm_df.groupby('EMP ID').size().reset_index(name='Total Tagging')

    # Merge
    final = total_login.merge(total_break, on='EMP ID', how='left')
    final = final.merge(total_meeting, on='EMP ID', how='left')
    final = final.merge(talk_time, on='EMP ID', how='left')
    final = final.merge(total_mature, on='EMP ID', how='left')
    final = final.merge(ib_mature, on='EMP ID', how='left')
    final = final.merge(tagging, on='EMP ID', how='left')

    final = final.fillna(0)

    # Calculations
    final['Total Net Login'] = final['LOGIN_SEC'] - final['BREAK_SEC']
    final['OB Mature'] = final['Total Mature'] - final['IB Mature']
    final['AHT_SEC'] = final.apply(lambda r: r['TALK_SEC']/r['Total Mature'] if r['Total Mature']>0 else 0, axis=1)

    # Convert
    final['Total Login'] = final['LOGIN_SEC'].apply(to_hms)
    final['Total Break'] = final['BREAK_SEC'].apply(to_hms)
    final['Total Net Login'] = final['Total Net Login'].apply(to_hms)
    final['Total Meeting'] = final['MEET_SEC'].apply(to_hms)
    final['Total Talk Time'] = final['TALK_SEC'].apply(to_hms)
    final['AHT'] = final['AHT_SEC'].apply(to_hms)

    # Agent Name
    name_map = agent_df[['EMP ID','Agent Name']].drop_duplicates()
    final = final.merge(name_map, on='EMP ID', how='left')

    final = final[[
        'EMP ID','Agent Name','Total Login','Total Net Login','Total Break','Total Meeting',
        'Total Talk Time','AHT','Total Mature','IB Mature','OB Mature','Total Tagging'
    ]]

    file_name = f"Final_Agent_Performance_{uuid.uuid4().hex}.xlsx"
    path = f"/mnt/data/{file_name}"
    final.to_excel(path, index=False)

    return send_file(path, as_attachment=True, download_name="Final Agent Performance.xlsx")

if __name__ == "__main__":
    app.run(debug=True)
