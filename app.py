from flask import Flask, request, render_template, send_file, flash, redirect, url_for
import pandas as pd
import io
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)
app.secret_key = 'super_secret_key'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        files = {
            'login': request.files.get('login_report'),
            'agent_perf': request.files.get('agent_performance'),
            'cdr': request.files.get('cdr_report'),
            'crm': request.files.get('crm_report')
        }
        
        if not all(files.values()):
            flash('Please upload all 4 files.')
            return redirect(request.url)
        
        # Save files temporarily
        file_paths = {}
        for name, file in files.items():
            if file and file.filename:
                filename = secure_filename(file.filename)
                path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(path)
                file_paths[name] = path
        
        try:
            final_df = process_reports(file_paths)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, sheet_name='Final Agent Performance', index=False)
            output.seek(0)
            
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='Final Agent Performance.xlsx'
            )
        except Exception as e:
            flash(f'Error processing files: {str(e)}')
            return redirect(url_for('upload_files'))
    
    return '''
    <!DOCTYPE html>
    <html>
    <head><title>Agent Performance Report Generator</title></head>
    <body>
        <h1>Upload Reports</h1>
        {% with messages = get_flashed_messages() %}
            {% if messages %}
                <ul>
                {% for message in messages %}
                    <li style="color: red;">{{ message }}</li>
                {% endfor %}
                </ul>
            {% endif %}
        {% endwith %}
        <form method="post" enctype="multipart/form-data">
            <p>Login Report: <input type="file" name="login_report" accept=".xlsx,.xls" required></p>
            <p>Agent Performance Report: <input type="file" name="agent_performance" accept=".xlsx,.xls" required></p>
            <p>CDR Report: <input type="file" name="cdr_report" accept=".xlsx,.xls" required></p>
            <p>CRM Report: <input type="file" name="crm_report" accept=".xlsx,.xls" required></p>
            <p><input type="submit" value="Start Processing"></p>
        </form>
    </body>
    </html>
    '''

def process_reports(file_paths):
    # Read all reports [web:1][web:5][web:7]
    login_df = pd.read_excel(file_paths['login'])
    agent_df = pd.read_excel(file_paths['agent_perf'])
    cdr_df = pd.read_excel(file_paths['cdr'])
    crm_df = pd.read_excel(file_paths['crm'])
    
    # Rename columns to EMP ID as per mapping
    if 'UserName' in login_df.columns:
        login_df['EMP ID'] = login_df['UserName']
    if 'Username' in cdr_df.columns:
        cdr_df['EMP ID'] = cdr_df['Username']
    if 'Agent Name' in agent_df.columns:
        agent_df = agent_df.rename(columns={'Agent Name': 'EMP ID'})  # Assume needs mapping or direct
    if 'CreatedByID' in crm_df.columns:
        crm_df['EMP ID'] = crm_df['CreatedByID']
    
    # 1. Total Login: sum login time per EMP ID, assume column 'Total Login Time', convert to timedelta [web:20][web:21]
    def time_to_td(time_str):
        if pd.isna(time_str):
            return pd.Timedelta(0)
        if isinstance(time_str, str):
            return pd.to_timedelta(time_str)
        elif isinstance(time_str, (int, float)):
            # Excel time serial
            return pd.to_timedelta(time_str, unit='D')
        return pd.Timedelta(0)
    
    login_df['Total_Login_td'] = login_df['Total Login Time'].apply(time_to_td) if 'Total Login Time' in login_df else pd.Timedelta(0)
    total_login = login_df.groupby('EMP ID')['Total_Login_td'].sum().reset_index()
    total_login['Total Login'] = total_login['Total_Login_td'].astype(str).str[:-7]  # HH:MM:SS
    
    # 2-3. Total Break and Net Login, assume columns exist
    login_df['Total_Break_td'] = (
        login_df.get('LUNCHBREAK', pd.Series(pd.Timedelta(0), index=login_df.index)).apply(time_to_td) +
        login_df.get('SHORTBREAK', pd.Series(pd.Timedelta(0), index=login_df.index)).apply(time_to_td) +
        login_df.get('TEABREAK', pd.Series(pd.Timedelta(0), index=login_df.index)).apply(time_to_td)
    )
    total_break = login_df.groupby('EMP ID')['Total_Break_td'].sum().reset_index()
    total_break['Total Break'] = total_break['Total_Break_td'].astype(str).str[:-7]
    
    total_net_login = total_login.merge(total_break, on='EMP ID')
    total_net_login['Total Net Login_td'] = total_net_login['Total_Login_td'] - total_net_login['Total_Break_td']
    total_net_login['Total Net Login'] = total_net_login['Total Net Login_td'].astype(str).str[:-7]
    
    # 4. Total Meeting
    login_df['Total_Meeting_td'] = (
        login_df.get('MEETING', pd.Series(pd.Timedelta(0), index=login_df.index)).apply(time_to_td) +
        login_df.get('SYSTEMDOWN', pd.Series(pd.Timedelta(0), index=login_df.index)).apply(time_to_td)
    )
    total_meeting = login_df.groupby('EMP ID')['Total_Meeting_td'].sum().reset_index()
    total_meeting['Total Meeting'] = total_meeting['Total_Meeting_td'].astype(str).str[:-7]
    
    # 5. Total Talk Time from Agent Performance, assume column 'Total Talk Time'
    talk_time = agent_df.groupby('EMP ID')['Total Talk Time'].sum().reset_index()
    talk_time['Total Talk Time'] = pd.to_numeric(talk_time['Total Talk Time'], errors='coerce').fillna(0)
    
    # 6. Total Mature: count Disposition in ['Callmatured', 'Transfer'] [web:14]
    cdr_df['mature'] = cdr_df['Disposition'].isin(['Callmatured', 'Transfer']).astype(int)
    total_mature = cdr_df.groupby('EMP ID')['mature'].sum().reset_index()
    total_mature.rename(columns={'mature': 'Total Mature'}, inplace=True)
    
    # 7. IB Mature: same + Campaign == 'CSRINBOUND'
    cdr_df['ib_mature'] = ((cdr_df['Disposition'].isin(['Callmatured', 'Transfer'])) & 
                           (cdr_df['Campaign'] == 'CSRINBOUND')).astype(int)
    ib_mature = cdr_df.groupby('EMP ID')['ib_mature'].sum().reset_index()
    ib_mature.rename(columns={'ib_mature': 'IB Mature'}, inplace=True)
    
    # 8. OB Mature
    matures = total_mature.merge(ib_mature, on='EMP ID', how='left')
    matures['OB Mature'] = matures['Total Mature'] - matures['IB Mature'].fillna(0)
    
    # 9. Total Tagging: count CRM
    total_tagging = crm_df.groupby('EMP ID').size().reset_index(name='Total Tagging')
    
    # 10. AHT = Total Talk Time / Total Mature (seconds? assume talk time in seconds)
    final = (total_net_login[['EMP ID', 'Total Login', 'Total Net Login', 'Total Break']]
             .merge(total_meeting[['EMP ID', 'Total Meeting']], on='EMP ID', how='outer')
             .merge(talk_time, on='EMP ID', how='outer')
             .merge(matures[['EMP ID', 'Total Mature', 'IB Mature', 'OB Mature']], on='EMP ID', how='outer')
             .merge(total_tagging, on='EMP ID', how='outer')
             .fillna(0))
    
    final['AHT'] = final['Total Talk Time'] / final['Total Mature'].replace(0, float('nan'))
    
    # Get Agent Name from agent_df, assume has Agent Name originally? Wait, mapping says Agent Name = EMP ID, but output needs Agent Name
    # Assume agent_df has both or use merge
    if 'Agent Name_original' in agent_df.columns:  # adjust if needed
        agent_name_map = agent_df[['EMP ID', 'Agent Name_original']].drop_duplicates()
    else:
        agent_name_map = pd.DataFrame({'EMP ID': final['EMP ID'].unique(), 'Agent Name': final['EMP ID']})
    final = final.merge(agent_name_map, on='EMP ID', how='left')
    
    # Reorder columns
    cols = ['EMP ID', 'Agent Name', 'Total Login', 'Total Net Login', 'Total Break', 
            'Total Meeting', 'Total Talk Time', 'AHT', 'Total Mature', 'IB Mature', 
            'OB Mature', 'Total Tagging']
    final = final[cols]
    
    return final.fillna(0)

if __name__ == '__main__':
    app.run(debug=True)
