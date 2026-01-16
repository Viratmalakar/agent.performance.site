from flask import Flask, request, send_file
import pandas as pd
import os

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process():
    files = request.files.getlist('files')
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        dfs.append(df)
    final = pd.concat(dfs, ignore_index=True)
    output = 'FinalAgentPerformance.xlsx'
    final.to_excel(output, index=False)
    return send_file(output, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
