from flask import Flask, request, send_file
import pandas as pd
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

app = Flask(__name__)

RESULT_FILE = "Final_Result.xlsx"
PDF_FILE = "Final_Result.pdf"

THEME = """
<style>
body{
margin:0;
font-family:Arial;
background: linear-gradient(135deg,#7f5cff,#ff9acb,#6ec6ff);
background-size:400% 400%;
animation:bg 12s infinite;
}
@keyframes bg{
0%{background-position:0% 50%}
50%{background-position:100% 50%}
100%{background-position:0% 50%}
}
.box{
width:85%;
margin:60px auto;
background:white;
padding:25px;
border-radius:15px;
text-align:center;
}
table{
width:100%;
border-collapse:collapse;
font-size:13px;
}
th,td{
border:1px solid #999;
padding:6px;
}
th{
background:#ffd27f;
}
.btn{
display:inline-block;
padding:10px 20px;
background:#2196f3;
color:white;
border-radius:6px;
margin:10px;
text-decoration:none;
font-weight:bold;
}
</style>
"""

@app.route("/")
def home():
    return THEME + """
    <div class='box'>
    <h2>Upload Excel Files</h2>
    <form action="/process" method="post" enctype="multipart/form-data">
        <input type="file" name="files" multiple><br><br>
        <button class='btn'>Upload & Calculate</button>
    </form>
    </div>
    """

@app.route("/process", methods=["POST"])
def process():
    files = request.files.getlist("files")

    # ---- SAFE DEMO CALCULATION ----
    data = []
    for f in files:
        data.append({"File Name": f.filename, "Status": "Processed"})

    df = pd.DataFrame(data)

    # save excel
    df.to_excel(RESULT_FILE, index=False)

    # create pdf
    create_pdf(df)

    table_html = df.to_html(index=False)

    return THEME + f"""
    <div class='box'>
    <h2>Calculation Completed ✅</h2>
    {table_html}
    <br>
    <a class='btn' href='/download_excel'>⬇ Download Excel</a>
    <a class='btn' href='/download_pdf'>⬇ Download PDF</a>
    </div>
    """

def create_pdf(df):
    c = canvas.Canvas(PDF_FILE, pagesize=A4)
    width, height = A4

    # Top right name
    c.setFont("Helvetica-Bold", 10)
    c.drawRightString(width-40, height-40, "Chandan Malakar")

    y = height - 80
    c.setFont("Helvetica", 10)

    for col in df.columns:
        c.drawString(40, y, str(col))
        y -= 15

    y -= 10

    for _, row in df.iterrows():
        for val in row:
            c.drawString(40, y, str(val))
            y -= 15
        y -= 10

        if y < 50:
            c.showPage()
            y = height - 80
            c.drawRightString(width-40, height-40, "Chandan Malakar")

    c.save()

@app.route("/download_excel")
def download_excel():
    return send_file(RESULT_FILE, as_attachment=True)

@app.route("/download_pdf")
def download_pdf():
    return send_file(PDF_FILE, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
