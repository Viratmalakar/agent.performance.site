from flask import Flask, request
import os

app = Flask(__name__)

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
width:480px;
margin:80px auto;
background:white;
padding:25px;
border-radius:15px;
text-align:center;
}
</style>
"""

@app.route("/")
def home():
    return THEME + """
    <div class='box'>
    <h2>Agent Performance Upload</h2>

    <form action="/process" method="post" enctype="multipart/form-data">
        <input type="file" name="files" multiple><br><br>
        <button>Upload</button>
    </form>

    </div>
    """

@app.route("/process", methods=["POST"])
def process():
    files = request.files.getlist("files")

    return THEME + f"""
    <div class='box'>
    <h2>✅ Files Uploaded</h2>
    <p>{len(files)} file(s) received</p>
    <p>⚙️ Work in progress… Calculations running</p>
    </div>
    """

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
