from flask import Flask, render_template, request, redirect, session
import os

app = Flask(__name__)
app.secret_key = "testkey"

@app.route("/")
def home():
    return "Server is running fine ✅"

@app.route("/process", methods=["GET","POST"])
def process():
    return "Process route working fine ✅"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
