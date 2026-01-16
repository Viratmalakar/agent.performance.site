@app.route("/process", methods=["POST"])
def process():
    return """
    <h2>Server is running fine ✅</h2>
    <p>Excel processing logic will be added step by step.</p>
    <a href='/upload'>Go Back</a>
    """
