import pandas as pd

@app.route("/process", methods=["POST"])
def process():
    files = request.files.getlist("files")

    df = pd.read_excel(files[0])   # first file preview

    table = df.head(10).to_html(index=False)

    html = THEME + """
    <div class='box' style='width:80%'>
    <h2>Excel Preview (First 10 Rows)</h2>
    """ + table + """
    <br>
    <a href='/download'>⬇ Download Excel</a><br><br>
    <a href='/upload'>Upload More</a>
    </div>
    """

    df.to_excel("Final_Output.xlsx", index=False)

    return html
