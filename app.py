<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Upload Reports</title>

<style>
:root{
  --bg:#f6f4fb;
  --card:#ffffff;
  --primary:#7c6ee6;
  --secondary:#9f8cff;
  --accent:#5a4fdc;
  --text:#2f2f3a;
  --muted:#7a7a99;
  --border:#e3defa;
}

*{box-sizing:border-box}

body{
  margin:0;
  min-height:100vh;
  display:flex;
  align-items:center;
  justify-content:center;
  font-family:"Segoe UI",system-ui,sans-serif;
  background:linear-gradient(135deg,#f6f4fb,#efeafd);
  color:var(--text);
}

.wrapper{
  width:900px;
  display:grid;
  grid-template-columns:1fr 1fr;
  gap:26px;
}

.card{
  background:var(--card);
  border-radius:22px;
  padding:30px;
  box-shadow:0 18px 40px rgba(92,79,220,.15);
  border:1px solid var(--border);
}

.card h3{
  margin:0 0 6px;
  color:var(--accent);
}

.card p{
  margin:0 0 18px;
  font-size:13px;
  color:var(--muted);
}

.file-box{
  border:2px dashed var(--secondary);
  border-radius:16px;
  padding:22px;
  text-align:center;
  background:#faf9ff;
}

input[type=file]{
  width:100%;
  color:var(--muted);
}

button{
  grid-column:1/3;
  margin-top:10px;
  padding:14px;
  border:none;
  border-radius:18px;
  background:linear-gradient(90deg,var(--primary),var(--accent));
  color:#fff;
  font-size:16px;
  font-weight:600;
  cursor:pointer;
  box-shadow:0 12px 28px rgba(124,110,230,.35);
}

button:hover{
  transform:translateY(-2px);
}

.footer{
  grid-column:1/3;
  text-align:center;
  font-size:12px;
  color:var(--muted);
  margin-top:8px;
}
</style>
</head>

<body>

<form action="/process" method="post" enctype="multipart/form-data">
  <div class="wrapper">

    <div class="card">
      <h3>Agent Performance Report</h3>
      <p>Upload agent performance Excel</p>
      <div class="file-box">
        <input type="file" name="agent_files" multiple required>
      </div>
    </div>

    <div class="card">
      <h3>CDR Report</h3>
      <p>Upload call detail record Excel</p>
      <div class="file-box">
        <input type="file" name="cdr_files" multiple>
      </div>
    </div>

    <button type="submit">Upload & Generate Dashboard</button>

    <div class="footer">
      Secure upload • Agent-wise analytics • Premium purple theme
    </div>

  </div>
</form>

</body>
</html>
