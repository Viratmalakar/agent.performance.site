<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Upload Reports</title>
<style>
body{
  margin:0;min-height:100vh;
  display:flex;justify-content:center;align-items:center;
  background:#f6f4fb;font-family:Segoe UI,sans-serif;
}
.card{
  width:500px;background:#fff;padding:30px;
  border-radius:20px;
  box-shadow:0 15px 40px rgba(124,110,230,.25);
}
h2{color:#6a5acd}
input,button{width:100%;margin-top:12px}
button{
  padding:12px;border:none;border-radius:14px;
  background:#6a5acd;color:#fff;font-size:15px
}
</style>
</head>
<body>

<div class="card">
<h2>Upload Reports</h2>
<form action="/process" method="post" enctype="multipart/form-data">
  <label>Agent Performance</label>
  <input type="file" name="agent_files" multiple required>

  <label>CDR Report</label>
  <input type="file" name="cdr_files" multiple>

  <button type="submit">Generate Dashboard</button>
</form>
</div>

</body>
</html>
