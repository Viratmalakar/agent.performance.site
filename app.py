@app.route("/upload")
def upload():
    if "user" not in session:
        return redirect("/")

    return THEME + """
    <div class='box'>
    <h2>Upload Excel Files</h2>

    <input type='file' id='files' multiple><br><br>
    <div id='list'></div>

    <button onclick='uploadFiles()'>Start Upload</button>
    <br><br><a href='/logout'>Logout</a>
    </div>

<script>
function kb(x){ return (x/1024).toFixed(1)+" KB"; }

function uploadFiles(){
    let files=document.getElementById("files").files;
    let list=document.getElementById("list");
    list.innerHTML="";

    let formData=new FormData();

    for(let i=0;i<files.length;i++){
        let f=files[i];
        formData.append("files",f);

        list.innerHTML+=`
        <div style='margin:10px 0;text-align:left'>
        ${f.name} (${kb(f.size)})
        <div style='width:100%;background:#ddd;height:8px;border-radius:5px'>
        <div id='bar${i}' style='width:0%;height:8px;background:#2196f3'></div>
        </div>
        <span id='txt${i}'></span>
        </div>`;
    }

    let xhr=new XMLHttpRequest();
    xhr.open("POST","/process",true);

    xhr.upload.onprogress=function(e){
        if(e.lengthComputable){
            let p=(e.loaded/e.total)*100;
            for(let i=0;i<files.length;i++){
                document.getElementById("bar"+i).style.width=p+"%";
                document.getElementById("txt"+i).innerHTML=
                kb(e.loaded)+" / "+kb(e.total);
            }
        }
    };

    xhr.onload=function(){
        for(let i=0;i<files.length;i++){
            document.getElementById("txt"+i).innerHTML="Uploaded ✔";
        }
    };

    xhr.send(formData);
}
</script>
    """
