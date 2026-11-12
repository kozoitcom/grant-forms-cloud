import os, io, re, json, uuid, hmac, hashlib, html
from pathlib import Path
from fastapi import FastAPI, Request, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse
from supabase import create_client
from pypdf import PdfReader

app = FastAPI(title="Grant Forms Cloud")
BUCKET = "application-documents"
TYPES = {"text","email","number","date","textarea","select","file"}

def db():
    if not os.getenv("SUPABASE_URL") or not os.getenv("SUPABASE_SERVICE_ROLE_KEY"):
        raise HTTPException(503,"Supabase is not configured")
    return create_client(os.environ["SUPABASE_URL"],os.environ["SUPABASE_SERVICE_ROLE_KEY"])

def esc(s):
    return html.escape(str(s or ""),quote=True)

def token():
    secret=os.getenv("ADMIN_PASSWORD","")
    return hmac.new(secret.encode(),b"grant-admin-v1",hashlib.sha256).hexdigest()

def admin(req):
    return len(os.getenv("ADMIN_PASSWORD",""))>=16 and hmac.compare_digest(req.cookies.get("admin",""),token())

def guard(req):
    if not admin(req): raise HTTPException(401,"Admin login required")

def page(title,body):
    return HTMLResponse('<!doctype html><html lang="ja"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>'+esc(title)+'</title><style>body{font:16px system-ui;margin:24px auto;padding:0 16px;max-width:850px;line-height:1.6}input,select,textarea,button{font:inherit;padding:9px;max-width:100%;box-sizing:border-box}input:not([type=checkbox]),textarea,select{width:100%}label{display:block;margin:16px 0 5px}button{cursor:pointer}article{padding:14px;border:1px solid #ccc;border-radius:8px;margin:14px 0}small{color:#666}a{color:#1458a5}</style><h1>'+esc(title)+'</h1>'+body+'</html>')

def validate(schema):
    if not isinstance(schema,dict) or not isinstance(schema.get("fields"),list): raise ValueError("fields is required")
    ids=set()
    for f in schema["fields"]:
        fid=str(f.get("id",""))
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,79}",fid) or fid in ids: raise ValueError("Invalid or duplicate field ID")
        ids.add(fid)
        if f.get("type") not in TYPES or not f.get("label"): raise ValueError("Invalid field")
        if f["type"]=="select" and not isinstance(f.get("options"),list): raise ValueError("Select requires options")
    return schema

@app.get("/health")
def health(): return {"status":"ok"}

@app.get("/")
def root(req:Request): return RedirectResponse("/admin" if admin(req) else "/login")

@app.get("/login")
def login(): return page("管理者ログイン",'<form method="post"><label>パスワード</label><input type="password" name="password" required><p><button>ログイン</button></p></form>')

@app.post("/login")
def login_post(password:str=Form(...)):
    expected=os.getenv("ADMIN_PASSWORD","")
    if len(expected)<16 or not hmac.compare_digest(password,expected): raise HTTPException(401,"Invalid password or admin password not configured")
    r=RedirectResponse("/admin",303)
    r.set_cookie("admin",token(),httponly=True,secure=True,samesite="strict",max_age=3600*4)
    return r

@app.get("/admin")
def dashboard(req:Request):
    if not admin(req): return RedirectResponse("/login",303)
    forms=db().table("forms").select("id,title,published").order("created_at",desc=True).execute().data
    submissions=db().table("submissions").select("id,form_id,created_at").order("created_at",desc=True).limit(100).execute().data
    body='<h2>募集要項PDFからフォーム案を生成</h2><form method="post" action="/admin/generate" enctype="multipart/form-data"><input type="file" name="pdf" accept=".pdf" required><p><button>AI解析・下書き作成</button></p></form><h2>フォーム一覧</h2>'
    for f in forms:
        fid=esc(f["id"])
        body+='<article><b>'+esc(f["title"])+'</b> ('+('公開中' if f["published"] else '非公開')+')<p><a href="/admin/edit/'+fid+'">編集・公開設定</a> | <a href="/f/'+fid+'">フォームを表示</a></p></article>'
    body+='<h2>企業からの回答</h2>'
    for s in submissions: body+='<article><a href="/admin/submission/'+esc(s["id"])+'">'+esc(s["created_at"])+' / '+esc(s["id"])+'</a></article>'
    return page("助成金フォーム管理",body)

@app.post("/admin/generate")
async def generate(req:Request,pdf:UploadFile=File(...)):
    guard(req)
    if not (pdf.filename or "").lower().endswith(".pdf"): raise HTTPException(400,"PDFのみ対応")
    content=await pdf.read(15*1024*1024+1)
    if len(content)>15*1024*1024: raise HTTPException(413,"PDFは15MB以内")
    try: text="\n".join(p.extract_text() or "" for p in PdfReader(io.BytesIO(content)).pages)
    except Exception: raise HTTPException(400,"PDFを読み取れません")
    if len(text.strip())<100: raise HTTPException(400,"画像PDFは未対応です")
    if not os.getenv("OPENAI_API_KEY"): raise HTTPException(503,"OPENAI_API_KEY is not configured")
    from openai import OpenAI
    instruction='日本の助成金の募集要項から申請時に企業が入力・提出する情報を抽出してください。JSONのみ返答。形式: {"title":"名称","fields":[{"id":"ascii_snake_case","label":"日本語","type":"text|email|number|date|textarea|select|file","required":true,"options":["選択肢"],"help":"補足"}]}。書類はfile、選択肢はselect。申請時以外の実績報告書類は除外。曖昧な必須はfalse。出典のない項目は追加しない。'
    try:
        response=OpenAI().chat.completions.create(model=os.getenv("OPENAI_MODEL","gpt-4.1-mini"),messages=[{"role":"system","content":instruction},{"role":"user","content":text[:90000]}],response_format={"type":"json_object"})
        schema=validate(json.loads(response.choices[0].message.content))
        result=db().table("forms").insert({"title":schema.get("title","新規フォーム"),"source_name":pdf.filename,"fields":schema["fields"],"published":False}).execute().data[0]
    except Exception: raise HTTPException(502,"AI解析または保存に失敗しました")
    return RedirectResponse("/admin/edit/"+result["id"],303)

@app.get("/admin/edit/{fid}")
def edit(req:Request,fid:str):
    guard(req)
    f=db().table("forms").select("*").eq("id",fid).single().execute().data
    body='<p>AIの生成結果を必ず確認してください。保存すると非公開に戻ります。</p><form method="post"><label>フォーム名</label><input name="title" value="'+esc(f["title"])+'" required><label>項目JSON</label><textarea name="fields" rows="22" required>'+esc(json.dumps(f["fields"],ensure_ascii=False,indent=2))+'</textarea><p><button>下書きを保存</button></p></form>'
    body+='<form method="post" action="/admin/publish/'+esc(fid)+'"><button>'+('非公開にする' if f["published"] else '確認済み・公開する')+'</button></form><p><a href="/admin">管理画面へ戻る</a></p>'
    return page("フォーム確認・編集",body)

@app.post("/admin/edit/{fid}")
def edit_post(req:Request,title:str=Form(...),fields:str=Form(...),fid:str=""):
    guard(req)
    try: fs=validate({"fields":json.loads(fields)})["fields"]
    except Exception as e: raise HTTPException(400,str(e))
    db().table("forms").update({"title":title,"fields":fs,"published":False}).eq("id",fid).execute()
    return RedirectResponse("/admin/edit/"+fid,303)

@app.post("/admin/publish/{fid}")
def publish(req:Request,fid:str):
    guard(req)
    client=db(); f=client.table("forms").select("published").eq("id",fid).single().execute().data
    client.table("forms").update({"published":not f["published"]}).eq("id",fid).execute()
    return RedirectResponse("/admin/edit/"+fid,303)

@app.get("/f/{fid}")
def form(req:Request,fid:str):
    f=db().table("forms").select("*").eq("id",fid).single().execute().data
    if not f["published"] and not admin(req): raise HTTPException(404,"フォームは非公開です")
    body='<form method="post" enctype="multipart/form-data">'
    for field in f["fields"]:
        fid2=esc(field["id"]); typ=field["type"]; required=" required" if field.get("required") else ""
        body+='<label>'+esc(field["label"])+(' *' if required else '')+'</label>'
        if field.get("help"): body+='<small>'+esc(field["help"])+'</small>'
        if typ=="textarea": body+='<textarea name="'+fid2+'"'+required+'></textarea>'
        elif typ=="select":
            body+='<select name="'+fid2+'"'+required+'><option value="">選択してください</option>'
            for opt in field.get("options",[]): body+='<option value="'+esc(opt)+'">'+esc(opt)+'</option>'
            body+='</select>'
        else: body+='<input type="'+("file" if typ=="file" else typ)+'" name="'+fid2+'"'+required+(' accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg,.png"' if typ=="file" else '')+'>'
    body+='<p><button>回答・書類を送信</button></p></form>'
    return page(f["title"],body)

@app.post("/f/{fid}")
async def submit(req:Request,fid:str):
    client=db(); f=client.table("forms").select("*").eq("id",fid).single().execute().data
    if not f["published"]: raise HTTPException(403,"受付停止中")
    data=await req.form(); answers={}; uploads=[]
    for field in f["fields"]:
        key=field["id"]; value=data.get(key); typ=field["type"]
        if typ=="file":
            if isinstance(value,UploadFile) and value.filename:
                payload=await value.read(10*1024*1024+1)
                if len(payload)>10*1024*1024: raise HTTPException(413,"ファイルは10MB以内")
                if not (payload.startswith(b"%PDF-") or payload.startswith(b"PK\x03\x04") or payload.startswith(b"\xff\xd8\xff") or payload.startswith(b"\x89PNG")): raise HTTPException(400,"非対応のファイル形式")
                uploads.append((key,value.filename,payload))
            elif field.get("required"): raise HTTPException(400,key+" は必須")
        else:
            v=str(value or "").strip()
            if field.get("required") and not v: raise HTTPException(400,key+" は必須")
            if typ=="select" and v and v not in field.get("options",[]): raise HTTPException(400,"不正な選択肢")
            if typ=="number" and v and not re.fullmatch(r"[0-9]+",v): raise HTTPException(400,"数字を入力")
            if typ=="email" and v and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+",v): raise HTTPException(400,"メール形式が不正")
            answers[key]=v[:20000]
    sid=str(uuid.uuid4()); paths=[]
    try:
        for key,filename,payload in uploads:
            extension=Path(filename).suffix.lower()[:8]
            path=f"{fid}/{sid}/{key}-{uuid.uuid4().hex}{extension}"
            client.storage.from_(BUCKET).upload(path,payload,{"content-type":"application/octet-stream"})
            paths.append((key,filename,path))
        client.table("submissions").insert({"id":sid,"form_id":fid,"answers":answers}).execute()
        if paths: client.table("attachments").insert([{"submission_id":sid,"field_id":k,"filename":n,"object_path":p} for k,n,p in paths]).execute()
    except Exception:
        for _,_,path in paths:
            try: client.storage.from_(BUCKET).remove([path])
            except Exception: pass
        raise HTTPException(500,"保存に失敗しました")
    return page("送信完了","<p>回答を受け付けました。受付番号："+esc(sid)+"</p>")

@app.get("/admin/submission/{sid}")
def detail(req:Request,sid:str):
    guard(req); client=db()
    row=client.table("submissions").select("*").eq("id",sid).single().execute().data
    attachments=client.table("attachments").select("*").eq("submission_id",sid).execute().data
    body="<h2>回答</h2><pre style='white-space:pre-wrap'>"+esc(json.dumps(row["answers"],ensure_ascii=False,indent=2))+"</pre><h2>添付書類</h2>"
    for a in attachments: body+='<p><a href="/admin/file/'+esc(a["id"])+'">'+esc(a["filename"])+"</a></p>"
    return page("企業回答詳細",body)

@app.get("/admin/file/{aid}")
def download(req:Request,aid:str):
    guard(req)
    client=db(); a=client.table("attachments").select("*").eq("id",aid).single().execute().data
    result=client.storage.from_(BUCKET).create_signed_url(a["object_path"],60)
    return RedirectResponse(result.get("signedURL") or result.get("signedUrl"))
