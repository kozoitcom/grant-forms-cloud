"""Google Drive storage for Grant Forms Cloud (single-admin prototype).
Uses an OAuth refresh token kept only in Render environment variables.
"""
import io, json, os, time, threading, urllib.parse, uuid
import requests
from fastapi import HTTPException

_lock = threading.RLock()
_cache = {"token": None, "expires": 0}

def access_token():
    with _lock:
        if _cache["token"] and time.time() < _cache["expires"] - 60:
            return _cache["token"]
        required = ("GOOGLE_CLIENT_ID", "GOOGLE_CLIENT_SECRET", "GOOGLE_REFRESH_TOKEN", "GOOGLE_DRIVE_FOLDER_ID")
        if any(not os.getenv(k) for k in required):
            raise HTTPException(503, "Google Drive is not configured. See CLOUD_SETUP.md")
        r = requests.post("https://oauth2.googleapis.com/token", data={
            "client_id": os.environ["GOOGLE_CLIENT_ID"],
            "client_secret": os.environ["GOOGLE_CLIENT_SECRET"],
            "refresh_token": os.environ["GOOGLE_REFRESH_TOKEN"],
            "grant_type": "refresh_token"}, timeout=25)
        if not r.ok:
            raise HTTPException(503, "Google OAuth token refresh failed; check credentials and consent")
        data = r.json()
        _cache.update(token=data["access_token"], expires=time.time()+int(data.get("expires_in", 3600)))
        return _cache["token"]

def call(method, path, *, params=None, json_data=None, data=None, headers=None):
    h={"Authorization": "Bearer "+access_token()}
    h.update(headers or {})
    r=requests.request(method, "https://www.googleapis.com/drive/v3/"+path,
        params=params, json=json_data, data=data, headers=h, timeout=60)
    if not r.ok:
        raise HTTPException(502, "Google Drive request failed ("+str(r.status_code)+"): "+r.text[:250])
    return r

def find_folder(name, parent):
    q="name = '"+name.replace("'", "\\'")+"' and '"+parent+"' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    files=call("GET","files",params={"q":q,"fields":"files(id,name),nextPageToken","pageSize":100}).json()["files"]
    if files: return files[0]["id"]
    return call("POST","files",json_data={"name":name,"mimeType":"application/vnd.google-apps.folder","parents":[parent]},params={"fields":"id"}).json()["id"]

def collection(name):
    return find_folder(name,os.environ["GOOGLE_DRIVE_FOLDER_ID"])

def list_rows(table):
    folder=collection(table)
    result=[]; token=None
    while True:
        params={"q":"'"+folder+"' in parents and trashed = false","fields":"nextPageToken,files(id,name,createdTime)","pageSize":1000}
        if token: params["pageToken"]=token
        resp=call("GET","files",params=params).json()
        for f in resp.get("files",[]):
            if not f["name"].endswith(".json"): continue
            try:
                row=call("GET","files/"+f["id"],params={"alt":"media"}).json()
                row["_drive_id"]=f["id"]
                result.append(row)
            except Exception: continue
        token=resp.get("nextPageToken")
        if not token: break
    return result

def create_row(table,record):
    record=dict(record)
    record.setdefault("id",str(uuid.uuid4()))
    record.setdefault("created_at",__import__("datetime").datetime.now(__import__("datetime").timezone.utc).isoformat())
    body=json.dumps(record,ensure_ascii=False).encode()
    return upload_file(record["id"]+".json",body,collection(table),"application/json") and record

def upload_file(name, content, parent, mime="application/octet-stream"):
    boundary="grantforms"+uuid.uuid4().hex
    meta=json.dumps({"name":name,"parents":[parent]}).encode()
    payload=(b"--"+boundary.encode()+b"\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n"+meta+
        b"\r\n--"+boundary.encode()+b"\r\nContent-Type: "+mime.encode()+b"\r\n\r\n"+content+
        b"\r\n--"+boundary.encode()+b"--\r\n")
    r=requests.post("https://www.googleapis.com/upload/drive/v3/files",
        params={"uploadType":"multipart","fields":"id"},
        headers={"Authorization":"Bearer "+access_token(),"Content-Type":"multipart/related; boundary="+boundary},
        data=payload,timeout=120)
    if not r.ok: raise HTTPException(502,"Google Drive upload failed ("+str(r.status_code)+")")
    return r.json()["id"]

def update_row(table, row, changes):
    drive_id=row.pop("_drive_id")
    updated={**row,**changes}
    content=json.dumps(updated,ensure_ascii=False).encode()
    r=requests.patch("https://www.googleapis.com/upload/drive/v3/files/"+drive_id,
        params={"uploadType":"media"},headers={"Authorization":"Bearer "+access_token(),"Content-Type":"application/json"},
        data=content,timeout=60)
    if not r.ok: raise HTTPException(502,"Google Drive update failed")
    return updated

def get_row(table, row_id):
    for r in list_rows(table):
        if r["id"]==row_id: return r
    raise HTTPException(404,"Record not found")

def download_file(file_id):
    return call("GET","files/"+file_id,params={"alt":"media"}).content

def delete_file(file_id):
    call("DELETE","files/"+file_id)
