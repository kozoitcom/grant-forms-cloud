/**
 * Grant Forms Cloud - Google Apps Script edition (prototype).
 * Deploy as Web App, execute as owner. Set Script Properties before use.
 * NEVER use real applicant data until security and privacy review is complete.
 */
const ROOT_PROP = 'DRIVE_FOLDER_ID';
const SHEET_PROP = 'SPREADSHEET_ID';
const TYPES = ['text','email','number','date','textarea','select','file'];
const MAX_FILE_BYTES = 5 * 1024 * 1024;
const ALLOWED_MIME = ['application/pdf','image/jpeg','image/png',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.formId = String((e && e.parameter && e.parameter.form) || '');
  return t.evaluate().setTitle('助成金フォーム').addMetaTag('viewport','width=device-width, initial-scale=1');
}
function settings_() { return PropertiesService.getScriptProperties(); }
function root_() {
  const id = settings_().getProperty(ROOT_PROP);
  if (!id) throw new Error('DRIVE_FOLDER_ID をスクリプトプロパティに設定してください');
  return DriveApp.getFolderById(id);
}
function child_(parent,name) {
  const it=parent.getFoldersByName(name);
  return it.hasNext()?it.next():parent.createFolder(name);
}
function spreadsheet_() {
  const props=settings_();let id=props.getProperty(SHEET_PROP);
  if (!id) {
    const s=SpreadsheetApp.create('助成金フォーム_管理データ');
    id=s.getId();props.setProperty(SHEET_PROP,id);
    DriveApp.getFileById(id).moveTo(root_());
  }
  return SpreadsheetApp.openById(id);
}
function tab_(name,headers) {
  const ss=spreadsheet_();let sh=ss.getSheetByName(name);
  if (!sh) { sh=ss.insertSheet(name);sh.appendRow(headers);sh.setFrozenRows(1); }
  return sh;
}
function formTab_(){return tab_('Forms',['id','title','fields_json','published','created_at']);}
function answerTab_(){return tab_('Submissions',['id','form_id','company','answers_json','attachments_json','created_at']);}
function records_(sh) {
  const a=sh.getDataRange().getValues();
  if(a.length<2)return [];
  const h=a.shift().map(String);
  return a.filter(r=>r[0]).map((r,i)=>({row:i+2,...Object.fromEntries(h.map((k,j)=>[k,r[j]]))}));
}
function form_(id) {
  const r=records_(formTab_()).find(x=>String(x.id)===String(id));
  if(!r)throw new Error('フォームが見つかりません');
  return {id:String(r.id),title:String(r.title),fields:JSON.parse(r.fields_json),published:String(r.published)==='true',row:r.row};
}
function validate_(fields) {
  if(!Array.isArray(fields)||fields.length>100)throw new Error('フォーム項目が不正です');
  const ids=new Set();
  fields.forEach(f=>{
    if(!/^[A-Za-z][A-Za-z0-9_]{0,79}$/.test(f.id)||ids.has(f.id)||!TYPES.includes(f.type)||!f.label)
      throw new Error('フォーム項目のID・型・ラベルが不正です');
    ids.add(f.id);
    if(f.type==='select'&&!Array.isArray(f.options))throw new Error('選択肢が必要です');
  });
  return fields;
}
function hash_(s) {
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s);
  return Utilities.base64EncodeWebSafe(bytes);
}
function adminLogin(password) {
  const expected=settings_().getProperty('ADMIN_PASSWORD')||'';
  if(expected.length<16||hash_(String(password))!==hash_(expected))throw new Error('管理者パスワードが違います');
  const token=Utilities.getUuid()+Utilities.getUuid();
  CacheService.getScriptCache().put('admin:'+hash_(token),'1',3600);
  return token;
}
function requireAdmin_(token) {
  if(!token||CacheService.getScriptCache().get('admin:'+hash_(String(token)))!=='1')
    throw new Error('管理者ログインが必要です');
}
function adminList(token) {
  requireAdmin_(token);
  return {forms:records_(formTab_()).map(r=>({id:String(r.id),title:String(r.title),published:String(r.published)==='true'})),
    submissions:records_(answerTab_()).map(r=>({id:String(r.id),formId:String(r.form_id),company:String(r.company),createdAt:String(r.created_at)}))};
}
function adminSaveForm(token,id,title,fields,published) {
  requireAdmin_(token);fields=validate_(fields);
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {
    const sh=formTab_();const now=new Date().toISOString();
    if(id) {
      const old=form_(id);
      sh.getRange(old.row,2,1,3).setValues([[String(title),JSON.stringify(fields),Boolean(published)]]);
      return id;
    }
    id=Utilities.getUuid();sh.appendRow([id,String(title),JSON.stringify(fields),Boolean(published),now]);
    return id;
  } finally {lock.releaseLock();}
}
function adminGetForm(token,id){requireAdmin_(token);return form_(id);}
function publicGetForm(id) {
  const f=form_(id);
  if(!f.published)throw new Error('現在、このフォームは公開されていません');
  return {id:f.id,title:f.title,fields:f.fields};
}
function publicSubmit(id,answers,files) {
  const f=publicGetForm(id);
  if(!answers||typeof answers!=='object'||!Array.isArray(files))throw new Error('入力内容が不正です');
  if(files.length>15)throw new Error('添付ファイルが多すぎます');
  const clean={};const uploaded={};
  f.fields.forEach(field=>{
    if(field.type==='file')return;
    const v=String(answers[field.id]||'').trim();
    if(field.required&&!v)throw new Error(field.label+' は必須です');
    if(v.length>20000)throw new Error(field.label+' が長すぎます');
    if(field.type==='select'&&v&&!field.options.includes(v))throw new Error('選択肢が不正です');
    if(field.type==='email'&&v&&!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v))throw new Error('メールアドレスが不正です');
    if(field.type==='number'&&v&&!/^\d+$/.test(v))throw new Error('数字を入力してください');
    clean[field.id]=v;
  });
  const allowed=new Set(f.fields.filter(x=>x.type==='file').map(x=>x.id));
  files.forEach(file=>{
    if(!allowed.has(file.fieldId)||uploaded[file.fieldId])throw new Error('添付項目が不正です');
    if(!ALLOWED_MIME.includes(file.mimeType))throw new Error('対応していないファイル形式です');
    if(typeof file.base64!=='string'||file.base64.length>MAX_FILE_BYTES*1.4)throw new Error('ファイルが大きすぎます');
    uploaded[file.fieldId]=file;
  });
  f.fields.filter(x=>x.type==='file'&&x.required).forEach(x=>{
    if(!uploaded[x.id])throw new Error(x.label+' は必須です');
  });
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {
    const sid=Utilities.getUuid();
    const folder=child_(child_(root_(),'Submissions'),sid);
    const attachments=[];
    files.forEach(file=>{
      const bytes=Utilities.base64Decode(file.base64);
      if(bytes.length>MAX_FILE_BYTES)throw new Error('ファイルは5MB以内です');
      const name=String(file.name||'attachment').replace(/[\\/<>:"|?*\x00-\x1f]/g,'_').slice(0,120);
      const blob=Utilities.newBlob(bytes,file.mimeType,name);
      const driveFile=folder.createFile(blob);
      attachments.push({fieldId:file.fieldId,name:name,fileId:driveFile.getId()});
    });
    const company=String(clean.company_name||clean.company||'').replace(/^[=+@-]/, c => "'"+c);
    answerTab_().appendRow([sid,id,company,JSON.stringify(clean),JSON.stringify(attachments),new Date().toISOString()]);
    return sid;
  } finally {lock.releaseLock();}
}
function adminSubmission(token,id) {
  requireAdmin_(token);
  const r=records_(answerTab_()).find(x=>String(x.id)===String(id));
  if(!r)throw new Error('回答が見つかりません');
  return {id:String(r.id),formId:String(r.form_id),answers:JSON.parse(r.answers_json),
    attachments:JSON.parse(r.attachments_json),createdAt:String(r.created_at)};
}
function adminGenerate(token,pdfBase64,name) {
  requireAdmin_(token);
  // Placeholder only; API key is not used until PDF extraction is implemented.
  if(!/\.pdf$/i.test(name))throw new Error('PDFのみ対応です');
  const bytes=Utilities.base64Decode(pdfBase64);
  if(bytes.length>5*1024*1024)throw new Error('PDFは5MB以内です');
  // Apps Script alone cannot reliably extract text from arbitrary PDFs.
  // Avoid pretending a file was analyzed; PDF parsing will be added separately.
  throw new Error('PDF本文抽出は未実装です。募集要項のテキスト抽出機能を追加してから使用してください');
}
