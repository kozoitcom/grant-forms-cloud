/** System 3 owns immutable transfer snapshots; never updates the collection or ledger. */
const TRANSFER_FOLDER_KEY = 'GRANT_TRANSFER_SNAPSHOTS_FOLDER_ID';

function getTransferPreparation(receipt) {
  requireGrantAdmin_();
  const application = collectionExport_(String(receipt || ''));
  // The only inputs from Systems 1 and 2 are explicit, serializable data contracts.
  const materialTemplates = application.grant.startsWith('schema:') ? requirementRead_(application.grant.slice(7)).batchIds.flatMap(id=>materialBatch_(id).manifest.files.filter(f=>f.role!=='guidelines').map(f=>({key:f.documentKey||f.documentTitle,title:f.documentTitle,[f.role]:{id:f.fileId,name:f.name,url:f.url}}))) : listGrantDocumentSettings(application.grant);
  const templates = materialTemplates.map(document => ({
    key:document.key, title:document.title, blank:document.blank || null,
    sample:document.sample || null, uploadNotRequired:Boolean(document.uploadNotRequired)
  }));
  return {schemaVersion:1, application, templates, generationAvailable:false};
}

function saveTransferPreparation(receipt) {
  requireGrantAdmin_();
  const input = getTransferPreparation(receipt);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let file;
  try {
    const props = PropertiesService.getScriptProperties();
    const folderId = props.getProperty(TRANSFER_FOLDER_KEY);
    const folder = folderId ? DriveApp.getFolderById(folderId) :
      DriveApp.createFolder('Grant Forms Cloud - 転記用確定データ');
    if (!folderId) props.setProperty(TRANSFER_FOLDER_KEY, folder.getId());
    const snapshot = {...input, id:Utilities.getUuid(), savedAt:new Date().toISOString(), state:'prepared'};
    file = folder.createFile(snapshot.id + '.json', JSON.stringify(snapshot, null, 2), MimeType.PLAIN_TEXT);
    file.setDescription(JSON.stringify({id:snapshot.id, company:input.application.company,
      receipt:input.application.receipt, grant:input.application.grant, savedAt:snapshot.savedAt, state:'prepared'}));
    return {id:snapshot.id, url:file.getUrl()};
  } catch (_) {
    if (file) { try { file.setTrashed(true); } catch (_) {} }
    throw new Error('転記用データを保存できませんでした。');
  } finally { lock.releaseLock(); }
}

function listTransferPreparations() {
  requireGrantAdmin_();
  const folderId = PropertiesService.getScriptProperties().getProperty(TRANSFER_FOLDER_KEY);
  if (!folderId) return [];
  const files = DriveApp.getFolderById(folderId).getFiles();
  const records = [];
  while (files.hasNext()) {
    const file = files.next();
    try {
      const record = JSON.parse(file.getDescription());
      if (record.id && record.state === 'prepared') records.push({...record, url:file.getUrl()});
    } catch (_) {}
  }
  return records.sort((a,b) => b.savedAt.localeCompare(a.savedAt));
}

function getSavedTransferPreparation(id) {
  requireGrantAdmin_();
  if (!/^[a-f0-9-]{36}$/.test(String(id || ''))) throw new Error('保存データを選び直してください。');
  const folderId = PropertiesService.getScriptProperties().getProperty(TRANSFER_FOLDER_KEY);
  if (folderId) {
    const files = DriveApp.getFolderById(folderId).getFilesByName(id + '.json');
    if (files.hasNext()) return JSON.parse(files.next().getBlob().getDataAsString('UTF-8'));
  }
  throw new Error('転記用データが見つかりません。');
}

/** Transfer workspaces are independent of source ledger, applications and original templates. */
function transferWorkspaceFolder_(){
 const p=PropertiesService.getScriptProperties();let id=p.getProperty('GRANT_TRANSFER_WORKSPACES_FOLDER_ID');
 if(id)return DriveApp.getFolderById(id);const f=DriveApp.createFolder('Grant Forms Cloud - 書類作成の準備');p.setProperty('GRANT_TRANSFER_WORKSPACES_FOLDER_ID',f.getId());return f;
}
function transferWorkspace_(id){
 if(!/^[a-f0-9-]{36}$/.test(String(id||'')))throw Error('作成データを選び直してください。');
 const rootId=PropertiesService.getScriptProperties().getProperty('GRANT_TRANSFER_WORKSPACES_FOLDER_ID');if(!rootId)throw Error('準備データが見つかりません。');
 const fs=DriveApp.getFolderById(rootId).getFoldersByName(id);if(!fs.hasNext())throw Error('準備データが見つかりません。');const folder=fs.next(),files=folder.getFilesByName('準備データ.json');if(!files.hasNext())throw Error('準備データが見つかりません。');const file=files.next();return {folder,file,data:JSON.parse(file.getBlob().getDataAsString('UTF-8'))};
}
function beginTransferWorkspace(input){
 requireGrantAdmin_();const grantName=String(input?.grantName||'').trim(),year=String(input?.year||'').trim(),notes=String(input?.notes||'').trim();
 if(!grantName||grantName.length>160||!year||year.length>40||notes.length>10000||!['original','original_pdf'].includes(input.output))throw Error('助成金名・年度・出力形式を確認してください。');
 if(!Array.isArray(input.files)||!input.files.length||input.files.length>20)throw Error('資料は1〜20件で登録してください。');
 let total=0;const files=input.files.map((f,index)=>{const name=String(f.name||'');if(!name||name.length>180||/[\\/\r\n]/.test(name)||!['blank','sample','guidelines'].includes(f.role)||!Number.isInteger(f.size)||f.size<1||f.size>15*1024*1024||!(/\.(docx?|xlsx?|pdf)$/i.test(name)))throw Error('ファイル名・種類・サイズを確認してください（1資料15 MBまで）。');if(f.role==='blank'&&!/\.(docx?|xlsx?)$/i.test(name))throw Error('空欄様式はWord・Excelで登録してください。');total+=f.size;return {index,name,role:f.role,size:f.size};});
 if(total>50*1024*1024||!files.some(f=>f.role==='blank'))throw Error('空欄のWord・Excel様式を含め、合計50 MB以内にしてください。');
 const record=companyLedgerRecords_().find(r=>r.id===input.companyId);if(!record)throw Error('対象の会社台帳を選択してください。');
 const details=companyLedgerDetails_(record),base={company:record.name,...record.profile},application=input.receipt?collectionExport_(input.receipt):null;
 if(application){const raw=PropertiesService.getScriptProperties().getProperties(),linked=Object.entries(raw).some(([key,value])=>{if(!key.startsWith('GRANT_INVITE_'))return false;try{const i=JSON.parse(value);return i.receipt===application.receipt&&i.companyId===record.id;}catch(_){return false;}});if(!linked&&application.company!==record.name)throw Error('選択した申請データと会社台帳の会社が一致しません。');}
 const lock=LockService.getScriptLock();lock.waitLock(20000);try{
 const id=Utilities.getUuid(),folder=transferWorkspaceFolder_().createFolder(id),data={id,state:'uploading',savedAt:new Date().toISOString(),companyId:record.id,company:record.name,grantName,year,output:input.output,notes,files,base:{values:{...(details.values||{}),...base},fields:details.fields||[],entries:record.entries||[],aliases:record.fieldAliases||{},employees:details.employees||[]},application,policy:{ledgerBase:true,keepOriginals:true,sampleReferenceOnly:true,missingValues:'review'},generationAvailable:false};
 folder.createFile('準備データ.json',JSON.stringify(data),MimeType.PLAIN_TEXT);return {id};
 }finally{lock.releaseLock();}
}
function uploadTransferWorkspaceFile(input){
 requireGrantAdmin_();const lock=LockService.getScriptLock();lock.waitLock(20000);try{
 const {folder,file,data}=transferWorkspace_(input?.id),slot=data.files[input.index];if(!slot||data.state!=='uploading')throw Error('資料の登録状態を確認してください。');if(slot.fileId)return {index:slot.index,name:slot.name};
 const encoded=String(input.base64||'');if(encoded.length>21*1024*1024||!/^[A-Za-z0-9+/]*={0,2}$/.test(encoded))throw Error('ファイルのデータを確認してください。');const bytes=Utilities.base64Decode(encoded);if(bytes.length!==slot.size)throw Error('ファイルサイズが一致しません。');
 const blob=Utilities.newBlob(bytes,'application/octet-stream',slot.name),ext=slot.name.split('.').pop().toLowerCase();requirementUploadExtension_({type:ext==='pdf'?'file':['doc','docx'].includes(ext)?'file_word':'file_excel',label:slot.name},blob,bytes);
 const name=String(slot.index+1).padStart(2,'0')+'_'+slot.name;const previous=folder.getFilesByName(name),saved=previous.hasNext()?previous.next():folder.createFile(blob.setName(name));slot.fileId=saved.getId();slot.url=saved.getUrl();file.setContent(JSON.stringify(data));return {index:slot.index,name:slot.name};
 }finally{lock.releaseLock();}
}
function completeTransferWorkspace(id){
 requireGrantAdmin_();const lock=LockService.getScriptLock();lock.waitLock(20000);try{const {file,data}=transferWorkspace_(id);if(data.files.some(f=>!f.fileId))throw Error('すべての資料を登録してください。');data.state='prepared';file.setContent(JSON.stringify(data));return data;}finally{lock.releaseLock();}
}
function listTransferWorkspaces(){
 requireGrantAdmin_();const id=PropertiesService.getScriptProperties().getProperty('GRANT_TRANSFER_WORKSPACES_FOLDER_ID');if(!id)return [];const folders=DriveApp.getFolderById(id).getFolders(),out=[];while(folders.hasNext()){const fs=folders.next().getFilesByName('準備データ.json');if(!fs.hasNext())continue;try{const d=JSON.parse(fs.next().getBlob().getDataAsString('UTF-8'));if(d.state==='prepared')out.push({id:d.id,company:d.company,grantName:d.grantName,year:d.year,savedAt:d.savedAt});}catch(_){}}return out.sort((a,b)=>b.savedAt.localeCompare(a.savedAt));
}
function getTransferWorkspace(id){requireGrantAdmin_();const {data}=transferWorkspace_(id);if(data.state!=='prepared')throw Error('資料の登録が完了していません。');return data;}
