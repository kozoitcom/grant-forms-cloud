/** Administrator-only bulk material upload. Classification is confirmed by the user. */
const MATERIAL_BATCH_ROOT = 'GRANT_MATERIAL_BATCH_ROOT_ID';
const MATERIAL_FILE_LIMIT = 20 * 1024 * 1024;
const MATERIAL_BATCH_LIMIT = 50 * 1024 * 1024;
const MATERIAL_BATCH_COUNT = 30;

function getGrantMaterialCatalog() {
  requireGrantAdmin_();
  return {grants:[{id:'women-r8', title:GRANT_FORMS['women-r8'], year:'令和8年度'}],
    documents:WOMEN_R8_DOCUMENTS.map(([key,title]) => ({key,title})),
    maxFileBytes:MATERIAL_FILE_LIMIT,maxBatchBytes:MATERIAL_BATCH_LIMIT,maxFiles:MATERIAL_BATCH_COUNT};
}

function materialMetadata_(input) {
  if (!input || !/^[a-f0-9-]{36}$/.test(String(input.requestId || ''))) throw new Error('アップロードを最初から選び直してください。');
  const grant = String(input.grant || '');
  if (grant && grant !== 'women-r8') throw new Error('登録先の助成金を選び直してください。');
  const title = grant ? GRANT_FORMS[grant] : String(input.title || '').trim();
  const year = grant ? '令和8年度' : String(input.year || '').trim();
  if (!title || title.length > 180 || !year || year.length > 32 || /[\r\n]/.test(title + year)) throw new Error('助成金名と年度を入力してください。');
  if (!Array.isArray(input.files) || !input.files.length || input.files.length > MATERIAL_BATCH_COUNT) throw new Error('ファイルは1〜30件で選択してください。');
  let total = 0;
  const pairs = new Set();
  const files = input.files.map(item => {
    if (!item || typeof item !== 'object') throw new Error('ファイル情報を確認してください。');
    const name = String(item.name || '');
    const extension = name.match(/\.(pdf|docx|xlsx)$/i)?.[1].toLowerCase();
    const role = String(item.role || '');
    const documentKey = String(item.documentKey || '');
    const known = WOMEN_R8_DOCUMENTS.find(document => document[0] === documentKey);
    const documentTitle = known ? known[1] : String(item.documentTitle || '').trim();
    if (!name || name.length > 240 || /[\r\n\x00]/.test(name) || !extension ||
        !Number.isInteger(item.size) || item.size <= 0 || item.size > MATERIAL_FILE_LIMIT) throw new Error('PDF・Word（docx）・Excel（xlsx）を1件20 MB以内で選択してください。');
    if (!['guidelines','blank','sample'].includes(role)) throw new Error(name + 'の種類を選択してください。');
    if (role === 'guidelines' && extension !== 'pdf') throw new Error('募集要項はPDFで登録してください。');
    if (role !== 'guidelines' && (!documentTitle || documentTitle.length > 180 || (documentKey && (!grant || !known)))) throw new Error(name + 'に対応する書類を選択してください。');
    if (role !== 'guidelines' && documentKey) {
      const pair = role + ':' + documentKey;
      if (pairs.has(pair)) throw new Error('同じ提出書類の同じ種類は1ファイルにまとめてください。');
      pairs.add(pair);
    }
    total += item.size;
    return {name,size:item.size,role,documentKey:role === 'guidelines' ? '' : documentKey,
      documentTitle:role === 'guidelines' ? '募集要項' : documentTitle};
  });
  if (total > MATERIAL_BATCH_LIMIT) throw new Error('1回の合計を50 MB以内にしてください。');
  return {id:input.requestId,grant,title,year,files};
}

function materialRoot_(create) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(MATERIAL_BATCH_ROOT);
  if (id) return DriveApp.getFolderById(id);
  if (!create) return null;
  const folder = DriveApp.createFolder('Grant Forms Cloud - 助成金資料の一括登録');
  props.setProperty(MATERIAL_BATCH_ROOT,folder.getId());
  return folder;
}

function materialBatch_(id) {
  if (!/^[a-f0-9-]{36}$/.test(String(id || ''))) throw new Error('アップロード情報を確認してください。');
  const root = materialRoot_(false);
  const folders = root ? root.getFoldersByName(id) : null;
  if (!folders || !folders.hasNext()) throw new Error('アップロード情報が見つかりません。');
  const folder = folders.next();
  const manifests = folder.getFilesByName('_batch.json');
  if (!manifests.hasNext()) throw new Error('アップロード情報が見つかりません。');
  const file = manifests.next();
  const manifest = JSON.parse(file.getBlob().getDataAsString('UTF-8'));
  return {folder,file,manifest};
}

function beginGrantMaterialBatch(input) {
  requireGrantAdmin_();
  const metadata = materialMetadata_(input);
  const signature = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(metadata)));
  const lock = LockService.getScriptLock();lock.waitLock(20000);
  let created;
  try {
    const root = materialRoot_(true);
    const existing = root.getFoldersByName(metadata.id);
    if (existing.hasNext()) {
      const batch = materialBatch_(metadata.id);
      if (batch.manifest.signature !== signature) throw new Error('選択ファイルが変更されています。新しい一括登録を開始してください。');
      return {id:metadata.id,state:batch.manifest.state};
    }
    created = root.createFolder(metadata.id);
    const manifest = {schemaVersion:1,...metadata,signature,state:'uploading',createdAt:new Date().toISOString()};
    created.createFile('_batch.json',JSON.stringify(manifest),MimeType.PLAIN_TEXT);
    return {id:metadata.id,state:'uploading'};
  } catch (error) {
    if (created) { try { created.setTrashed(true); } catch (_) {} }
    throw error;
  } finally {lock.releaseLock();}
}

function uploadGrantMaterialFile(input) {
  requireGrantAdmin_();
  if (!input || !Number.isInteger(input.index) || typeof input.base64 !== 'string' ||
      input.base64.length > Math.ceil(MATERIAL_FILE_LIMIT / 3) * 4 || !/^[A-Za-z0-9+/]*={0,2}$/.test(input.base64)) throw new Error('アップロードするファイルを確認してください。');
  const lock = LockService.getScriptLock();lock.waitLock(20000);
  let created;
  try {
    const batch = materialBatch_(input.id), manifest = batch.manifest;
    const item = manifest.files[input.index];
    if (!item) throw new Error('対象ファイルが見つかりません。');
    const bytes = Utilities.base64Decode(input.base64);
    if (bytes.length !== item.size || bytes.length > MATERIAL_FILE_LIMIT) throw new Error(item.name + 'のファイルサイズが一致しません。');
    const extension = item.name.match(/\.(pdf|docx|xlsx)$/i)[1].toLowerCase();
    const valid = extension === 'pdf' ? bytes.slice(0,5).join(',') === '37,80,68,70,45' : bytes[0] === 80 && bytes[1] === 75 && bytes[2] === 3 && bytes[3] === 4;
    if (!valid) throw new Error(item.name + 'のファイル形式を確認してください。');
    const digest = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,bytes));
    if (item.fileId) {
      if (item.digest !== digest) throw new Error('同じ登録のファイル内容は変更できません。');
      return {index:input.index,name:item.name};
    }
    if (manifest.state !== 'uploading') throw new Error('完了した登録には追加できません。新しい一括登録を開始してください。');
    const mime = {pdf:'application/pdf',docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',xlsx:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}[extension];
    const filename = (input.index + 1) + '_' + item.name.replace(/[\\/\x00-\x1f]/g,'_');
    created = batch.folder.createFile(Utilities.newBlob(bytes,mime,filename));
    Object.assign(item,{fileId:created.getId(),url:created.getUrl(),digest});
    batch.file.setContent(JSON.stringify(manifest));
    return {index:input.index,name:item.name};
  } catch (error) {
    if (created) { try { created.setTrashed(true); } catch (_) {} }
    throw error;
  } finally {lock.releaseLock();}
}

function completeGrantMaterialBatch(id) {
  requireGrantAdmin_();
  const lock = LockService.getScriptLock();lock.waitLock(20000);
  try {
    const batch = materialBatch_(id), manifest = batch.manifest;
    if (manifest.state === 'completed') return {id,grant:manifest.grant,url:batch.folder.getUrl()};
    if (manifest.files.some(item => !item.fileId)) throw new Error('まだアップロードされていないファイルがあります。再試行してください。');
    for (const item of manifest.files) if (DriveApp.getFileById(item.fileId).isTrashed()) throw new Error('削除されたファイルがあります。新しい登録を開始してください。');
    if (manifest.grant === 'women-r8') {
      const settings = getDocumentSettings_();
      let changed = false;
      for (const item of manifest.files) {
        if (item.role === 'guidelines' || !item.documentKey) continue;
        const entry = settings[item.documentKey] || {};
        entry[item.role] = {id:item.fileId,name:item.name,url:item.url};
        settings[item.documentKey] = entry;changed = true;
      }
      if (changed) writeDocumentSettings_(settings);
    }
    manifest.state = 'completed';manifest.completedAt = new Date().toISOString();
    batch.file.setContent(JSON.stringify(manifest));
    return {id,grant:manifest.grant,url:batch.folder.getUrl()};
  } finally {lock.releaseLock();}
}

function listGrantMaterialBatches() {
  requireGrantAdmin_();
  const root = materialRoot_(false);
  if (!root) return [];
  const folders = root.getFolders(), results = [];
  while (folders.hasNext()) {
    const folder = folders.next(), manifests = folder.getFilesByName('_batch.json');
    if (!manifests.hasNext()) continue;
    try {
      const batch = JSON.parse(manifests.next().getBlob().getDataAsString('UTF-8'));
      results.push({id:batch.id,title:batch.title,year:batch.year,grant:batch.grant,state:batch.state,
        createdAt:batch.createdAt,url:folder.getUrl(),files:batch.files.map(item=>({name:item.name,role:item.role,documentTitle:item.documentTitle,uploaded:Boolean(item.fileId),url:item.url || ''}))});
    } catch (_) {}
  }
  return results.sort((a,b)=>b.createdAt.localeCompare(a.createdAt));
}
