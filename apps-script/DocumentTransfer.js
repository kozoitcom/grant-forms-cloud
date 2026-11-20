/** System 3 owns immutable transfer snapshots; never updates the collection or ledger. */
const TRANSFER_FOLDER_KEY = 'GRANT_TRANSFER_SNAPSHOTS_FOLDER_ID';

function getTransferPreparation(receipt) {
  requireGrantAdmin_();
  const application = collectionExport_(String(receipt || ''));
  // The only inputs from Systems 1 and 2 are explicit, serializable data contracts.
  const templates = listGrantDocumentSettings(application.grant).map(document => ({
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
