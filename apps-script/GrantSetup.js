/** System 1: grant configuration and document templates. */
const GRANT_DRAFT_FOLDER = 'GRANT_DRAFT_FOLDER_ID';

function listGrantDrafts() {
  requireGrantAdmin_();
  const id = PropertiesService.getScriptProperties().getProperty(GRANT_DRAFT_FOLDER);
  if (!id) return [];
  const files = DriveApp.getFolderById(id).getFiles();
  const drafts = [];
  while (files.hasNext()) {
    const file = files.next();
    try {
      const metadata = JSON.parse(file.getDescription());
      if (metadata && metadata.title && metadata.year) drafts.push({...metadata, url: file.getUrl()});
    } catch (_) {
      // Files without draft metadata do not represent draft forms.
    }
  }
  return drafts.sort((a, b) => b.createdAt - a.createdAt);
}

function saveGrantDraft(form) {
  requireGrantAdmin_();
  const title = String(form.title || '').trim();
  const year = String(form.year || '').trim();
  const fields = String(form.fields || '').trim().slice(0, 4000);
  const documents = String(form.documents || '').trim().slice(0, 4000);
  const blob = form.guidelines;
  if (!title || title.length > 180 || !year || year.length > 32) {
    throw new Error('助成金名と年度を確認してください。');
  }
  if (!blob || typeof blob.getBytes !== 'function') throw new Error('募集要項PDFを選んでください。');
  const bytes = blob.getBytes();
  if (bytes.length > 20 * 1024 * 1024 || bytes.slice(0, 5).join(',') !== '37,80,68,70,45') {
    throw new Error('募集要項は20 MB以内のPDFを添付してください。');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let file;
  try {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty(GRANT_DRAFT_FOLDER);
    const folder = id ? DriveApp.getFolderById(id) : DriveApp.createFolder('Grant Forms Cloud - 募集要項・フォーム下書き');
    if (!id) props.setProperty(GRANT_DRAFT_FOLDER, folder.getId());
    file = folder.createFile(blob.setName(year + ' ' + title + ' 募集要項.pdf'));
    file.setDescription(JSON.stringify({title, year, fields, documents, createdAt: Date.now()}));
    return true;
  } catch (error) {
    if (file) {
      try { file.setTrashed(true); } catch (_) {}
    }
    throw new Error('下書きを保存できませんでした。');
  } finally {
    lock.releaseLock();
  }
}

function getGrantRequirements(grant) {
  requireGrantAdmin_();
  if (grant !== 'women-r8') throw new Error('この助成金の必要項目は未登録です。');
  const conditions = {
    all: '必須', company: '法人のみ必須', sole: '個人事業主のみ必須',
    location: '登記上の本店が都外、または本社機能のある事業所と所在地が異なる場合',
    otherCity: '個人事業主で居住地と事業所地が異なる場合'
  };
  return {
    fields: [
      '労働者・係長職・管理職・役員それぞれの男性人数・女性人数・必須（合計と女性割合を自動計算）',
      '事業者区分（法人／個人事業主）・必須',
      '企業名・屋号・必須',
      '業種・必須',
      '事業所所在地・必須',
      '代表者の役職・氏名・必須',
      '代表者のふりがな（ひらがな）・任意',
      '担当者名・必須',
      'メールアドレス・必須',
      '電話番号・任意',
      '登記上の本店の所在地条件・法人は回答必須',
      '居住地と事業所地が異なるか・個人事業主は回答必須',
      '備考・任意'
    ],
    documents: WOMEN_R8_DOCUMENTS.map(([, title, condition, note]) =>
      title + '【' + conditions[condition] + '】 ' + note)
  };
}

function getDocumentSettings_() {
  const props = PropertiesService.getScriptProperties();
  const fileId = props.getProperty('WOMEN_R8_DOCUMENT_SETTINGS_FILE_ID');
  if (fileId) return JSON.parse(DriveApp.getFileById(fileId).getBlob().getDataAsString('UTF-8'));
  const raw = props.getProperty('WOMEN_R8_DOCUMENT_SETTINGS');
  return raw ? JSON.parse(raw) : {};
}

function listGrantDocumentSettings(grant) {
  requireGrantAdmin_();
  if (grant !== 'women-r8') throw new Error('助成金を選び直してください。');
  const settings = getDocumentSettings_();
  return WOMEN_R8_DOCUMENTS.map(([key,title]) => ({key,title,...settings[key]}));
}

function saveGrantDocumentSettings(form) {
  requireGrantAdmin_();
  const key = String(form.documentKey || '');
  if (form.grant !== 'women-r8' || !WOMEN_R8_DOCUMENTS.some(doc => doc[0] === key)) {
    throw new Error('提出書類を選び直してください。');
  }
  const uploads = [];
  for (const type of ['blank','sample']) {
    const blob = form[type];
    if (!blob || typeof blob.getBytes !== 'function' || !blob.getName()) continue;
    const name = blob.getName();
    const bytes = blob.getBytes();
    const pdf = /\.pdf$/i.test(name) && bytes.slice(0,5).join(',') === '37,80,68,70,45';
    const office = /\.(docx|xlsx)$/i.test(name) && bytes[0] === 80 && bytes[1] === 75;
    if ((!pdf && !office) || bytes.length > 20 * 1024 * 1024) {
      throw new Error('PDF・Word（docx）・Excel（xlsx）を20 MB以内で登録してください。');
    }
    uploads.push({type,blob,name});
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  const created = [];
  try {
    const props = PropertiesService.getScriptProperties();
    const settings = getDocumentSettings_();
    const entry = settings[key] || {};
    if (uploads.length) {
      const folderId = props.getProperty('GRANT_DOCUMENT_TEMPLATE_FOLDER');
      const folder = folderId ? DriveApp.getFolderById(folderId) :
        DriveApp.createFolder('Grant Forms Cloud - 公式様式・管理者用見本');
      if (!folderId) props.setProperty('GRANT_DOCUMENT_TEMPLATE_FOLDER', folder.getId());
      for (const upload of uploads) {
        const file = folder.createFile(upload.blob);
        created.push(file);
        entry[upload.type] = {id:file.getId(),name:upload.name,url:file.getUrl()};
      }
    }
    entry.uploadNotRequired = form.uploadNotRequired === 'on' || form.uploadNotRequired === 'true';
    settings[key] = entry;
    writeDocumentSettings_(settings);
    return true;
  } catch (error) {
    created.forEach(file => { try { file.setTrashed(true); } catch (_) {} });
    throw new Error('書類設定を保存できませんでした。');
  } finally {
    lock.releaseLock();
  }
}


/** Keep legacy property reads compatible while supporting larger sets of templates. Caller holds the script lock. */
function writeDocumentSettings_(settings) {
  const props = PropertiesService.getScriptProperties();
  const json = JSON.stringify(settings);
  const fileId = props.getProperty('WOMEN_R8_DOCUMENT_SETTINGS_FILE_ID');
  if (fileId) { DriveApp.getFileById(fileId).setContent(json); return; }
  if (Utilities.newBlob(json).getBytes().length <= 8000) {
    props.setProperty('WOMEN_R8_DOCUMENT_SETTINGS',json);return;
  }
  const folderId = props.getProperty('GRANT_DOCUMENT_TEMPLATE_FOLDER');
  const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.createFolder('Grant Forms Cloud - 公式様式・管理者用見本');
  if (!folderId) props.setProperty('GRANT_DOCUMENT_TEMPLATE_FOLDER',folder.getId());
  const file = folder.createFile('_document-settings.json',json,MimeType.PLAIN_TEXT);
  props.setProperty('WOMEN_R8_DOCUMENT_SETTINGS_FILE_ID',file.getId());
}
