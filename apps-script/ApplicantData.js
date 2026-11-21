/** System 2: company ledger, collection, invitations and submitted records. */
function submitWomenR8(form) {
  if (!form || form.website || form.previewMode === 'true') throw new Error('プレビューからは送信できません。');
  const industry = String(form.industry || '').trim();
  if (!industry || industry.length > 120) throw new Error('業種を入力してください。');
  const employees = validateWomenR8Employees_(form.employeesJson);
  const genderCounts = validateWomenR8Counts_(form);
  const fields = [['company',120],['address',250],['representative',120],
    ['person',80],['email',180],['phone',40],['note',1000]];
  const values = fields.map(([key, max]) => String(form[key] || '').trim().slice(0, max));
  const [company, address, representative, person, email] = values;
  const token = String(form.inviteToken || '');
  if (!token) requireGrantAdmin_();
  const {entity, location, otherCity} = form;
  const corporate = entity === 'company';
  if ((!corporate && entity !== 'sole') || !company || !address || !representative || !person
      || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)
      || !['yes','no'].includes(corporate ? location : otherCity)) {
    throw new Error('必須項目とメールアドレスを確認してください。');
  }
  const applicable = new Set(['all', entity]);
  if (corporate && location === 'yes') applicable.add('location');
  if (!corporate && otherCity === 'yes') applicable.add('otherCity');
  let total = 0;
  const documentSettings = getDocumentSettings_();
  const uploads = WOMEN_R8_DOCUMENTS.filter(([key, , group]) => applicable.has(group) && !documentSettings[key]?.uploadNotRequired)
    .map(([key, label]) => {
      const blob = form[key];
      if (!blob || typeof blob.getBytes !== 'function') throw new Error(label + 'を添付してください。');
      const bytes = blob.getBytes();
      total += bytes.length;
      if (total > WOMEN_R8_LIMIT) throw new Error('添付は合計20 MB以内にしてください。');
      if (bytes.slice(0, 5).join(',') !== '37,80,68,70,45') {
        throw new Error(label + 'はPDFで添付してください。');
      }
      return [key, blob];
    });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let folder;
  let recorded = false;
  try {
    const invitation = token ? readGrantInvitation_(token) : null;
    if (invitation && invitation.receipt) return {receipt: invitation.receipt};
    if (invitation && (company !== invitation.company || email.toLowerCase() !== invitation.email.toLowerCase())) {
      throw new Error('案内メールの会社名と宛先メールアドレスを確認してください。');
    }
    const {root, sheet} = getWomenR8Storage_();
    const receipt = Utilities.getUuid();
    const destinationId = PropertiesService.getScriptProperties().getProperty('GRANT_SUBMISSION_FOLDER_ID');
    const destination = destinationId ? DriveApp.getFolderById(destinationId) : root;
    if (destination.isTrashed()) throw new Error('保存先フォルダが削除されています。');
    folder = destination.createFolder(company + ' - 女性活躍R8 - ' + receipt);
    uploads.forEach(([key, blob]) => {
      const title = WOMEN_R8_DOCUMENTS.find(document => document[0] === key)[1];
      folder.createFile(blob.setName(title + '.pdf'));
    });
    folder.createFile('入力内容.json', JSON.stringify({
      receipt, submittedAt: new Date().toISOString(), company, industry, entity,
      ...Object.fromEntries(fields.map(([key], index) => [key, values[index]])),
      representativeKana: String(form.representativeKana || '').trim().slice(0, 120),
      location: corporate ? location : '', otherCity: corporate ? '' : otherCity,
      genderCounts, employees
    }, null, 2), MimeType.PLAIN_TEXT);
    const safe = text => /^[=+@-]/.test(text) ? "'" + text : text;
    if (!sheet.getRange(1, 15).getValue()) sheet.getRange(1, 15).setValue('代表者のふりがな');
    sheet.getRange(1, 16).setValue('男女別人数・女性割合判定');
    sheet.getRange(1, 17).setValue('従業員一覧');
    sheet.getRange(1, 18).setValue('業種');
    sheet.appendRow([new Date(), receipt, corporate ? '法人' : '個人事業主',
      ...values.map(safe), corporate ? location : '', corporate ? '' : otherCity,
      uploads.length, folder.getUrl(), safe(String(form.representativeKana || '').trim().slice(0, 120)), JSON.stringify(genderCounts), JSON.stringify(employees), safe(industry)]);
    recorded = true;
    SpreadsheetApp.flush();
    if (invitation) {
      invitation.genderCounts = genderCounts;
      invitation.receipt = receipt;
      invitation.folderUrl = folder.getUrl();
      writeGrantInvitation_(token, invitation);
    }
    return {receipt};
  } catch (error) {
    if (folder && !recorded) {
      try { folder.setTrashed(true); } catch (_) {}
    }
    throw new Error('書類を保存できませんでした。受付担当者にご連絡ください。');
  } finally {
    lock.releaseLock();
  }
}

function getWomenR8Storage_() {
  const props = PropertiesService.getScriptProperties();
  const rootId = props.getProperty('WOMEN_R8_FOLDER_ID');
  const sheetId = props.getProperty('WOMEN_R8_SHEET_ID');
  if (rootId && sheetId) {
    return {root: DriveApp.getFolderById(rootId), sheet: SpreadsheetApp.openById(sheetId).getSheets()[0]};
  }
  const root = DriveApp.createFolder('Grant Forms Cloud - 女性活躍情報公開促進奨励金（令和8年度）');
  const book = SpreadsheetApp.create('女性活躍情報公開促進奨励金（令和8年度）- 書類受付一覧');
  DriveApp.getFileById(book.getId()).moveTo(root);
  const sheet = book.getSheets()[0];
  sheet.appendRow(['受付日時','受付ID','事業者区分','企業名・屋号','所在地','代表者役職・氏名',
    '担当者名','メール','電話','備考','本店所在地条件','居住地と事業所地が異なる','添付数','書類フォルダ','代表者のふりがな']);
  sheet.setFrozenRows(1);
  props.setProperties({WOMEN_R8_FOLDER_ID: root.getId(), WOMEN_R8_SHEET_ID: book.getId()});
  return {root, sheet};
}

const GRANT_INVITE_PREFIX = 'GRANT_INVITE_';


function grantInvitationKey_(token) {
  if (!/^[a-f0-9]{64}$/.test(token)) throw new Error('受付リンクが無効です。');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token);
  return GRANT_INVITE_PREFIX + Utilities.base64EncodeWebSafe(digest);
}

function readGrantInvitation_(token) {
  const raw = PropertiesService.getScriptProperties().getProperty(grantInvitationKey_(token));
  const invitation = raw ? JSON.parse(raw) : null;
  if (!invitation || invitation.grant !== 'women-r8' || invitation.deletedAt || Date.now() > invitation.expiresAt) {
    throw new Error('受付リンクが無効か、期限が切れています。');
  }
  return invitation;
}

function writeGrantInvitation_(token, invitation) {
  PropertiesService.getScriptProperties().setProperty(grantInvitationKey_(token), JSON.stringify(invitation));
}

function listGrantInvitations() {
  requireGrantAdmin_();
  const properties = PropertiesService.getScriptProperties().getProperties();
  return Object.entries(properties).filter(([key]) => key.startsWith(GRANT_INVITE_PREFIX))
    .map(([key, value]) => ({...JSON.parse(value), historyId: key}))
    .sort((a, b) => b.createdAt - a.createdAt);
}

function createGrantInvitation(input) {
  requireGrantAdmin_();
  const companyId = String(input.companyId || '');
  const record = companyId ? listCompanyNumbers().find(item => item.id === companyId) : null;
  if (companyId && !record) throw new Error('会社台帳から会社を選び直してください。');
  const company = record ? record.name : String(input.company || '').trim();
  const email = String(input.email || '').trim();
  const grant = String(input.grant || '');
  if (!company || company.length > 120 || /[\r\n]/.test(company) || email.length > 180 ||
      !/^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(email) ||
      !Object.prototype.hasOwnProperty.call(GRANT_FORMS, grant)) {
    throw new Error('会社名・宛先メールアドレス・助成金を確認してください。');
  }
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  const createdAt = Date.now();
  const invitation = {company, email, grant, createdAt, expiresAt: createdAt + 30 * 86400000};
  const prefill = applicantPrefill_(companyId, input.sourceReceipt);
  if (prefill.company && prefill.company !== company) throw new Error('事前入力元の会社と案内先の会社が一致しません。');
  if (record) invitation.companyId = record.id;
  prefill.company = company;
  prefill.email = email;
  if (Utilities.newBlob(JSON.stringify(prefill)).getBytes().length > 5000) {
    invitation.prefillFileId = saveApplicantPrefill_(prefill);
  } else invitation.prefill = prefill;
  writeGrantInvitation_(token, invitation);
  return {token, ...grantInvitationMessage_(token, invitation)};
}

function grantInvitationMessage_(token, invitation) {
  const title = GRANT_FORMS[invitation.grant];
  const url = ScriptApp.getService().getUrl().replace(/\/dev$/, '/exec') +
    '?grant=' + encodeURIComponent(invitation.grant) + '&invite=' + encodeURIComponent(token);
  const deadline = Utilities.formatDate(new Date(invitation.expiresAt), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  return {email: invitation.email, url, subject: '【書類提出のご案内】' + title,
    body: invitation.company + ' ご担当者様\n\n' +
      title + 'の申請書類を、以下の貴社専用フォームからご提出ください。\n\n' +
      url + '\n\n受付リンクの有効期限 ' + deadline + '（日本時間）\n' +
      'フォームには、このメールの宛先と同じメールアドレスを入力してください。\n' +
      'このリンクは関係者以外に共有しないでください。'};
}

function sendGrantInvitation(token) {
  requireGrantAdmin_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const invitation = readGrantInvitation_(token);
    if (invitation.sentAt) return {sentAt: invitation.sentAt};
    if (invitation.sending) throw new Error('送信結果を確認してください。二重送信を防ぐため再送を停止しました。');
    const message = grantInvitationMessage_(token, invitation);
    invitation.sending = true;
    writeGrantInvitation_(token, invitation);
    MailApp.sendEmail(message.email, message.subject, message.body);
    invitation.sentAt = Date.now();
    delete invitation.sending;
    writeGrantInvitation_(token, invitation);
    return {sentAt: invitation.sentAt};
  } finally {
    lock.releaseLock();
  }
}

const COMPANY_NUMBER_PREFIX = 'COMPANY_NUMBERS_';

function listCompanyNumbers() {
  requireGrantAdmin_();
  const data = PropertiesService.getScriptProperties().getProperties();
  return Object.entries(data).filter(([key]) => key.startsWith(COMPANY_NUMBER_PREFIX))
    .map(([, value]) => JSON.parse(value)).sort((a, b) => a.name.localeCompare(b.name, 'ja'));
}

function saveCompanyNumbers(input) {
  requireGrantAdmin_();
  const name = String(input.name || '').trim();
  const id = String(input.id || '');
  if (!name || name.length > 120 || (id && !/^[a-f0-9-]{36}$/.test(id)) ||
      !Array.isArray(input.entries) || input.entries.length > 20) {
    throw new Error('会社名を入力し、項目は20件以内にしてください。');
  }
  const labels = new Set();
  const entries = input.entries.map(([rawLabel, rawValue]) => {
    const label = String(rawLabel || '').trim();
    const value = String(rawValue || '').trim();
    if (!label || label.length > 80 || value.length > 120 || labels.has(label)) {
      throw new Error('項目名と入力値の長さ、項目名の重複を確認してください。');
    }
    labels.add(label);
    return [label, value];
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    const records = listCompanyNumbers();
    if (id && !records.some(record => record.id === id)) throw new Error('会社の登録が見つかりません。');
    if (records.some(record => record.id !== id && record.name === name)) {
      throw new Error('同名の会社は登録済みです。会社一覧から選択してください。');
    }
    const recordId = id || Utilities.getUuid();
    const profile = Object.fromEntries(Object.entries({"industry":["業種",120],"address":["事業所所在地",250],"representative":["代表者の役職・氏名",120],"representativeKana":["代表者のふりがな",120],"person":["担当者名",80],"phone":["電話番号",40]})
      .map(([key, [, max]]) => [key, String(input.profile?.[key] || '').trim().slice(0, max)]));
    const json = JSON.stringify({id: recordId, name, entries, profile});
    if (Utilities.newBlob(json).getBytes().length > 8500) throw new Error('入力内容が長すぎます。');
    props.setProperty(COMPANY_NUMBER_PREFIX + recordId, json);
    return recordId;
  } finally {
    lock.releaseLock();
  }
}

function setGrantHistoryDeleted(historyId, deleted) {
  requireGrantAdmin_();
  if (typeof historyId !== 'string' ||
      !/^GRANT_INVITE_[A-Za-z0-9_-]{43}=?$/.test(historyId) ||
      typeof deleted !== 'boolean') throw new Error('履歴を選び直してください。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(historyId);
    if (!raw) throw new Error('履歴が見つかりません。');
    const invitation = JSON.parse(raw);
    if (deleted) invitation.deletedAt = Date.now();
    else delete invitation.deletedAt;
    props.setProperty(historyId, JSON.stringify(invitation));
    return true;
  } finally {
    lock.releaseLock();
  }
}

function validateWomenR8Counts_(form) {
  const groups = [["workers","労働者"],["chiefs","係長職"],["managers","管理職"],["officers","役員"]];
  const counts = groups.map(([key, label]) => {
    const values = ['Male', 'Female'].map(suffix => {
      const raw = String(form[key + suffix] ?? '').trim();
      if (!/^\d+$/.test(raw) || Number(raw) > 100000) {
        throw new Error(label + 'の男女別人数を0以上の整数で入力してください。');
      }
      return Number(raw);
    });
    const [male, female] = values;
    const total = male + female;
    return {key, label, male, female, total, below40: total > 0 && female * 5 < total * 2};
  });
  return {groups: counts, eligible: counts.some(group => group.below40)};
}

function validateWomenR8Employees_(raw) {
  let employees;
  try { employees = JSON.parse(String(raw || '[]')); }
  catch (_) { throw new Error('従業員一覧を確認してください。'); }
  if (!Array.isArray(employees)) throw new Error('従業員一覧を確認してください。');
  const cleaned = employees.map(employee => {
    const name = String(employee?.name || '').trim();
    const position = String(employee?.position || '').trim();
    const employment = String(employee?.employment || '');
    if (!name || name.length > 120 || !position || position.length > 120 ||
        !['regular', 'nonregular'].includes(employment)) {
      throw new Error('従業員の氏名・役職・雇用区分を確認してください。');
    }
    return {name, position, employment};
  });
  if (JSON.stringify(cleaned).length > 45000) throw new Error('従業員一覧の入力内容が長すぎます。');
  return cleaned;
}

/** Admin-only setting; never accepts a destination from applicant input. */
function getSubmissionStorageSettings() {
  requireGrantAdmin_();
  const id = PropertiesService.getScriptProperties().getProperty('GRANT_SUBMISSION_FOLDER_ID');
  if (!id) return {url:'', name:'標準の保存先（助成金別フォルダ）'};
  const folder = DriveApp.getFolderById(id);
  return {url:folder.getUrl(), name:folder.getName()};
}

function saveSubmissionStorageSettings(input) {
  requireGrantAdmin_();
  const value = String(input || '').trim();
  const match = value.match(/^https:\/\/drive\.google\.com\/drive\/(?:u\/\d+\/)?folders\/([A-Za-z0-9_-]+)(?:[?/#].*)?$/);
  const id = match ? match[1] : (/^[A-Za-z0-9_-]{10,}$/.test(value) ? value : '');
  if (!id) throw new Error('GoogleドライブのフォルダURLを入力してください。');
  const folder = DriveApp.getFolderById(id);
  if (folder.isTrashed()) throw new Error('削除済みのフォルダは指定できません。');
  // Confirm actual write access before storing the destination; do not change permissions.
  const probe = folder.createFile('Grant Forms Cloud - 保存先確認.txt', '保存先の書き込み確認', MimeType.PLAIN_TEXT);
  probe.setTrashed(true);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    PropertiesService.getScriptProperties().setProperty('GRANT_SUBMISSION_FOLDER_ID', id);
    return {url:folder.getUrl(), name:folder.getName()};
  } finally { lock.releaseLock(); }
}


/** Read-only data contract exported by System 2. No storage is created on reads. */
function collectionRows_() {
  const id = PropertiesService.getScriptProperties().getProperty('WOMEN_R8_SHEET_ID');
  if (!id) return [];
  return SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues().slice(1)
    .filter(row => row[1] && row[3]);
}

function collectionMetadata_(row) {
  const unescape = value => String(value || '').replace(/^'(?=[=+@-])/, '');
  const date = row[0] instanceof Date ? row[0] : new Date(row[0]);
  return {receipt:String(row[1]), grant:'women-r8', company:unescape(row[3]),
    submittedAt:isNaN(date.getTime()) ? '' : date.toISOString(), folderUrl:String(row[13] || '')};
}

function listCollectedApplications() {
  requireGrantAdmin_();
  return collectionRows_().map(collectionMetadata_).reverse();
}

function collectionExport_(receipt) {
  if (!/^[a-f0-9-]{36}$/.test(String(receipt || ''))) throw new Error('申請を選び直してください。');
  const row = collectionRows_().find(item => String(item[1]) === receipt);
  if (!row) throw new Error('保存済みの申請データが見つかりません。');
  const text = value => String(value || '').replace(/^'(?=[=+@-])/, '');
  const json = (value, fallback) => {
    if (!value) return fallback;
    try { return JSON.parse(value); } catch (_) { throw new Error('申請データの形式を確認してください。'); }
  };
  return {schemaVersion:1, ...collectionMetadata_(row), entity:row[2] === '法人' ? 'company' : 'sole',
    values:{company:text(row[3]), address:text(row[4]), representative:text(row[5]),
      person:text(row[6]), email:text(row[7]), phone:text(row[8]), note:text(row[9]),
      location:text(row[10]), otherCity:text(row[11]), representativeKana:text(row[14]), industry:text(row[17])},
    genderCounts:json(row[15], null), employees:json(row[16], [])};
}
