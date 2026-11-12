/**
 * Grant Forms Cloud - initial Apps Script prototype.
 * Only basic organization details; do not collect sensitive personal data yet.
 */
const GRANT_FOLDER_ID = '1XJovIo8J46rq_VQl-u_pVdjzQ2cAjsm_';
const HEADER = ['受付日時','受付ID','法人名','担当者名','メールアドレス','電話番号','申請する奨励金','備考'];

function doGet(event) {
  const params = (event && event.parameter) || {};
  if (params.page === 'admin') {
    requireGrantAdmin_();
    return HtmlService.createHtmlOutputFromFile('GrantAdmin')
      .setTitle('Grant Forms Cloud 管理画面')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (event && event.parameter && event.parameter.grant === 'women-r8') {
    const page = HtmlService.createTemplateFromFile('WomenR8');
    page.documents = WOMEN_R8_DOCUMENTS;
    page.invitation = params.invite ? readGrantInvitation_(params.invite) : null;
    if (!page.invitation) requireGrantAdmin_();
    page.inviteToken = page.invitation ? params.invite : '';
    return page.evaluate()
      .setTitle('令和8年度 女性活躍情報公開促進奨励金 | 書類受付')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('助成金申請 事前受付フォーム')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function submitApplication(input) {
  requireGrantAdmin_();
  if (!input || typeof input !== 'object') throw new Error('入力内容がありません。');
  const value = (key, max) => String(input[key] || '').trim().slice(0,max);
  if (value('website',100)) throw new Error('送信できませんでした。');
  const company = value('company',120);
  const person = value('person',80);
  const email = value('email',180);
  const phone = value('phone',40);
  const grant = value('grant',150);
  const note = value('note',1000);
  if (!company || !person || !email || !grant) throw new Error('必須項目を入力してください。');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('メールアドレスを確認してください。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getResponseSheet_();
    const id = Utilities.getUuid();
    // Prefix text cells to prevent spreadsheet formula injection.
    const safe = s => /^[=+@-]/.test(s) ? "'" + s : s;
    sheet.appendRow([new Date(), id, safe(company), safe(person), safe(email), safe(phone), safe(grant), safe(note)]);
    SpreadsheetApp.flush();
    return {ok:true, receipt:id};
  } finally {
    lock.releaseLock();
  }
}

function getResponseSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('RESPONSE_SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id).getSheets()[0];
  const folder = DriveApp.getFolderById(GRANT_FOLDER_ID);
  const book = SpreadsheetApp.create('Grant Forms Cloud - 申請受付一覧');
  const file = DriveApp.getFileById(book.getId());
  file.moveTo(folder);
  const sheet = book.getSheets()[0];
  sheet.setName('受付一覧');
  sheet.appendRow(HEADER);
  sheet.setFrozenRows(1);
  props.setProperty('RESPONSE_SPREADSHEET_ID', book.getId());
  return sheet;
}

/** Run once from the editor to initialize storage and authorize Drive/Sheets. */
function setupStorage() {
  requireGrantAdmin_();
  const sheet = getResponseSheet_();
  Logger.log('保存先: ' + sheet.getParent().getUrl());
}

const WOMEN_R8_DOCUMENTS = [
  [
    "application",
    "支給申請書（様式第1号）",
    "all",
    "p.12"
  ],
  [
    "offices",
    "事業所一覧（様式第1号別紙）",
    "all",
    "p.12 · 本社・都外事業所も記載"
  ],
  [
    "pledge",
    "誓約書（様式第2号）",
    "all",
    "p.12 · 代表者の自署入り"
  ],
  [
    "insurance",
    "雇用保険被保険者資格取得等確認通知書（事業主通知用）",
    "all",
    "p.12 · 都内で常時雇用する労働者1名分"
  ],
  [
    "rules",
    "就業規則（本則）",
    "all",
    "p.12 · 直近の届出分。受付印が確認できる箇所を含める"
  ],
  [
    "profile",
    "会社案内または会社概要",
    "all",
    "p.12 · 所在地・代表者・事業内容が確認できるもの"
  ],
  [
    "registry",
    "履歴事項全部証明書",
    "company",
    "p.12 · 発行から3か月以内"
  ],
  [
    "opening",
    "個人事業の開業・廃業等届出書",
    "sole",
    "p.12"
  ],
  [
    "residence",
    "住民票記載事項証明書",
    "sole",
    "p.12 · 発行から3か月以内"
  ],
  [
    "locationProof",
    "水道光熱費の領収書または賃貸借契約書",
    "location",
    "pp.12–13 · 事業所の所在地が確認できるもの"
  ],
  [
    "corpCityTax",
    "法人都民税の納税証明書",
    "company",
    "p.13 · 納期確定の直近分"
  ],
  [
    "corpBusinessTax",
    "法人事業税の納税証明書",
    "company",
    "p.13 · 納期確定の直近分"
  ],
  [
    "soleCityTax",
    "個人都民税の納税証明書",
    "sole",
    "p.13 · 納期確定の直近分"
  ],
  [
    "soleOtherCityTax",
    "もう一方の所在地の個人都民税納税証明書",
    "otherCity",
    "p.13 · 居住地と事業所地が異なる場合"
  ],
  [
    "soleBusinessTax",
    "個人事業税の納税証明書",
    "sole",
    "p.13 · 納期確定の直近分"
  ],
  [
    "chart",
    "組織図",
    "all",
    "p.13 · 役職・雇用管理区分が分かるもの"
  ],
  [
    "gender",
    "男女別一覧",
    "all",
    "p.13 · 対象区分のいずれかで女性割合が4割未満と確認できるもの"
  ]
];
const WOMEN_R8_LIMIT = 20 * 1024 * 1024;

function submitWomenR8(form) {
  if (!form || form.website) throw new Error('送信できませんでした。');
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
  const uploads = WOMEN_R8_DOCUMENTS.filter(([, , group]) => applicable.has(group))
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
    folder = root.createFolder(receipt);
    uploads.forEach(([key, blob]) => folder.createFile(blob.setName(key + '.pdf')));
    const safe = text => /^[=+@-]/.test(text) ? "'" + text : text;
    sheet.appendRow([new Date(), receipt, corporate ? '法人' : '個人事業主',
      ...values.map(safe), corporate ? location : '', corporate ? '' : otherCity,
      uploads.length, folder.getUrl()]);
    recorded = true;
    SpreadsheetApp.flush();
    if (invitation) {
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
    '担当者名','メール','電話','備考','本店所在地条件','居住地と事業所地が異なる','添付数','書類フォルダ']);
  sheet.setFrozenRows(1);
  props.setProperties({WOMEN_R8_FOLDER_ID: root.getId(), WOMEN_R8_SHEET_ID: book.getId()});
  return {root, sheet};
}

const GRANT_FORMS = {
  'women-r8': '令和8年度 女性活躍情報公開促進奨励金'
};
const GRANT_INVITE_PREFIX = 'GRANT_INVITE_';

function requireGrantAdmin_() {
  const active = Session.getActiveUser().getEmail().toLowerCase();
  const owner = Session.getEffectiveUser().getEmail().toLowerCase();
  if (!active || active !== owner) throw new Error('管理者アカウントで開いてください。');
}

function grantInvitationKey_(token) {
  if (!/^[a-f0-9]{64}$/.test(token)) throw new Error('受付リンクが無効です。');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token);
  return GRANT_INVITE_PREFIX + Utilities.base64EncodeWebSafe(digest);
}

function readGrantInvitation_(token) {
  const raw = PropertiesService.getScriptProperties().getProperty(grantInvitationKey_(token));
  const invitation = raw ? JSON.parse(raw) : null;
  if (!invitation || invitation.grant !== 'women-r8' || Date.now() > invitation.expiresAt) {
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
    .map(([, value]) => JSON.parse(value)).sort((a, b) => b.createdAt - a.createdAt);
}

function createGrantInvitation(input) {
  requireGrantAdmin_();
  const company = String(input.company || '').trim();
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
    const json = JSON.stringify({id: recordId, name, entries});
    if (Utilities.newBlob(json).getBytes().length > 8500) throw new Error('入力内容が長すぎます。');
    props.setProperty(COMPANY_NUMBER_PREFIX + recordId, json);
    return recordId;
  } finally {
    lock.releaseLock();
  }
}
