/**
 * Grant Forms Cloud - initial Apps Script prototype.
 * Only basic organization details; do not collect sensitive personal data yet.
 */
const GRANT_FOLDER_ID = '1XJovIo8J46rq_VQl-u_pVdjzQ2cAjsm_';
const HEADER = ['受付日時','受付ID','法人名','担当者名','メールアドレス','電話番号','申請する奨励金','備考'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('助成金申請 事前受付フォーム')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function submitApplication(input) {
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
  const sheet = getResponseSheet_();
  Logger.log('保存先: ' + sheet.getParent().getUrl());
}
