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
    const systems = [
      {id:'setup', title:'① 助成金・様式設定', file:'GrantSetupView'},
      {id:'applications', title:'② 申請データ・会社台帳', file:'ApplicantDataView'},
      {id:'transfer', title:'③ AI転記・書類作成', file:'DocumentTransferView'},
      {id:'forms', title:'④ 申請者フォーム・リンク発行', file:'ApplicantFormsView'}
    ];
    const selected = systems.find(item => item.id === params.system) || systems[0];
    const page = HtmlService.createTemplateFromFile('GrantAdmin');
    page.systems = systems;
    page.activeSystem = selected.id;
    page.systemFile = selected.file;
    page.adminBaseUrl = ScriptApp.getService().getUrl().replace(/\/dev$/, '/exec') + '?page=admin';
    return page.evaluate()
      .setTitle(selected.title + ' | Grant Forms Cloud')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (event && event.parameter && event.parameter.grant === 'women-r8') {
    const invitation = params.invite ? readGrantInvitation_(params.invite) : null;
    if (!invitation) requireGrantAdmin_();
    const previewMode = !invitation && params.preview === '1';
    const prefill = invitation ? readApplicantInvitationPrefill_(invitation) :
      previewMode ? applicantPrefill_(params.companyId, params.sourceReceipt) : {};
    return buildApplicantPage_(invitation, prefill, invitation ? params.invite : '', previewMode);
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

const GRANT_FORMS = {
  'women-r8': '令和8年度 女性活躍情報公開促進奨励金'
};

function requireGrantAdmin_() {
  const active = Session.getActiveUser().getEmail().toLowerCase();
  const owner = Session.getEffectiveUser().getEmail().toLowerCase();
  if (!active || active !== owner) throw new Error('管理者アカウントで開いてください。');
}

