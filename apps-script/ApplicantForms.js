/** System 4: real applicant form previews and immutable invitation prefills. */
function applicantPrefill_(companyId, sourceReceipt) {
  requireGrantAdmin_();
  companyId = String(companyId || '');
  sourceReceipt = String(sourceReceipt || '');
  const record = companyId ? listCompanyNumbers().find(item => item.id === companyId) : null;
  if (companyId && !record) throw new Error('会社台帳から会社を選び直してください。');
  const prefill = {};
  const limits = {company:120, industry:120, address:250, representative:120,
    representativeKana:120, person:80, phone:40, email:180, note:1000};
  const copyValues = values => {
    for (const [key,max] of Object.entries(limits)) {
      if (values[key] !== undefined) prefill[key] = String(values[key] || '').trim().slice(0,max);
    }
  };
  if (record) { copyValues(record.profile || {}); prefill.company = record.name; }
  if (sourceReceipt) {
    const application = collectionExport_(sourceReceipt);
    if (record && record.name !== application.company) throw new Error('会社台帳と申請データの会社が一致しません。');
    copyValues(application.values);
    if(application.schemaVersion===2)for(const [key,value] of Object.entries(application.values)){if(/^[a-zA-Z][a-zA-Z0-9_]{0,63}$/.test(key))prefill[key]=String(value??'').slice(0,4000);}
    prefill.company = application.company;
    prefill.entity = application.entity;
    if (['yes','no'].includes(application.values.location)) prefill.location = application.values.location;
    if (['yes','no'].includes(application.values.otherCity)) prefill.otherCity = application.values.otherCity;
    prefill.employees = validateWomenR8Employees_(JSON.stringify(application.employees));
    for (const group of application.genderCounts?.groups || []) {
      if (!['workers','chiefs','managers','officers'].includes(group.key)) continue;
      for (const [suffix,value] of [['Male',group.male],['Female',group.female]]) {
        if (Number.isInteger(value) && value >= 0 && value <= 100000) prefill[group.key + suffix] = value;
      }
    }
  }
  if (prefill.representativeKana) prefill.representativeKana = prefill.representativeKana.normalize('NFKC')
    .replace(/[ァ-ヶヽヾ]/g, character => String.fromCharCode(character.charCodeAt(0) - 0x60));
  return prefill;
}

function getApplicantFormPreview(input) {
  requireGrantAdmin_();
  if (input && String(input.grant).startsWith('schema:')) {
    const definition = requirementRead_(input.grant.slice(7));
    applicantPrefill_(input.companyId,input.sourceReceipt);
    return {url:ScriptApp.getService().getUrl().replace(/\/dev$/, '/exec')+'?grant='+encodeURIComponent(input.grant)+'&preview=1&v='+definition.revision+(input.companyId?'&companyId='+encodeURIComponent(input.companyId):'')+(input.sourceReceipt?'&sourceReceipt='+encodeURIComponent(input.sourceReceipt):'')};
  }
  if (!input || input.grant !== 'women-r8') throw new Error('助成金を選び直してください。');
  // Resolve both identifiers now. Never put company details or private ledger entries in the URL.
  applicantPrefill_(input.companyId, input.sourceReceipt);
  const base = ScriptApp.getService().getUrl().replace(/\/dev$/, '/exec');
  return {url:base + '?grant=women-r8&preview=1' +
    (input.companyId ? '&companyId=' + encodeURIComponent(input.companyId) : '') +
    (input.sourceReceipt ? '&sourceReceipt=' + encodeURIComponent(input.sourceReceipt) : '')};
}

function saveApplicantPrefill_(prefill) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty('GRANT_INVITATION_PREFILL_FOLDER_ID');
    const folder = id ? DriveApp.getFolderById(id) : DriveApp.createFolder('Grant Forms Cloud - 案内時点の事前入力データ');
    if (!id) props.setProperty('GRANT_INVITATION_PREFILL_FOLDER_ID', folder.getId());
    return folder.createFile(Utilities.getUuid() + '.json', JSON.stringify(prefill), MimeType.PLAIN_TEXT).getId();
  } finally { lock.releaseLock(); }
}

function readApplicantInvitationPrefill_(invitation) {
  // Called only after readGrantInvitation_ has validated the bearer token and expiry.
  if (!invitation.prefillFileId) return invitation.prefill || {};
  return JSON.parse(DriveApp.getFileById(invitation.prefillFileId).getBlob().getDataAsString('UTF-8'));
}

/** The inline preview and actual applicant URL share one rendering implementation. */
function buildApplicantPage_(invitation, prefill, inviteToken, previewMode) {
  if(invitation?.legacyAnswerDraftId){const saved=JSON.parse(DriveApp.getFileById(invitation.legacyAnswerDraftId).getBlob().getDataAsString('UTF-8'));prefill={...prefill,...saved.values,company:saved.values.company||invitation.company,email:invitation.email};}
  const page = HtmlService.createTemplateFromFile('WomenR8');
  page.documents = WOMEN_R8_DOCUMENTS;
  page.documentSettings = getDocumentSettings_();
  page.invitation = invitation;
  page.prefill = prefill;
  page.inviteToken = inviteToken;
  page.previewMode = previewMode;
  page.initialFormJson = JSON.stringify({values:prefill,preview:previewMode}).replace(/</g, '\\u003c');
  return page.evaluate().setTitle('令和8年度 女性活躍情報公開促進奨励金 | 書類受付')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Partial applicant answers are private and only accessible with the validated invitation. */
function saveWomenR8AnswerDraft(input){
 const token=String(input?.token||''),invite=readGrantInvitation_(token);
 if(invite.grant!=='women-r8'||invite.receipt)throw new Error('このリンクは編集できません。');
 const raw=input.values||{},values={};
 const limits={company:120,industry:120,address:250,representative:120,representativeKana:120,person:80,phone:40,note:1000,entity:20,location:10,otherCity:10,workersMale:10,workersFemale:10,chiefsMale:10,chiefsFemale:10,managersMale:10,managersFemale:10,officersMale:10,officersFemale:10};
 for(const [key,max] of Object.entries(limits)){const v=String(raw[key]??'');if(v.length>max)throw new Error('入力内容が長すぎます。');values[key]=v;}
 if(!Array.isArray(raw.employees)||raw.employees.length>300)throw new Error('従業員数を確認してください。');
 values.employees=raw.employees.map(e=>{const out={};for(const key of ['name','position','employment']){out[key]=String(e[key]??'');if(out[key].length>120)throw new Error('従業員の入力内容が長すぎます。');}return out;});
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{const latest=readGrantInvitation_(token);if(latest.receipt)throw new Error('提出済みです。');const json=JSON.stringify({values,updatedAt:new Date().toISOString()});if(latest.legacyAnswerDraftId)DriveApp.getFileById(latest.legacyAnswerDraftId).setContent(json);else{latest.legacyAnswerDraftId=requirementStore_(true).createFile('legacy-answers-'+Utilities.getUuid()+'.json',json,MimeType.PLAIN_TEXT).getId();writeGrantInvitation_(token,latest);}return true;}finally{lock.releaseLock();}
}
