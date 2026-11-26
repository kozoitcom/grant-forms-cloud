const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');
const app = path.join(__dirname,'../apps-script');
const read = file => fs.readFileSync(path.join(app,file),'utf8');
const files = ['コード.js','GrantSetup.js','ApplicantData.js','DocumentTransfer.js','ApplicantForms.js'];
const code = files.map(read).join('\n');
function harness() {
  let n=0;
  const props = {WOMEN_R8_SHEET_ID:'sheet',COMPANY_NUMBERS_KEEP:'preserved',WOMEN_R8_DOCUMENT_SETTINGS:'{"rules":{"sample":{"id":"template-1","name":"sample.docx","url":"https://drive.google.com/file/d/template-1/view"}}}'};
  const rows = [['header'],[new Date('2026-11-20T00:00:00Z'),'12345678-1234-1234-1234-123456789abc','法人','会社A','住所A','代表A','担当A','test@example.test','00-0000-0000','備考','no','',1,'https://drive.google.com/drive/folders/submission','かな','{"groups":[],"eligible":false}','[{"name":"従業員A","position":"なし","employment":"regular"}]','通信業']];
  const saved=[];
  const iterator = items => {let i=0;return {hasNext:()=>i<items.length,next:()=>items[i++]};};
  const folder={createFile:(name,body)=>{const file={getId:()=> 'prefill-file-'+saved.length,name,body,description:'',getUrl:()=> 'https://drive.google.com/file/d/snapshot/view',setDescription(value){this.description=value;},getDescription(){return this.description;},getBlob(){return {getDataAsString:()=>this.body}},setTrashed(){}};saved.push(file);return file;},getId:()=> 'transfer-folder',getFiles:()=>iterator(saved),getFilesByName:name=>iterator(saved.filter(f=>f.name===name))};
  let currentEmail='admin@example.test';
  const context={Date,console,
    Session:{getActiveUser:()=>({getEmail:()=>currentEmail}),getEffectiveUser:()=>({getEmail:()=> 'admin@example.test'})},
    PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k],getProperties:()=>({...props}),setProperty:(k,v)=>props[k]=v})},
    SpreadsheetApp:{openById:()=>({getSheets:()=>[{getDataRange:()=>({getValues:()=>rows})}]})},
    DriveApp:{createFolder:()=>folder,getFolderById:()=>folder,getFileById:id=>saved.find(file=>file.getId()===id)},
    LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
    Utilities:{formatDate:date=>date.toISOString(),DigestAlgorithm:{SHA_256:'sha256'},computeDigest:(algorithm,text)=>crypto.createHash(algorithm).update(text).digest(),base64EncodeWebSafe:bytes=>Buffer.from(bytes).toString('base64url'),newBlob:text=>({getBytes:()=>Buffer.from(text)}),getUuid:()=> 'aaaaaaaa-aaaa-aaaa-aaaa-'+String(++n).padStart(12,'0')},MimeType:{PLAIN_TEXT:'text/plain'},
    ScriptApp:{getService:()=>({getUrl:()=> 'https://script.google.com/macros/s/deployment/exec'})},
    HtmlService:{createTemplateFromFile:()=>({evaluate(){return {setTitle(){return this},addMetaTag(){return this},getContent(){return '<html>preview</html>'},template:this}}})}
  };
  vm.createContext(context);vm.runInContext(code,context);
  return {context,props,rows,saved,setUser:email=>currentEmail=email};
}
test('each system loads its own view; unknown route safely selects setup',()=>{
  const {context}=harness();
  for(const [system,view] of [['setup','GrantSetupView'],['applications','ApplicantDataView'],['transfer','DocumentTransferView'],['forms','GrantSetupView'],['<script>','GrantSetupView']]) {
    const page=context.doGet({parameter:{page:'admin',system}}).template;
    assert.equal(page.systemFile,view);
    assert.equal(page.systems.length,3);
  }
});
test('views and scripts are separate and do not include other systems form controls',()=>{
  for(const file of ['GrantSetupView.html','ApplicantDataView.html','DocumentTransferView.html','ApplicantFormsView.html']) {
    const html=read(file);new vm.Script(html.match(/<script>([\s\S]*?)<\/script>/)[1],{filename:file});
    const ids=[...html.matchAll(/\bid="([^"]+)"/g)].map(m=>m[1]);
    assert.equal(new Set(ids).size,ids.length,'duplicate DOM ids');
  }
  assert(!read('ApplicantDataView.html').includes('id="document-settings"'));
  assert(read('ApplicantDataView.html').includes('id="company-data"'));
  assert(!read('GrantSetupView.html').includes('id="company-data"'));
  assert(!read('DocumentTransferView.html').includes('saveCompanyNumbers'));
});
test('preparing transfer reads submitted values and templates without changing source data',()=>{
  const {context,props,rows,saved}=harness();const original=JSON.stringify({props,rows});
  const applications=context.listCollectedApplications();assert.equal(applications.length,1);
  const result=context.getTransferPreparation(applications[0].receipt);
  assert.equal(result.application.values.company,'会社A');assert.equal(result.application.values.industry,'通信業');
  assert.equal(result.application.employees[0].name,'従業員A');
  assert.equal(result.templates.find(d=>d.key==='rules').sample.id,'template-1');
  assert.equal(result.generationAvailable,false);assert.equal(saved.length,0);
  assert.equal(JSON.stringify({props,rows}),original);
});
test('saved transfer snapshot survives later changes to ledger, submission and templates',()=>{
  const {context,props,rows}=harness();const receipt=rows[1][1];
  const result=context.saveTransferPreparation(receipt);
  rows[1][3]='会社B';props.COMPANY_NUMBERS_KEEP='changed';props.WOMEN_R8_DOCUMENT_SETTINGS='{}';
  const snapshot=context.getSavedTransferPreparation(result.id);
  assert.equal(snapshot.application.company,'会社A');
  assert.equal(snapshot.templates.find(d=>d.key==='rules').sample.id,'template-1');
  assert.equal(context.listTransferPreparations().length,1);
  assert.equal(snapshot.state,'prepared');
});
test('all new administrative data endpoints reject non-admin; invalid receipt rejected',()=>{
  const {context,setUser}=harness();
  assert.throws(()=>context.getTransferPreparation('not-a-receipt'));
  setUser('other@example.test');
  for(const name of ['listCollectedApplications','getTransferPreparation','saveTransferPreparation','listTransferPreparations','getSavedTransferPreparation'])assert.throws(()=>context[name]('12345678-1234-1234-1234-123456789abc'));
});
test('empty collection returns no records and never creates new storage',()=>{
  const {context,props,saved}=harness();delete props.WOMEN_R8_SHEET_ID;
  assert.equal(context.listCollectedApplications().length,0);assert.equal(saved.length,0);
});

test('form preview loads the saved application and never copies internal ledger numbers',()=>{
  const {context,rows,props}=harness();
  props.COMPANY_NUMBERS_KEEP=JSON.stringify({id:'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa',name:'会社A',profile:{industry:'旧業種',address:'旧住所'},entries:[['管理者番号','DO-NOT-EXPOSE']]});
  const input={grant:'women-r8',companyId:'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa',sourceReceipt:rows[1][1]};
  const preview=context.getApplicantFormPreview(input);
  assert(preview.url.includes('preview=1'));assert(!preview.url.includes('会社A'));
  const page=context.doGet({parameter:{grant:'women-r8',preview:'1',companyId:input.companyId,sourceReceipt:input.sourceReceipt}}).template;
  assert.equal(page.previewMode,true);assert.equal(page.prefill.company,'会社A');assert.equal(page.prefill.industry,'通信業');
  assert.equal(page.prefill.employees[0].name,'従業員A');assert(!page.initialFormJson.includes('DO-NOT-EXPOSE'));
  assert.throws(()=>context.submitWomenR8({previewMode:'true'}));
});
test('preview rejects another company source and is restricted to administrators',()=>{
  const {context,rows,props,setUser}=harness();
  props.COMPANY_NUMBERS_KEEP=JSON.stringify({id:'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa',name:'別会社',profile:{},entries:[]});
  assert.throws(()=>context.getApplicantFormPreview({grant:'women-r8',companyId:'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa',sourceReceipt:rows[1][1]}));
  setUser('other@example.test');
  assert.throws(()=>context.getApplicantFormPreview({grant:'women-r8'}));
  assert.throws(()=>context.doGet({parameter:{grant:'women-r8',preview:'1'}}));
});
test('initial form JSON safely escapes markup in company input',()=>{
  const {context,props}=harness();
  props.COMPANY_NUMBERS_KEEP=JSON.stringify({id:'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa',name:'</script><script>alert(1)</script>',profile:{},entries:[]});
  const page=context.doGet({parameter:{grant:'women-r8',preview:'1',companyId:'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'}}).template;
  assert(!page.initialFormJson.includes('<'));
  assert.equal(JSON.parse(page.initialFormJson).values.company,'</script><script>alert(1)</script>');
});
test('fourth screen owns invitation form; second screen keeps ledger and history only',()=>{
  const collect=read('ApplicantDataView.html'),forms=read('ApplicantFormsView.html');
  assert(!collect.includes('id="issue"'));assert(!collect.includes('createGrantInvitation('));
  assert(forms.includes('id="issue"'));assert(forms.includes('getApplicantFormPreview('));
  assert(collect.includes('id="company-data"'));assert(collect.includes('id="history"'));
});

test('issued link snapshots prefilled data and uses the actual recipient email',()=>{
  const {context,rows}=harness();
  const result=context.createGrantInvitation({company:'会社A',email:'recipient@example.test',grant:'women-r8',sourceReceipt:rows[1][1]});
  const invitation=context.readGrantInvitation_(result.token);
  assert.equal(invitation.email,'recipient@example.test');
  assert.equal(context.readApplicantInvitationPrefill_(invitation).email,'recipient@example.test');
  rows[1][3]='変更後の会社';rows[1][17]='変更後の業種';
  const page=context.doGet({parameter:{grant:'women-r8',invite:result.token,preview:'1'}}).template;
  assert.equal(page.previewMode,false);assert.equal(page.prefill.company,'会社A');assert.equal(page.prefill.industry,'通信業');
  assert.equal(page.prefill.employees[0].name,'従業員A');
});
test('large prefill is stored outside script properties and old invitation format still works',()=>{
  const {context,rows,saved}=harness();
  rows[1][16]=JSON.stringify(Array.from({length:80},(_,i)=>({name:'従業員'+i,position:'担当者',employment:'regular'})));
  const result=context.createGrantInvitation({company:'会社A',email:'recipient@example.test',grant:'women-r8',sourceReceipt:rows[1][1]});
  const invitation=context.readGrantInvitation_(result.token);
  assert(invitation.prefillFileId);assert.equal(saved.length,1);
  assert.equal(context.readApplicantInvitationPrefill_(invitation).employees.length,80);
  assert.equal(context.readApplicantInvitationPrefill_({prefill:{address:'以前の住所'}}).address,'以前の住所');
});

test('legacy applicant drafts save partial fields and employee rows, reject invalid or submitted invitations',()=>{
 const {context}=harness();let invite={grant:'women-r8',company:'Company',email:'test@example.test'},stored='',n=0;
 context.readGrantInvitation_=token=>{if(token!=='valid-test-token')throw Error('invalid');return invite;};
 context.writeGrantInvitation_=(_token,value)=>{invite=value;};
 context.requirementStore_=()=>({createFile(_name,json){stored=json;n++;return {getId:()=> 'test-draft'};}});
 context.DriveApp.getFileById=()=>({setContent(json){stored=json;},getBlob(){return {getDataAsString:()=>stored};}});
 const input={token:'valid-test-token',values:{person:'途中',employees:[{name:'従業員',position:'',employment:''}]}};
 assert.equal(context.saveWomenR8AnswerDraft(input),true);assert.equal(n,1);assert.equal(JSON.parse(stored).values.person,'途中');assert.equal(JSON.parse(stored).values.employees[0].position,'');
 input.values.person='再編集';context.saveWomenR8AnswerDraft(input);assert.equal(n,1);assert.equal(JSON.parse(stored).values.person,'再編集');
 assert.throws(()=>context.saveWomenR8AnswerDraft({...input,token:'invalid'}));
 invite.receipt='submitted';assert.throws(()=>context.saveWomenR8AnswerDraft(input));
});

test('company ledger spreadsheet rows preserve identifiers and escape spreadsheet formulas',()=>{
 const {context}=harness();const rows=context.companyLedgerRows_([{id:'id1',name:'=formula',profile:{phone:'00123',industry:'通信'},entries:[['管理番号','0001']]}]);
 assert.equal(rows[0][1],'会社名');assert.equal(rows[1][1],"'=formula");assert.equal(rows[1][7],'00123');assert.equal(rows[1][8],'0001');
});
test('company ledger folder configuration exports existing records and returns folder and spreadsheet links',()=>{
 const {context,props,setUser}=harness();let moved=0,sync=0;
 const folder={getUrl:()=> 'https://drive.google.com/drive/folders/test-folder-123',getName:()=> '台帳保存先',isTrashed:()=>false,createFile:()=>({setTrashed(){}})};
 context.DriveApp.getFolderById=()=>folder;context.DriveApp.getFileById=()=>({moveTo(f){assert.equal(f,folder);moved++;},setTrashed(){}});
 context.SpreadsheetApp.create=()=>({getId:()=> 'ledger-sheet'});context.SpreadsheetApp.openById=()=>({getUrl:()=> 'https://docs.google.com/spreadsheets/d/ledger-sheet/edit'});
 context.listCompanyNumbers=()=>[{id:'company1',name:'会社'}];context.syncCompanyLedgerSheet_=(records,id)=>{assert.equal(records[0].name,'会社');assert.equal(id,'ledger-sheet');sync++;};
 const result=context.saveCompanyLedgerStorageSettings('https://drive.google.com/drive/folders/test-folder-123');assert.equal(props.COMPANY_LEDGER_FOLDER_ID,'test-folder-123');assert.equal(props.COMPANY_LEDGER_SPREADSHEET_ID,'ledger-sheet');assert.equal(sync,1);assert.equal(moved,1);assert(result.spreadsheetUrl.includes('ledger-sheet'));
 setUser('other@example.test');assert.throws(()=>context.getCompanyLedgerStorageSettings());assert.throws(()=>context.saveCompanyLedgerStorageSettings('bad'));
});
