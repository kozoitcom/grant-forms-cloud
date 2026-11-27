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

test('legacy applicant drafts save partial fields and employee rows, reject invalid invitations and allow corrections after submission',()=>{
 const {context}=harness();let invite={grant:'women-r8',company:'Company',email:'test@example.test'},stored='',n=0;
 context.readGrantInvitation_=token=>{if(token!=='valid-test-token')throw Error('invalid');return invite;};
 context.writeGrantInvitation_=(_token,value)=>{invite=value;};
 context.requirementStore_=()=>({createFile(_name,json){stored=json;n++;return {getId:()=> 'test-draft'};}});
 context.DriveApp.getFileById=()=>({setContent(json){stored=json;},getBlob(){return {getDataAsString:()=>stored};}});
 const input={token:'valid-test-token',values:{person:'途中',employees:[{name:'従業員',position:'',employment:''}]}};
 assert.equal(context.saveWomenR8AnswerDraft(input),true);assert.equal(n,1);assert.equal(JSON.parse(stored).values.person,'途中');assert.equal(JSON.parse(stored).values.employees[0].position,'');
 input.values.person='再編集';context.saveWomenR8AnswerDraft(input);assert.equal(n,1);assert.equal(JSON.parse(stored).values.person,'再編集');
 assert.throws(()=>context.saveWomenR8AnswerDraft({...input,token:'invalid'}));
 invite.receipt='submitted';assert.equal(context.saveWomenR8AnswerDraft(input),true);
});

test('company ledger spreadsheet rows preserve identifiers and escape spreadsheet formulas',()=>{
 const {context}=harness();const rows=context.companyLedgerRows_([{id:'id1',name:'=formula',profile:{phone:'00123',industry:'通信'},entries:[['管理番号','0001']]}]);
 assert.equal(rows[0][1],'会社名');assert.equal(rows[1][1],"'=formula");assert.equal(rows[1][7],'00123');assert.equal(rows[1][8],'0001');
});
test('company ledger folder setting exports records separately and keeps old shared spreadsheet',()=>{
 const {context,props,setUser}=harness();let sync=0;props.COMPANY_LEDGER_SPREADSHEET_ID='old-shared';
 const folder={getUrl:()=> 'https://drive.google.com/drive/folders/test-folder-123',getName:()=> '台帳保存先',isTrashed:()=>false,createFile:()=>({setTrashed(){}})};
 context.DriveApp.getFolderById=()=>folder;context.companyLedgerRecords_=()=>[{id:'company1',name:'会社'}];context.syncCompanyLedgerSheet_=(records,id)=>{assert.equal(records[0].name,'会社');assert.equal(id,'test-folder-123');sync++;};
 const result=context.saveCompanyLedgerStorageSettings('https://drive.google.com/drive/folders/test-folder-123');assert.equal(props.COMPANY_LEDGER_FOLDER_ID,'test-folder-123');assert.equal(props.COMPANY_LEDGER_SPREADSHEET_ID,'old-shared');assert.equal(sync,1);assert.equal(result.spreadsheetUrl,'');assert.equal(result.perCompany,true);
 setUser('other@example.test');assert.throws(()=>context.getCompanyLedgerStorageSettings());assert.throws(()=>context.saveCompanyLedgerStorageSettings('bad'));assert.throws(()=>context.createCompanyLedgerSpreadsheets());
});

test('submitted company data accumulates across grants and prefills matching fields without erasing blanks',()=>{
 const {context,props}=harness();delete props.COMPANY_NUMBERS_KEEP;const files=new Map();let fileCounter=0,syncs=0;
 const folder={createFile(name,text){const id='details-'+(++fileCounter);const file={text,getId:()=>id,getBlob(){return {getDataAsString:()=>this.text};},setContent(text){this.text=text;}};files.set(id,file);return file;}};
 context.requirementStore_=()=>folder;context.DriveApp.getFileById=id=>files.get(id);context.syncCompanyLedgerSheet_=()=>{syncs++;};
 const invite={grant:'schema:test',company:'会社A',email:'a@example.test'};
 const birth={id:'birth',label:'生年月日',type:'date',enabled:true},first={receipt:'r1',company:'会社A',grant:'schema:a',submittedAt:'2026-11-26T00:00:00Z',values:{company:'会社A',address:'住所A',email:'a@example.test',industry:'通信業',birth:'1990-01-02'},definition:{fields:[birth]},employees:[]};
 const id=context.upsertSubmittedCompany_(invite,first);assert.equal(invite.companyId,id);assert.equal(context.companyLedgerRecords_().length,1);
 context.upsertSubmittedCompany_(invite,{receipt:'r2',company:'会社A改',grant:'schema:b',submittedAt:'2026-11-27T00:00:00Z',values:{address:'',person:'担当者B',email:'a@example.test'},definition:{fields:[]},employees:[]});
 let record=context.companyLedgerRecords_()[0];assert.equal(record.name,'会社A');assert.equal(record.profile.address,'住所A');assert.equal(record.profile.person,'担当者B');assert.equal(context.companyLedgerDetails_(record).fields[0].value,'1990-01-02');
 context.requirementRead_=()=>({published:{fields:[{id:'dateOfBirth',label:'生年月日',type:'date',enabled:true}]}});
 const prefill=context.applicantPrefill_(id,'','schema:next');assert.equal(prefill.dateOfBirth,'1990-01-02');assert.equal(prefill.address,'住所A');assert.equal(prefill.company,'会社A');
 context.upsertSubmittedCompany_(invite,{...first,receipt:'r-new',company:'変更名',values:{company:'変更名',address:'変更住所',email:'changed@example.test',birth:'2001-03-04',newInfo:'追加情報'},definition:{fields:[birth,{id:'newInfo',label:'追加項目',type:'text',enabled:true}]}});
 record=context.companyLedgerRecords_()[0];const detail=context.companyLedgerDetails_(record);assert.equal(record.name,'会社A');assert.equal(record.profile.address,'住所A');assert.equal(record.profile.email,'a@example.test');assert.equal(detail.values.birth,'1990-01-02');assert.equal(detail.fields.find(f=>f.label==='生年月日').value,'1990-01-02');assert.equal(detail.fields.find(f=>f.label==='追加項目').value,'追加情報');assert.equal(detail.values.newInfo,'追加情報');assert.equal(files.size,1);
 const prior=files.size;context.upsertSubmittedCompany_(invite,{...first,receipt:'r2'});assert.equal(files.size,prior);assert.equal(context.companyLedgerRecords_().length,1);
 context.syncCompanyLedgerSheet_=()=>{throw Error('temporary sheet failure');};context.accumulateSubmittedCompany_(invite,{...first,receipt:'r3'});assert.equal(invite.ledgerSyncPending,true);context.syncCompanyLedgerSheet_=()=>{};context.accumulateSubmittedCompany_(invite,{...first,receipt:'r3'});assert.equal(invite.ledgerSyncPending,false);
});
test('company-name alone never overwrites an unrelated ledger record',()=>{
 const {context,props}=harness();delete props.COMPANY_NUMBERS_KEEP;const files=new Map();let n=0;context.requirementStore_=()=>({createFile(name,text){const id='detail-'+(++n),file={getId:()=>id,getBlob:()=>({getDataAsString:()=>text}),setContent(value){text=value;}};files.set(id,file);return file;}});context.DriveApp.getFileById=id=>files.get(id);context.syncCompanyLedgerSheet_=()=>{};
 for(const email of ['first@example.test','second@example.test'])context.upsertSubmittedCompany_({company:'同名会社',email},{company:'同名会社',receipt:email,values:{email},grant:'schema:test',submittedAt:'now'});
 assert.equal(context.companyLedgerRecords_().length,2);
});


test('legacy resubmission reuses receipt, folder, JSON and response row',()=>{
 const {context}=harness();vm.runInContext(read('RequirementForms.js'),context);let invite={grant:'women-r8',email:'test@example.test'},folderCount=0,next=0;const saved=[],rows=[['header']];
 const iter=a=>{let i=0;return {hasNext:()=>i<a.length,next:()=>a[i++]}};
 const folder={getId:()=> 'submission',getUrl:()=> 'https://drive.google.com/drive/folders/submission',isTrashed:()=>false,createFile(name,text){if(typeof name==='object'){text=name.getBytes();name=name.getName()}const f={id:'file'+(++next),name,text,trashed:false,getId(){return this.id},getSize(){return this.text.length},isTrashed(){return this.trashed},setTrashed(v){this.trashed=v},setName(v){this.name=v},setContent(v){this.text=v},getBlob(){return {getDataAsString:()=>this.text}}};saved.push(f);return f;},getFilesByName:n=>iter(saved.filter(f=>!f.trashed&&f.name===n))};
 const root={isTrashed:()=>false,createFolder(){folderCount++;return folder}};context.DriveApp.getFolderById=()=>folder;context.DriveApp.getFileById=id=>saved.find(f=>f.id===id);
 const sheet={getDataRange:()=>({getValues:()=>rows}),appendRow:r=>rows.push(r),getRange:(row,col)=>({getValue:()=>true,setValue(){},setValues(values){rows[row-1]=values[0]}})};
 context.getWomenR8Storage_=()=>({root,sheet});context.SpreadsheetApp.flush=()=>{};context.readGrantInvitation_=()=>({...invite});context.writeGrantInvitation_=(_,v)=>invite=v;context.validateWomenR8Employees_=()=>[];context.validateWomenR8Counts_=()=>({});context.accumulateSubmittedCompany_=()=>{};context.getDocumentSettings_=()=>({});vm.runInContext("WOMEN_R8_DOCUMENTS.splice(0,WOMEN_R8_DOCUMENTS.length,['doc','書類','all'])",context);
 const blob=()=>({name:'source.pdf',getBytes:()=>[37,80,68,70,45],getName(){return this.name},setName(v){this.name=v;return this}});
 const input={inviteToken:'token',industry:'通信',entity:'company',location:'no',company:'初回',address:'住所',representative:'代表',person:'担当',email:'test@example.test',submissionVersion:0,submissionRequestId:'first',doc:blob()};
 const first=context.submitWomenR8(input);const second=context.submitWomenR8({...input,submissionVersion:1,submissionRequestId:'second',company:'更新',doc:blob()});assert.equal(first.receipt,second.receipt);assert.equal(folderCount,1);assert.equal(rows.length,2);assert.equal(rows[1][3],'更新');assert.equal(saved.filter(f=>!f.trashed&&f.name==='書類.pdf').length,1);assert.equal(saved.filter(f=>!f.trashed&&f.name==='入力内容.json').length,1);assert.equal(second.submissionVersion,2);
 context.submitWomenR8({...input,doc:null,submissionVersion:2,submissionRequestId:'third'});assert.equal(rows.length,2);assert.equal(folderCount,1);
});

test('multiple company field aliases persist, prefill alternate labels and reject ambiguity',()=>{
 const {context,props}=harness();delete props.COMPANY_NUMBERS_KEEP;
 const id='11111111-1111-1111-1111-111111111111';props['COMPANY_NUMBERS_'+id]=JSON.stringify({id,name:'会社A',profile:{address:'住所A'},entries:[]});context.syncCompanyLedgerSheet_=()=>{};
 context.saveCompanyFieldAliases({companyId:id,fields:[{key:'common:company',aliases:'企業名\n法人名\n企業名'},{key:'common:address',aliases:'所在地\n住所'}]});
 assert.equal(context.getCompanyFieldAliases(id).find(f=>f.key==='common:company').aliases.length,2);
 context.requirementRead_=()=>({published:{fields:[{id:'corp',label:'企業名',type:'text',enabled:true},{id:'legal',label:'法人名',type:'text',enabled:true},{id:'place',label:'住所',type:'text',enabled:true},{id:'wrongType',label:'企業名',type:'number',enabled:true}]}});
 const prefill=context.applicantPrefill_(id,'','schema:test');assert.equal(prefill.corp,'会社A');assert.equal(prefill.legal,'会社A');assert.equal(prefill.place,'住所A');assert.equal(prefill.wrongType,undefined);
 assert.throws(()=>context.saveCompanyFieldAliases({companyId:id,fields:[{key:'common:address',aliases:'企業名'}]}),/複数/);
 context.saveCompanyFieldAliases({companyId:id,fields:[{key:'common:company',aliases:''}]});assert.equal(context.getCompanyFieldAliases(id).find(f=>f.key==='common:company').aliases.length,0);assert.equal(context.companyLedgerRecords_()[0].name,'会社A');
});

test('transfer workspace saves mixed templates with frozen company base, validates company and retries upload once',()=>{
 const {context,props}=harness();let next=0;const iter=a=>{let i=0;return {hasNext:()=>i<a.length,next:()=>a[i++]}};const folders=new Map();
 function folder(id){const fs=[],children=[];const f={id,getId:()=>id,createFolder(name){const child=folder(name);children.push(child);return child;},getFoldersByName:n=>iter(children.filter(x=>x.id===n)),getFolders:()=>iter(children),getFilesByName:n=>iter(fs.filter(x=>x.name===n)),createFile(name,text){if(typeof name==='object'){text=name.bytes;name=name.name;}const file={name,text,id:'f'+(++next),getId(){return this.id},getUrl(){return 'https://drive.google.com/file/d/'+this.id+'/view'},getBlob(){return {getDataAsString:()=>this.text}},setContent(text){this.text=text;}};fs.push(file);return file;}};folders.set(id,f);return f;}
 context.DriveApp={createFolder:()=>folder('root'),getFolderById:id=>folders.get(id)};context.Utilities.base64Decode=s=>[...Buffer.from(s,'base64')];context.Utilities.newBlob=(bytes,mime,name)=>({bytes,name,getName(){return this.name},getBytes(){return this.bytes},setName(v){this.name=v;return this}});context.requirementUploadExtension_=(_field,_blob,bytes)=>{if(bytes[0]!==80)throw Error('format');};
 const record={id:'company',name:'会社A',profile:{address:'住所A'}},details={values:{custom:'保存値'},fields:[]};context.companyLedgerRecords_=()=>[record];context.companyLedgerDetails_=()=>details;context.collectionExport_=()=>({company:'別会社',receipt:'wrong'});
 const input={companyId:'company',grantName:'助成金',year:'令和8年度',output:'original_pdf',notes:'今回の計画',files:[{name:'申請書.docx',size:2,role:'blank'},{name:'別紙.xlsx',size:2,role:'blank'}]};
 assert.throws(()=>context.beginTransferWorkspace({...input,receipt:'wrong'}),/一致/);assert.throws(()=>context.beginTransferWorkspace({...input,files:[{name:'申請書.pdf',size:2,role:'blank'}]}),/Word/);
 const {id}=context.beginTransferWorkspace(input);assert.throws(()=>context.completeTransferWorkspace(id),/すべて/);assert.throws(()=>context.uploadTransferWorkspaceFile({id,index:0,base64:'AA=='}),/サイズ/);
 context.uploadTransferWorkspaceFile({id,index:0,base64:'UEs='});const prior=next;context.uploadTransferWorkspaceFile({id,index:0,base64:'UEs='});assert.equal(next,prior);context.uploadTransferWorkspaceFile({id,index:1,base64:'UEs='});
 const result=context.completeTransferWorkspace(id);assert.equal(result.base.values.address,'住所A');assert.equal(result.files.length,2);assert.equal(result.generationAvailable,false);record.profile.address='変更後';details.values.custom='変更値';const saved=context.getTransferWorkspace(id);assert.equal(saved.base.values.address,'住所A');assert.equal(saved.base.values.custom,'保存値');assert.equal(context.listTransferWorkspaces().length,1);assert.equal(saved.output,'original_pdf');
});
test('transfer workspace endpoints require administrator before storage access',()=>{
 const {context,setUser}=harness();setUser('outsider@example.test');for(const [name,arg]of [['beginTransferWorkspace',{}],['uploadTransferWorkspaceFile',{}],['completeTransferWorkspace','bad'],['listTransferWorkspaces'],['getTransferWorkspace','bad']])assert.throws(()=>context[name](arg));
});
test('transfer tab provides mixed uploads and company base without advertising completed document generation',()=>{
 const html=read('DocumentTransferView.html');assert.match(html,/ws-company/);assert.match(html,/multiple accept="\.docx,\.doc,\.xlsx,\.xls,\.pdf"/);assert.match(html,/ondrop/);assert.match(html,/ws-notes/);assert.match(html,/ws-output/);assert.match(html,/自動記入とPDF出力は未実装/);for(const m of html.matchAll(/<script>([\s\S]*?)<\/script>/g))assert.doesNotThrow(()=>new vm.Script(m[1]));
});

test('each company receives one stable spreadsheet and never writes another company or legacy shared file',()=>{
 const {context,props}=harness();props.COMPANY_LEDGER_FOLDER_ID='folder';props.COMPANY_LEDGER_SPREADSHEET_ID='legacy';const books=new Map(),writes=[];let creates=0,moves=0,fail=false;
 const folder={isTrashed:()=>false};context.DriveApp.getFolderById=()=>folder;context.DriveApp.getFileById=id=>({isTrashed:()=>false,moveTo(f){assert.equal(f,folder);moves++;}});
 context.SpreadsheetApp={create(name){const id='sheet-'+(++creates),book={getId:()=>id,name};books.set(id,book);return book;},openById:id=>{assert.notEqual(id,'legacy');return books.get(id)},flush(){}};
 context.writeCompanyLedgerRows_=(book,record)=>{if(fail)throw Error('temporary');writes.push([book.getId(),record.id]);};
 const companies=[{id:'a',name:'会社A'},{id:'b',name:'会社B'}];context.syncCompanyLedgerSheet_(companies);assert.equal(creates,2);assert.notEqual(props.COMPANY_LEDGER_SHEET_a,props.COMPANY_LEDGER_SHEET_b);
 context.syncCompanyLedgerSheet_([{...companies[0],name:'変更名'}]);assert.equal(creates,2);assert.equal(writes[2][0],props.COMPANY_LEDGER_SHEET_a);assert.equal(writes[2][1],'a');assert.equal(props.COMPANY_LEDGER_SPREADSHEET_ID,'legacy');
 fail=true;assert.throws(()=>context.syncCompanyLedgerSheet_([{id:'c',name:'会社C'}]),/temporary/);const stable=props.COMPANY_LEDGER_SHEET_c;fail=false;context.syncCompanyLedgerSheet_([{id:'c',name:'会社C'}]);assert.equal(creates,3);assert.equal(props.COMPANY_LEDGER_SHEET_c,stable);
});
