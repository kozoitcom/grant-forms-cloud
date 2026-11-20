const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const app = path.join(__dirname,'../apps-script');
const read = file => fs.readFileSync(path.join(app,file),'utf8');
const files = ['コード.js','GrantSetup.js','ApplicantData.js','DocumentTransfer.js'];
const code = files.map(read).join('\n');
function harness() {
  let n=0;
  const props = {WOMEN_R8_SHEET_ID:'sheet',COMPANY_NUMBERS_KEEP:'preserved',WOMEN_R8_DOCUMENT_SETTINGS:'{"rules":{"sample":{"id":"template-1","name":"sample.docx","url":"https://drive.google.com/file/d/template-1/view"}}}'};
  const rows = [['header'],[new Date('2026-11-20T00:00:00Z'),'12345678-1234-1234-1234-123456789abc','法人','会社A','住所A','代表A','担当A','test@example.test','00-0000-0000','備考','no','',1,'https://drive.google.com/drive/folders/submission','かな','{"groups":[],"eligible":false}','[{"name":"従業員A","position":"なし","employment":"regular"}]','通信業']];
  const saved=[];
  const iterator = items => {let i=0;return {hasNext:()=>i<items.length,next:()=>items[i++]};};
  const folder={createFile:(name,body)=>{const file={name,body,description:'',getUrl:()=> 'https://drive.google.com/file/d/snapshot/view',setDescription(value){this.description=value;},getDescription(){return this.description;},getBlob(){return {getDataAsString:()=>this.body}},setTrashed(){}};saved.push(file);return file;},getId:()=> 'transfer-folder',getFiles:()=>iterator(saved),getFilesByName:name=>iterator(saved.filter(f=>f.name===name))};
  let currentEmail='admin@example.test';
  const context={Date,console,
    Session:{getActiveUser:()=>({getEmail:()=>currentEmail}),getEffectiveUser:()=>({getEmail:()=> 'admin@example.test'})},
    PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k],getProperties:()=>({...props}),setProperty:(k,v)=>props[k]=v})},
    SpreadsheetApp:{openById:()=>({getSheets:()=>[{getDataRange:()=>({getValues:()=>rows})}]})},
    DriveApp:{createFolder:()=>folder,getFolderById:()=>folder},
    LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
    Utilities:{getUuid:()=> 'aaaaaaaa-aaaa-aaaa-aaaa-'+String(++n).padStart(12,'0')},MimeType:{PLAIN_TEXT:'text/plain'},
    ScriptApp:{getService:()=>({getUrl:()=> 'https://script.google.com/macros/s/deployment/exec'})},
    HtmlService:{createTemplateFromFile:()=>({evaluate(){return {setTitle(){return this},addMetaTag(){return this},template:this}}})}
  };
  vm.createContext(context);vm.runInContext(code,context);
  return {context,props,rows,saved,setUser:email=>currentEmail=email};
}
test('each system loads its own view; unknown route safely selects setup',()=>{
  const {context}=harness();
  for(const [system,view] of [['setup','GrantSetupView'],['applications','ApplicantDataView'],['transfer','DocumentTransferView'],['<script>','GrantSetupView']]) {
    const page=context.doGet({parameter:{page:'admin',system}}).template;
    assert.equal(page.systemFile,view);
    assert.equal(page.systems.length,3);
  }
});
test('views and scripts are separate and do not include other systems form controls',()=>{
  for(const file of ['GrantSetupView.html','ApplicantDataView.html','DocumentTransferView.html']) {
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
