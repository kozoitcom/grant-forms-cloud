const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),crypto=require('node:crypto');
const app=path.join(__dirname,'../apps-script');
const source=['コード.js','GrantSetup.js','MaterialBatch.js'].map(name=>fs.readFileSync(path.join(app,name),'utf8')).join('\n');
const batchId='11111111-1111-1111-1111-111111111111';
function harness(){
  const props={},files=new Map(),folders=new Map();let next=0,currentUser='admin@example.test';
  const iterator=items=>{let index=0;return {hasNext:()=>index<items.length,next:()=>items[index++]};};
  const blob=(content,mime,name)=>{const bytes=Buffer.isBuffer(content)?content:typeof content==='string'?Buffer.from(content):Buffer.from(content);return {getBytes:()=>[...bytes],getName:()=>name,getDataAsString:()=>bytes.toString('utf8'),setName(value){name=value;return this;}};};
  function createFile(folder,name,content){
    let data=typeof name==='object'?name:blob(content,'text/plain',name);
    const file={id:'file-'+(++next),name:data.getName(),data,trashed:false,getId(){return this.id},getUrl(){return 'https://drive.google.com/file/d/'+this.id+'/view'},getBlob(){return this.data},setContent(text){this.data=blob(text,'text/plain',this.name);return this},isTrashed(){return this.trashed},setTrashed(value){this.trashed=value}};
    files.set(file.id,file);folder.files.push(file);return file;
  }
  function createFolder(name){const folder={id:'folder-'+(++next),name,files:[],folders:[],trashed:false,getId(){return this.id},getUrl(){return 'https://drive.google.com/drive/folders/'+this.id},createFolder(name){const item=createFolder(name);this.folders.push(item);return item},createFile(name,content){return createFile(this,name,content)},getFolders(){return iterator(this.folders.filter(f=>!f.trashed))},getFoldersByName(name){return iterator(this.folders.filter(f=>f.name===name&&!f.trashed))},getFilesByName(name){return iterator(this.files.filter(f=>f.name===name&&!f.trashed))},setTrashed(value){this.trashed=value}};folders.set(folder.id,folder);return folder;}
  const ctx={console,Date,
    Session:{getActiveUser:()=>({getEmail:()=>currentUser}),getEffectiveUser:()=>({getEmail:()=> 'admin@example.test'})},
    PropertiesService:{getScriptProperties:()=>({getProperty:key=>props[key],setProperty:(key,value)=>{props[key]=value}})},
    DriveApp:{createFolder,getFolderById:id=>folders.get(id),getFileById:id=>{if(!files.has(id))throw Error('Missing file');return files.get(id)}},
    MimeType:{PLAIN_TEXT:'text/plain'},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
    Utilities:{newBlob:blob,DigestAlgorithm:{SHA_256:'sha256'},computeDigest:(algorithm,data)=>crypto.createHash(algorithm).update(typeof data==='string'?data:Buffer.from(data)).digest(),base64EncodeWebSafe:data=>Buffer.from(data).toString('base64url'),base64Decode:text=>[...Buffer.from(text,'base64')]}
  };
  vm.createContext(ctx);vm.runInContext(source,ctx);
  return {ctx,props,files,folders,setUser:value=>{currentUser=value}};
}
const pdf=Buffer.from('%PDF-1.4\ntest file'),docx=Buffer.from([80,75,3,4,0,0,0,0]);
function input(files=[{name:'募集要項.pdf',size:pdf.length,role:'guidelines'}]) {return {requestId:batchId,title:'テスト助成金',year:'令和8年度',grant:'',files};}
function upload(ctx,index,bytes=pdf){return ctx.uploadGrantMaterialFile({id:batchId,index,base64:bytes.toString('base64')});}
test('metadata validates roles, limits, duplicate targets and format of guidelines',()=>{
 const {ctx}=harness();
 for(const data of [input([]),input([{name:'bad.exe',size:10,role:'blank',documentTitle:'書類'}]),input([{name:'big.pdf',size:21*1024*1024,role:'guidelines'}]),input([{name:'unknown.pdf',size:10,role:''}]),input([{name:'要項.docx',size:8,role:'guidelines'}])])assert.throws(()=>ctx.beginGrantMaterialBatch(data));
 const duplicate=input([{name:'a.docx',size:8,role:'blank',documentKey:'rules'},{name:'b.docx',size:8,role:'blank',documentKey:'rules'}]);duplicate.grant='women-r8';assert.throws(()=>ctx.beginGrantMaterialBatch(duplicate));
 assert.throws(()=>ctx.beginGrantMaterialBatch(input(Array.from({length:3},(_,i)=>({name:i+'.pdf',size:20*1024*1024,role:'guidelines'})))));
});
test('retrying begin and file upload does not create duplicate folders or files',()=>{
 const {ctx,files,folders}=harness();
 ctx.beginGrantMaterialBatch(input());const folderCount=folders.size;ctx.beginGrantMaterialBatch(input());assert.equal(folders.size,folderCount);
 upload(ctx,0);const fileCount=files.size;upload(ctx,0);assert.equal(files.size,fileCount);
 const first=ctx.completeGrantMaterialBatch(batchId);const again=ctx.completeGrantMaterialBatch(batchId);assert.equal(first.url,again.url);
 assert.equal(ctx.listGrantMaterialBatches()[0].state,'completed');
 assert.throws(()=>ctx.beginGrantMaterialBatch({...input(),title:'changed'}));
});
test('incomplete uploads cannot publish template settings; retry completes the rest',()=>{
 const {ctx,props}=harness();props.WOMEN_R8_DOCUMENT_SETTINGS=JSON.stringify({rules:{uploadNotRequired:true}});
 const data=input([{name:'空欄.docx',size:docx.length,role:'blank',documentKey:'rules'},{name:'見本.pdf',size:pdf.length,role:'sample',documentKey:'rules'}]);data.grant='women-r8';
 ctx.beginGrantMaterialBatch(data);upload(ctx,0,docx);assert.throws(()=>ctx.completeGrantMaterialBatch(batchId));assert(!ctx.getDocumentSettings_().rules.blank);
 upload(ctx,1);ctx.completeGrantMaterialBatch(batchId);const settings=ctx.getDocumentSettings_();assert(settings.rules.blank.id);assert(settings.rules.sample.id);assert.equal(settings.rules.uploadNotRequired,true);
});
test('only one of blank or sample is sufficient and unrelated settings survive',()=>{
 const {ctx,props}=harness();props.WOMEN_R8_DOCUMENT_SETTINGS=JSON.stringify({chart:{sample:{id:'old',name:'old.pdf',url:'url'}}});
 const data=input([{name:'見本.pdf',size:pdf.length,role:'sample',documentKey:'rules'}]);data.grant='women-r8';ctx.beginGrantMaterialBatch(data);upload(ctx,0);ctx.completeGrantMaterialBatch(batchId);
 const settings=ctx.getDocumentSettings_();assert(settings.rules.sample);assert(!settings.rules.blank);assert.equal(settings.chart.sample.id,'old');
});
test('file signature and declared size are checked and changed bytes are rejected on retry',()=>{
 const {ctx}=harness();ctx.beginGrantMaterialBatch(input());
 assert.throws(()=>upload(ctx,0,Buffer.alloc(pdf.length)));assert.throws(()=>upload(ctx,0,Buffer.from('%PDF-1.4')));
 upload(ctx,0);const changed=Buffer.from(pdf);changed[changed.length-1]=0;assert.throws(()=>upload(ctx,0,changed));
});
test('new grant and other documents are archived without changing the applicant form configuration',()=>{
 const {ctx,props}=harness();const data=input([{name:'会社見本.docx',size:docx.length,role:'sample',documentTitle:'独自書類'}]);
 ctx.beginGrantMaterialBatch(data);upload(ctx,0,docx);ctx.completeGrantMaterialBatch(batchId);assert.equal(props.WOMEN_R8_DOCUMENT_SETTINGS,undefined);
 assert.equal(ctx.listGrantMaterialBatches()[0].files[0].documentTitle,'独自書類');
});
test('large document settings use Drive and remain editable through the same getter',()=>{
 const {ctx,props}=harness();const settings={rules:{sample:{name:'a'.repeat(8500)}}};ctx.writeDocumentSettings_(settings);assert(props.WOMEN_R8_DOCUMENT_SETTINGS_FILE_ID);assert.equal(ctx.getDocumentSettings_().rules.sample.name.length,8500);
 settings.rules.uploadNotRequired=true;ctx.writeDocumentSettings_(settings);assert.equal(ctx.getDocumentSettings_().rules.uploadNotRequired,true);
});
test('every public bulk endpoint rejects non-admin users',()=>{
 const {ctx,setUser}=harness();setUser('other@example.test');
 for(const [name,arg] of [['getGrantMaterialCatalog'],['beginGrantMaterialBatch',input()],['uploadGrantMaterialFile',{}],['completeGrantMaterialBatch',batchId],['listGrantMaterialBatches']])assert.throws(()=>ctx[name](arg));
});
test('empty listing does not initialize Drive folders',()=>{const {ctx,folders}=harness();assert.equal(ctx.listGrantMaterialBatches().length,0);assert.equal(folders.size,0);});

test('drop handler accepts mixed files, adds later drops, and rejects oversized files and duplicates',async()=>{
  const html=fs.readFileSync(path.join(app,'GrantSetupView.html'),'utf8');
  const script=html.match(/<script>([\s\S]*?)<\/script>/)[1];
  const nodes=new Map();
  const element=tag=>({tagName:tag,value:'',style:{},children:[],handlers:{},disabled:false,
    append(...items){this.children.push(...items)},add(item){this.children.push(item)},replaceChildren(...items){this.children=items},
    insertRow(){const item=element('tr');this.children.push(item);return item},insertCell(){const item=element('td');this.children.push(item);return item},
    setAttribute(){},addEventListener(name,handler){const old=this.handlers[name];this.handlers[name]=old?event=>{old(event);handler(event)}:handler},querySelectorAll(){return []}});
  const document={getElementById:id=>{if(!nodes.has(id))nodes.set(id,element(id));return nodes.get(id)},createElement:element};
  let success;
  const rpc=new Proxy({withSuccessHandler(fn){success=fn;return rpc},withFailureHandler(){return rpc}},{get(target,name){if(name in target)return target[name];return (...args)=>{success(name==='getGrantMaterialCatalog'?{documents:[],grants:[]}:[]);};}});
  const context={document,window:{addEventListener(){}},google:{script:{run:rpc}},Option:function(text,value){this.text=text;this.value=value},console};
  vm.createContext(context);vm.runInContext(script,context);await Promise.resolve();
  let prevented=0;
  const file=(name,size,lastModified=1)=>({name,size,lastModified});
  const drop=files=>nodes.get('material-dropzone').handlers.drop({preventDefault(){prevented++},dataTransfer:{files}});
  drop([file('募集要項.pdf',100),file('空欄様式.docx',200),file('記入済み見本.xlsx',300)]);
  assert.equal(vm.runInContext('materialQueue.length',context),3);
  assert.equal(vm.runInContext('materialQueue.map(item=>item.role).join(",")',context),'guidelines,blank,sample');
  drop([file('追加見本.pdf',400),file('募集要項.pdf',100),file('大きな.pdf',21*1024*1024)]);
  assert.equal(vm.runInContext('materialQueue.length',context),4);assert.equal(prevented,2);
  assert(nodes.get('material-upload-status').textContent.includes('選択済み'));
  assert(nodes.get('material-upload-status').textContent.includes('サイズ'));
});
