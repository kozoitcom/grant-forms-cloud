const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),crypto=require('node:crypto');
test('editing the same field after saving updates the newly returned draft object',async()=>{
 const nodes=new Map();
 const el=tag=>({tagName:tag.toUpperCase(),value:'',checked:false,style:{},children:[],setAttribute(k,v){this[k]=v},append(...x){this.children.push(...x)},replaceChildren(...x){this.children=x},add(x){this.children.push(x)}});
 const document={getElementById:id=>{if(!nodes.has(id))nodes.set(id,el('input'));return nodes.get(id)},createElement:el,querySelectorAll:()=>[]};
 let saved=null;const calls=[];
 const handlers={listRequirementForms:()=>saved?[saved]:[],listGrantMaterialBatches:()=>[],listGrantInvitations:()=>[],getRequirementAiStatus:()=>({configured:false}),saveRequirementForm:input=>{calls.push(structuredClone(input));saved={...structuredClone(input),id:'11111111-1111-1111-1111-111111111111',revision:input.revision+1};return structuredClone(saved)}};
 const run=()=>{let ok,fail;return new Proxy({withSuccessHandler(fn){ok=fn;return this},withFailureHandler(fn){fail=fn;return this}},{get(target,name){if(name in target)return target[name];return (...args)=>Promise.resolve().then(()=>handlers[name](...args)).then(ok,fail)}})};
 const scriptRun={};Object.defineProperty(scriptRun,'run',{get:run});
 const ctx={document,google:{script:scriptRun},Option:function(text,value){this.text=text;this.value=value},crypto,structuredClone,Date,console,confirm:()=>true};vm.createContext(ctx);
 const html=fs.readFileSync(path.join(__dirname,'../apps-script/RequirementEditor.html'),'utf8');vm.runInContext(html.match(/<script>([\s\S]*?)<\/script>/)[1],ctx);await new Promise(setImmediate);
 nodes.get('req-title').value='動作確認';nodes.get('req-year').value='令和8年度';nodes.get('req-add').onclick();await nodes.get('req-save').onclick();
 const row=nodes.get('req-fields').children[0];const details=row.children[1];assert.equal(details.tagName,'DETAILS');assert(!details.open);const labels=details.children[1].children;const control=title=>labels.find(x=>x.children?.[0]===title).children[1];
 const enabled=row.children[0].children[0];enabled.checked=false;enabled.onchange();assert.equal(row.children[0].children[1].textContent,'不要');
 const name=control('項目名');name.value='保存後に変更';name.oninput();await nodes.get('req-save').onclick();
 assert.equal(calls.length,2);assert.equal(calls[1].fields[0].enabled,false);assert.equal(calls[1].fields[0].label,'保存後に変更');assert.equal(saved.revision,2);
});
test('missing AI settings and API failures replace the progress beside the analysis button',async()=>{
 for(const configured of [false,true]){
  const nodes=new Map(),timers=new Set();let analysisCalls=0;
  const el=tag=>({tagName:tag.toUpperCase(),value:'',checked:false,style:{},children:[],setAttribute(k,v){this[k]=v},append(...x){this.children.push(...x)},replaceChildren(...x){this.children=x},add(x){this.children.push(x)}});
  const document={getElementById:id=>{if(!nodes.has(id))nodes.set(id,el('input'));return nodes.get(id)},createElement:el,querySelectorAll:()=>[]};
  const handlers={listRequirementForms:()=>[],listGrantMaterialBatches:()=>[{id:'batch',state:'completed',title:'test',year:'R8',files:[{name:'test.pdf'}]}],listGrantInvitations:()=>[],getRequirementAiStatus:()=>({configured}),analyzeRequirementMaterial:()=>{analysisCalls++;throw Error('AI解析に失敗しました（HTTP 429）。');}};
  const run=()=>{let ok,fail;return new Proxy({withSuccessHandler(fn){ok=fn;return this},withFailureHandler(fn){fail=fn;return this}},{get(target,name){if(name in target)return target[name];return (...args)=>Promise.resolve().then(()=>handlers[name](...args)).then(ok,fail)}})};
  const api={};Object.defineProperty(api,'run',{get:run});
  const ctx={document,google:{script:api},Option:function(text,value){this.text=text;this.value=value},crypto,structuredClone,Date,console,confirm:()=>true,setInterval(fn){timers.add(fn);return fn;},clearInterval(fn){timers.delete(fn);}};vm.createContext(ctx);
  const html=fs.readFileSync(path.join(__dirname,'../apps-script/RequirementEditor.html'),'utf8');vm.runInContext(html.match(/<script>([\s\S]*?)<\/script>/)[1],ctx);await new Promise(setImmediate);
  nodes.get('req-batch').value='batch';document.getElementById('req-consent').checked=true;await nodes.get('req-analyze').onclick();
  const text=nodes.get('req-analysis').textContent;assert(text.startsWith('解析を停止しました：'));assert(!text.includes('解析中'));assert(text.includes(configured?'HTTP 429':'未設定'));assert.equal(analysisCalls,configured?1:0);assert.equal(timers.size,0);
 }
});
