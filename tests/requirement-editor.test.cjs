const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),crypto=require('node:crypto');
test('editing the same field after saving updates the newly returned draft object',async()=>{
 const nodes=new Map();
 const el=tag=>({tagName:tag.toUpperCase(),value:'',checked:false,style:{},children:[],append(...x){this.children.push(...x)},replaceChildren(...x){this.children=x},add(x){this.children.push(x)}});
 const document={getElementById:id=>{if(!nodes.has(id))nodes.set(id,el('input'));return nodes.get(id)},createElement:el,querySelectorAll:()=>[]};
 let saved=null;const calls=[];
 const handlers={listRequirementForms:()=>saved?[saved]:[],listGrantMaterialBatches:()=>[],listGrantInvitations:()=>[],getRequirementAiStatus:()=>({configured:false}),saveRequirementForm:input=>{calls.push(structuredClone(input));saved={...structuredClone(input),id:'11111111-1111-1111-1111-111111111111',revision:input.revision+1};return structuredClone(saved)}};
 const run=()=>{let ok,fail;return new Proxy({withSuccessHandler(fn){ok=fn;return this},withFailureHandler(fn){fail=fn;return this}},{get(target,name){if(name in target)return target[name];return (...args)=>Promise.resolve().then(()=>handlers[name](...args)).then(ok,fail)}})};
 const scriptRun={};Object.defineProperty(scriptRun,'run',{get:run});
 const ctx={document,google:{script:scriptRun},Option:function(text,value){this.text=text;this.value=value},crypto,structuredClone,Date,console,confirm:()=>true};vm.createContext(ctx);
 const html=fs.readFileSync(path.join(__dirname,'../apps-script/RequirementEditor.html'),'utf8');vm.runInContext(html.match(/<script>([\s\S]*?)<\/script>/)[1],ctx);await new Promise(setImmediate);
 nodes.get('req-title').value='動作確認';nodes.get('req-year').value='令和8年度';nodes.get('req-add').onclick();await nodes.get('req-save').onclick();
 const labels=nodes.get('req-fields').children[0].children;const control=title=>labels.find(x=>x.children?.[0]===title).children[1];
 const enabled=control('使用する（外すと申請者画面から非表示）');enabled.checked=false;enabled.onchange();
 const name=control('項目名');name.value='保存後に変更';name.oninput();await nodes.get('req-save').onclick();
 assert.equal(calls.length,2);assert.equal(calls[1].fields[0].enabled,false);assert.equal(calls[1].fields[0].label,'保存後に変更');assert.equal(saved.revision,2);
});
