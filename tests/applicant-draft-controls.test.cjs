const test=require('node:test'),assert=require('node:assert/strict'),vm=require('node:vm'),fs=require('node:fs'),path=require('node:path');
function page(preview,snapshotExtra={}){
 const nodes=new Map(),timers=new Map();let next=0,saveCalls=0,pending;
 const el=tag=>({tagName:tag.toUpperCase(),style:{},children:[],value:'',append(...x){this.children.push(...x)},add(x){this.children.push(x)},addEventListener(){},setAttribute(){}});
 const get=id=>{if(!nodes.has(id))nodes.set(id,el('div'));return nodes.get(id);};
 get('answers').elements={inviteToken:{},revision:{},answersJson:{}};
 const initial={preview,submitted:false,token:'test-token',values:{},snapshot:{revision:1,title:'test',year:'R8',fields:[],...snapshotExtra}};
 const run=()=>{let ok,fail;return new Proxy({withSuccessHandler(fn){ok=fn;return this},withFailureHandler(fn){fail=fn;return this}},{get(target,name){if(name in target)return target[name];return input=>{if(name==='getRequirementApplicantState')ok(initial);else if(name==='saveRequirementAnswerDraft'){saveCalls++;pending={ok,fail};}};}});};
 const api={};Object.defineProperty(api,'run',{get:run});
 const ctx={document:{getElementById:get,createElement:el},google:{script:api},Option:function(){},setTimeout(fn,ms){const id=++next;timers.set(id,{fn,ms});return id;},clearTimeout(id){timers.delete(id);}};
 vm.createContext(ctx);const html=fs.readFileSync(path.join(__dirname,'../apps-script/RequirementApplicant.html'),'utf8');const script=html.match(/<script>([\s\S]*?)<\/script>/)[1].replace('<?!= initial ?>',JSON.stringify(initial));vm.runInContext(script,ctx);
 return {get,timers,calls:()=>saveCalls,pending:()=>pending};
}
test('preview explains why temporary save is disabled and never sends answers',()=>{const p=page(true);assert(p.get('save-draft-top').disabled);assert(p.get('save-draft-top').textContent.includes('プレビュー'));assert.equal(p.get('save-draft-top').style.cursor,'not-allowed');p.get('save-draft-top').onclick();assert.equal(p.calls(),0);});
test('applicant draft controls recover after timeout and after save error',()=>{const p=page(false);assert.equal(p.get('save-draft-top').disabled,false);p.get('save-draft-top').onclick();assert.equal(p.calls(),1);assert(p.get('save-draft-top').disabled);p.get('save-draft-top').onclick();assert.equal(p.calls(),1);const timeout=[...p.timers.values()].find(t=>t.ms===45000);timeout.fn();assert.equal(p.get('save-draft-top').disabled,false);assert(p.get('draft-status').textContent.includes('入力内容は画面に残っています'));p.get('save-draft-top').onclick();p.pending().fail({message:'test failure'});assert.equal(p.get('save-draft-bottom').disabled,false);p.get('save-draft-bottom').onclick();p.pending().ok();assert(p.get('draft-status').textContent.includes('一時保存しました'));});

test('heading displays fiscal year once even if title repeats it',()=>{const p=page(true,{year:'令和8年度',title:'令和8年度 令和8年度 女性活躍情報公開促進奨励金'});assert.equal(p.get('title').textContent,'令和8年度 女性活躍情報公開促進奨励金');});
test('company field can be corrected by applicant',()=>{const p=page(false,{fields:[{id:'company',label:'会社名',type:'text',enabled:true}]});const input=p.get('fields').children[0].children[1];assert.notEqual(input.readOnly,true);assert.equal(input.maxLength,120);});
