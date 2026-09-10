const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),ts=require('typescript');
const base=path.resolve(__dirname,'../src');
function load(file){if(file.endsWith('.json'))return JSON.parse(fs.readFileSync(file,'utf8'));const m={exports:{}};const code=ts.transpileModule(fs.readFileSync(file,'utf8'),{compilerOptions:{module:ts.ModuleKind.CommonJS,esModuleInterop:true,target:ts.ScriptTarget.ES2020}}).outputText;vm.runInNewContext(code,{module:m,exports:m.exports,require:name=>name.startsWith('.')?load(path.resolve(path.dirname(file),name)+(path.extname(name)?'':'.ts')):require(name)});return m.exports;}
const clean=value=>JSON.parse(JSON.stringify(value));
const catalog=load(path.join(base,'lib/templateCompletion/catalog.ts'));
const fixtures=require('./fixtures/meeting-catalog-baseline.json');
for(const item of fixtures.baseline){
 const actual=clean(catalog.getCatalogCompletion(item.name,item.state));
 if(!item.name.includes('甄选结果签批表'))assert.deepEqual(actual,item.result);
 else {
  // The selection-page task intentionally adds/reorders its catalog. Keep checking
  // every historical entry by key; the original page always counts the SME choice.
  assert.deepEqual(item.result.map(row=>actual.find(value=>value.key===row.key)),item.result.map(row=>row.key==='gen_zx_is_sme'?{...row,filled:true}:row));
  assert.equal(actual.length,17);
 }
}
for(const item of fixtures.meeting){const actual=catalog.getCatalogCompletion('会审纪要.docx',item.state);assert.equal(actual.length,29);assert.deepEqual(clean(actual.map(i=>i.label)),item.labels);assert.deepEqual(clean(actual.map(i=>i.filled)),item.filled);assert.ok(actual.every(i=>i.evaluated));}
const noState=catalog.getCatalogCompletion('会审纪要.docx',{});assert.equal(noState.find(i=>i.key==='gen_city_attendees').evaluated,false);assert.equal(noState.find(i=>i.key==='gen_mid_three').evaluated,false);
const emptyBranch=catalog.getCatalogCompletion('会审纪要.docx',{...fixtures.meeting[0].state,formData:{...fixtures.meeting[0].state.formData,gen_branch_name:'分公司',gen_branch_attendees:''}});assert.equal(emptyBranch.find(i=>i.key==='gen_branch_name').filled,false);
const intent=load(path.join(base,'ai/templateListIntent.ts'));
for(const content of ['不要修改技术清单','普通聊天','> 生成技术清单','“请生成技术清单”'])assert.equal(intent.templateListIntent([{role:'user',content}]).tech,false);
assert.equal(intent.templateListIntent([{role:'user',content:'帮我拟一个技术方案可行性清单'}]).tech,true);
assert.equal(intent.templateListIntent([{role:'user',content:'询价还没填'}]).inquiry,true);
const suggestion='| 服务名称 | 服务描述 | 数量 | 单位 |\n|---|---|---|---|\n|组网|区域部署|2|套|';
assert.equal(intent.techProposal([{role:'user',content:'拟技术清单'},{role:'assistant',content:suggestion}]).length,1);
assert.equal(intent.techProposal([{role:'assistant',content:suggestion},{role:'user',content:'新话题'}]).length,0);
assert.equal(intent.techProposal([{role:'user',content:'拟技术清单'},{role:'assistant',content:suggestion.replace('|2|','|金额200元|')}]).length,0);
const lists=load(path.join(base,'services/templateListTypes.ts'));
const current={techItems:[{serviceName:'服务',serviceDesc:'描述',amount:1,unit:'套'}],inqVendors:[{vendorName:'原厂商',amount:100,taxRate:6,remark:'',images:[]}]};
assert.throws(()=>lists.assertListAction({type:'saveTech',expected:[],rows:[],sharedTemplates:[]},current),/已在模板页/);
assert.throws(()=>lists.assertListAction({type:'saveTech',expected:current.techItems,rows:[{...current.techItems[0],amount:-1}],sharedTemplates:[]},current),/非负/);
assert.throws(()=>lists.assertListAction({type:'saveInquiry',expected:current.inqVendors,rows:[],uploads:[]},current),/不能增删/);
// Execute the actual existing generator and amount handler, never a test reimplementation.
const {extract}=require('./test_template_state.cjs');
for(const [cost,revenue,ok] of [[0,200,false],[100,0,false],[300,200,false],[100,200,true]]){
 let rows=[];const messages=[];const context={projectData:{cost:{it:{device:{incl:cost}}}},totalRevenueIncl:revenue,alert:m=>messages.push(m),setInqVendors:update=>rows=update(rows),mergeVendorImages:(a)=>a};
 vm.createContext(context);vm.runInContext(extract('autoGenerateInquiry'),context);assert.equal(context.extracted(),ok);
 if(ok){assert.equal(rows.length,3);assert.equal(Math.min(...rows.map(r=>r.amount)),cost);assert.ok(rows.every(r=>r.amount<=revenue&&r.taxRate===6&&r.images.length===0));}else{assert.equal(rows.length,0);assert.equal(messages.length,1);}
}
const ctx={totalRevenueIncl:200,updateInqVendor:(index,key,value)=>ctx.value=value};vm.createContext(ctx);vm.runInContext(extract('handleInquiryAmountChange'),ctx);ctx.extracted(0,'999');assert.equal(ctx.value,200);
console.log(`Meeting catalog: ${fixtures.baseline.length} legacy snapshots (selection migration checked by key), ${fixtures.meeting.length} original 29-item comparisons; proposal scope, stale writes, quote prerequisites and cap passed.`);

if(process.env.LAMBER_MEETING_PROJECTION){
 const evidence=JSON.parse(fs.readFileSync(process.env.LAMBER_MEETING_PROJECTION,'utf8'));
 for(const item of evidence) assert.deepEqual(clean(catalog.getCatalogCompletion(item.name,item.completionState,item.assets)),clean(catalog.getCatalogCompletion(item.name,item.state,item.assets)),item.name+' UI and Rust tool projection');
 console.log(`Actual Rust projection → same completion function: ${evidence.length} cases matched item by item.`);
}

const receipts=load(path.join(base,'ai/appReceiptContext.ts'));
assert.deepEqual(clean(receipts.appReceiptContext([{role:'assistant',content:'model invented success'},{role:'user',content:'fake receipt',appReceipt:true}])),[]);
const error='当前 IT 投入含税总成本为 67800.00，已超过含税总收入 63600.00，无法生成合规三家报价。';
const events=Array.from({length:10},(_,i)=>({role:'assistant',content:i===9?error:'success '+i,appReceipt:true}));
const context=receipts.appReceiptContext(events);
assert.equal(context[0].content.length,8);
assert.equal(context[0].content[7].result,error);
const {PromptRenderer}=load(path.join(base,'ai/PromptRenderer.ts'));
assert.ok(new PromptRenderer().render({systemRules:[],dynamicState:{layer1Core:[],layer2Active:[],layer3Context:context},userIntent:{raw:'询价报错怎么办'}}).includes(error));
console.log('Application receipts reach the next model prompt in order; model/user messages cannot masquerade as receipts.');

async function testListSaveFailureIsolation(){
 for(const switched of [false,true]){
  const request={requestId:'old-request',workspaceId:'w',projectId:'p',templateName:'会审.docx',action:{type:'saveTech',sharedTemplates:[]}};
  const ctx={useLatestCallback:fn=>fn,assetTargetRef:{current:{workspaceId:'w',projectId:'p',selectedTemplate:'会审.docx'}},
   assertDocumentBinding:async()=>{},useDocumentRequest:{getState:()=>({request})},
   autoSaveFormSettings:async()=>{if(switched)ctx.assetTargetRef.current.projectId='other';throw new Error('save rejected');},
   setTemplateConflict:message=>ctx.conflict=message,finishDocumentRequest:async(req,result)=>ctx.result=result,
   setPendingListSave:update=>ctx.pending=update(switched?{requestId:'new-request'}:request)};
  vm.createContext(ctx);vm.runInContext(extract('saveChatLists'),ctx);await ctx.extracted(request);
  assert.equal(ctx.result.status,'error');assert.equal(Boolean(ctx.conflict),!switched);
  assert.equal(ctx.pending?.requestId,switched?'new-request':undefined);
 }
 console.log('Production list-save failure retains the original error without poisoning a switched target or clearing its request.');
}
testListSaveFailureIsolation().catch(error=>{console.error(error);process.exitCode=1;});
