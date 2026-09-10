const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const {generate, extract} = require('./test_template_state.cjs');
const root = path.join(__dirname, '../src');
const listeners = new Map();
const emitTo = async (_window, name, payload) => { for (const callback of [...(listeners.get(name) ?? [])]) callback({payload}); };
let binding = {workspaceId:'w',projectId:'p',projectName:'合成项目'};
let workspace = 'w';
let assets = [];
let state = {formData:{gen_demand_env_require:''}};
let duringRead;
const mocks = {
  '@tauri-apps/api/core': {invoke: async command => {
    if(command === 'ai_get_session_binding') return binding;
    if(command === 'get_available_templates') return ['ICT项目需求导入表.docx','ICT项目立项签批表.docx'];
    throw Error(command);
  }},
  '@tauri-apps/api/window': {getCurrentWindow:()=>({label:'ai-assistant'})},
  '@tauri-apps/api/event': {emitTo, listen:async (name,callback)=>{ const set=listeners.get(name)??new Set();set.add(callback);listeners.set(name,set);return ()=>set.delete(callback); }},
  'zustand': {create: initializer => {let value=initializer();return {getState:()=>value,setState:patch=>{value={...value,...patch};}};}},
  '../utils/workspaceService': {workspaceService:{getState:async()=>({currentWorkspace:{workspaceId:workspace}})}},
  './domainSaveService': {domainSaveService:{loadTemplateState:async()=>{await duringRead?.();return {filledDataJson:state};},loadTemplateAssets:async()=>assets}},
  '../utils/projectService': {projectService:{getTemplateAssetPath:async id=>{if(id==='missing')throw Error('missing');return 'path';}}},
};
const cache = new Map();
function load(relative){
  const filename=path.join(root,relative);
  if(filename.endsWith('.json'))return JSON.parse(fs.readFileSync(filename,'utf8'));
  if(cache.has(filename))return cache.get(filename);
  const mod={exports:{}};cache.set(filename,mod.exports);
  const code=ts.transpileModule(fs.readFileSync(filename,'utf8'),{compilerOptions:{module:ts.ModuleKind.CommonJS,target:ts.ScriptTarget.ES2020,esModuleInterop:true}}).outputText;
  vm.runInNewContext(code,{exports:mod.exports,module:mod,console,crypto:require('node:crypto').webcrypto,setTimeout,clearTimeout,require:name=>{
    if(mocks[name])return mocks[name];
    const next=path.relative(root,path.resolve(path.dirname(filename),name));return load(path.extname(next)?next:`${next}.ts`);
  }});cache.set(filename,mod.exports);return mod.exports;
}
async function main(){
  const api=load('services/chatDocumentGeneration.ts');
  const intent=load('ai/documentGenerationIntent.ts');
  const ask=text=>intent.documentTemplateRequests([{role:'user',content:text}]);
  assert.equal(ask('帮我生成需求导入表')[0],'demand');
  assert.equal(ask('请生成立项签批表')[0],'approval');
  for(const text of ['不要生成需求导入表','怎么生成需求导入表','“帮我生成需求导入表”','你好'])assert.equal(ask(text).length,0,text);
  assert.equal(intent.documentTemplateRequests([{role:'user',content:'生成需求导入表'},{role:'user',content:'谢谢'},{role:'assistant',content:'生成需求导入表'}]).length,0);
  let targets=await api.loadDocumentTargets('session',['demand']);
  assert.equal(targets.length,1);assert.ok(api.completionSummary(targets[0].completion).missing.some(x=>x.label==='部署环境要求'));
  const summary=api.completionSummary((await api.loadDocumentTargets('session',['approval']))[0].completion);
  assert.equal(summary.unknown.length,3);assert.equal(summary.filledCount,4);
  assets=[{id:'missing',usage:'attach1'}];targets=await api.loadDocumentTargets('session',['demand']);
  assert.ok(api.completionSummary(targets[0].completion).missing.some(x=>x.key==='attach1'));
  const original=binding;binding=null;assert.equal((await api.loadDocumentTargets('session',['demand'])).length,0);
  binding={...original,projectId:null};assert.equal((await api.loadDocumentTargets('session',['demand'])).length,0);
  binding=original;duringRead=()=>{workspace='switched';};await assert.rejects(api.loadDocumentTargets('session',['demand']),/变更/);workspace='w';duringRead=null;
  const stop=await api.listenDocumentRequests(async request=>{await api.finishDocumentRequest(request,{status:'success',outputDir:'/tmp/真实目录'});});
  const result=await api.requestDocumentGeneration(targets[0]);assert.equal(result.status,'success');assert.match(api.documentReceipt(targets[0],result),/真实目录/);stop();
  const stopError=await api.listenDocumentRequests(async()=>{throw Error('原始错误：附件图片文件缺失');});
  const error=await api.requestDocumentGeneration(targets[0]);assert.match(error.message,/原始错误：附件图片文件缺失/);stopError();
  const fields={gen_demand_env_require:'机房提供独立电源',gen_demand_service_content:'本轮产品入口一致性测试'};
  const page=await generate('ICT项目需求导入表.docx',fields);
  let rechecked=0;
  const card=await generate('ICT项目需求导入表.docx',fields,{beforeDocumentGenerate:async()=>{rechecked++;}});
  assert.equal(JSON.stringify(page.captured),JSON.stringify(card.captured));assert.equal(rechecked,1);assert.equal(card.result.status,'success');
  const missing=await generate('ICT项目需求导入表.docx',{...fields,gen_demand_env_require:''});assert.equal(missing.result.status,'success');
  const failure=await generate('ICT项目需求导入表.docx',fields,{loadDemandImages:async()=>({attach1:[{assetId:'gone',error:true}],attach2:[]})});
  assert.equal(failure.result.status,'error');assert.equal(failure.result.message,failure.alerts[0]);assert.match(failure.result.message,/附件图片文件缺失/);
  const engine=await generate('ICT项目需求导入表.docx',fields,{invoke:async()=>{throw Error('Permission denied: /locked');}});assert.match(engine.result.message,/Permission denied: \/locked/);
  const cancel=await generate('ICT项目需求导入表.docx',fields,{invoke:async()=>{throw 'FILE_EXISTS::/tmp/existing.docx';}});assert.equal(cancel.result.status,'cancelled');
  const ast=ts.createSourceFile('calculations.ts',fs.readFileSync(path.join(root,'hooks/useIctCalculations.ts'),'utf8'),ts.ScriptTarget.Latest,true);
  let latestInput={value:1}, output, calculationState;
  const pending=[];
  const context={useCallback:fn=>fn, calculationSequence:{current:0},getInputDataPayload:()=>latestInput,
    invoke:async()=>new Promise(resolve=>pending.push(resolve)),setMetrics:value=>{output=value;},setCashflowTable:()=>{},setCalculationState:value=>{calculationState=value;},console};
  vm.createContext(context);vm.runInContext(extract('performCalculation',ast),context);
  const old=context.extracted();latestInput={value:2};const fresh=context.extracted();
  pending[1]({npv:200,cashflow:[]});await fresh;pending[0]({npv:100,cashflow:[]});await old;
  assert.equal(output.npv,200,'late old calculation cannot replace the current project result');assert.equal(calculationState.key,JSON.stringify(latestInput));
  const panel=fs.readFileSync(path.join(root,'components/ai/AiChatPanel.tsx'),'utf8');
  const condition=panel.match(/\{(documentTemplateIds.length > 0[^\n]+) && \(\n\s*<DocumentGenerationCards/)[1];
  const visible=(bound,ready,ids=['demand'])=>vm.runInNewContext(condition,{documentTemplateIds:ids,bindingReady:ready,bindingState:{binding:{projectId:bound}},currentSessionId:'session'});
  assert.ok(visible('p',true));assert.ok(!visible(null,true));assert.ok(!visible('p',false));assert.ok(!visible('p',true,[]));
  const expiredStop=await api.listenDocumentRequests(async()=>{throw Error('expired request must not run');});
  await emitTo('main','lamber-chat-document-request',{...targets[0],requestId:'expired',replyWindow:'ai-assistant',expiresAt:0});
  assert.equal(api.useDocumentRequest.getState().request,null);expiredStop();
  console.log('Chat document: intent, trusted targets, missing/unknown, stale binding, cross-window receipts, shared production generator parity, original failure and overwrite cancellation passed.');
}
main().catch(error=>{console.error(error);process.exitCode=1;});
