const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), vm = require('node:vm');
const ts = require('typescript'), { webcrypto } = require('node:crypto');
const { run } = require('./test_structure_reverse_success.cjs');
const root = path.join(__dirname, '../src'), listeners = new Map();
let binding = { workspaceId: 'w', projectId: 'p' }, workspace = 'w';
const project = { id:'p', name:'结构反算合成项目', default_scheme_id:'pre' };
const schemes = [{ id:'pre', project_id:'p', name:'方案甲', stage:'pre_selection', updated_at:'2', created_at:'1' },
  { id:'post', project_id:'p', name:'方案乙', stage:'post_selection', updated_at:'3', created_at:'2' }];
const emitTo = async (window, name, payload) => { for (const fn of [...(listeners.get(name) || [])]) fn({ payload }); };
const mocks = {
  '@tauri-apps/api/core': { invoke: async command => { assert.equal(command, 'ai_get_session_binding'); return binding; } },
  '@tauri-apps/api/window': { getCurrentWindow: () => ({ label: 'ai-assistant' }) },
  '@tauri-apps/api/event': { emitTo, listen: async (name, fn) => { const set = listeners.get(name) || new Set(); set.add(fn); listeners.set(name,set); return () => set.delete(fn); } },
  zustand: { create: initial => { let state=initial(); return { getState:()=>state, setState:patch=>{state={...state,...patch};} }; } },
  '../utils/workspaceService': { workspaceService: { getState: async () => ({ currentWorkspace: { workspaceId:workspace } }) } },
  '../utils/projectService': { projectService: { getProject:async()=>project,getSchemes:async()=>schemes } },
};
const cache = new Map();
function load(file) {
  file=path.resolve(root,file);
  if(file.endsWith('.json'))return JSON.parse(fs.readFileSync(file,'utf8'));
  if(cache.has(file))return cache.get(file);
  const mod={exports:{}};cache.set(file,mod.exports);
  const code=ts.transpileModule(fs.readFileSync(file,'utf8'),{compilerOptions:{module:ts.ModuleKind.CommonJS,target:ts.ScriptTarget.ES2020,esModuleInterop:true}}).outputText;
  vm.runInNewContext(code,{module:mod,exports:mod.exports,console,crypto:webcrypto,setTimeout,clearTimeout,require:name=>{
    if(mocks[name])return mocks[name];
    if(!name.startsWith('.'))return require(name);
    const resolved=path.resolve(path.dirname(file),name);return load(path.extname(resolved)?resolved:resolved+'.ts');
  }});return mod.exports;
}
async function main() {
  const api=load('services/chatStructureReverse.ts'), intent=load('ai/structureReverseIntent.ts');
  const receiptContext=load('ai/appReceiptContext.ts').appReceiptContext;
  const resultLib=load('lib/structureReverseResult.ts');
  const ask=text=>intent.structureReverseIntent([{role:'user',content:text}]);
  assert.equal(ask('把利润率做到12%').targetPercent,'12');
  assert.equal(ask('我要做智能结构反算').targetPercent,'');
  assert.equal(ask('我要做智能结构反算，暂时没决定目标值，不要替我选择科目或目标。').requested,true);
  assert.equal(ask('把净现值率调整到8%').metricType,'npv_rate');
  for(const text of ['你好','不要把利润率做到12%','“把利润率做到12%”','甄选费反算','```把利润率做到12%```'])assert.equal(ask(text).requested,false,text);
  assert.equal(intent.structureReverseIntent([{role:'user',content:'谢谢'},{role:'assistant',content:'把利润率做到15%'}]).requested,false);
  assert.equal(ask('利润率从8%做到12%').targetPercent,'','ambiguous percentages must not guess');
  for (const text of ['我要做结构反算，不要设为12%', '我要做结构反算，上次目标12%', '我要做结构反算，设备税率13%']) {
    assert.equal(ask(text).requested,true);
    assert.equal(ask(text).targetPercent,'','negated, historical and tax percentages are not confirmed targets');
  }
  assert.equal(ask('刚才结构反算为什么没成功？请把应用原错误完整告诉我，不要重新查询或试算。').requested,false);
  assert.equal(ask('请再发起一次甄选前方案的结构反算，毛利润率目标是20.005%，让我自己选择科目。').targetPercent,'20.005');
  assert.equal((await api.loadReverseProject('s')).project.id,'p');
  binding={workspaceId:'w',projectId:null};assert.equal(await api.loadReverseProject('s'),null);
  binding=null;assert.equal(await api.loadReverseProject('s'),null);
  binding={workspaceId:'w',projectId:'p'};workspace='other';await assert.rejects(api.loadReverseProject('s'),/变更/);workspace='w';
  for(const selector of ['pre','pre_selection','方案甲',''])assert.equal(api.selectReverseScheme(project,schemes,selector).id,'pre');
  assert.equal(api.selectReverseScheme(project,schemes,'post_selection').id,'post');
  assert.throws(()=>api.selectReverseScheme(project,schemes,'不存在'),/没有匹配/);
  const request={sessionId:'s',workspaceId:'w',projectId:'p',projectName:project.name,schemeId:'pre',subjectCode:'rev_it_integration',metricType:'margin',action:'apply',target:0.1};
  for(const changed of [{schemeId:'post'},{subjectCode:'cost_it_device'},{metricType:'npv_rate'},{sessionId:'other'},{workspaceId:'other'}]){
    const token=api.issueStructurePreview(request,'stamp');assert.throws(()=>api.consumeStructurePreview({...request,...changed,token},'stamp'),/预览已失效/);
  }
  let token=api.issueStructurePreview(request,'stamp');assert.throws(()=>api.consumeStructurePreview({...request,token},'changed input'),/预览已失效/);
  token=api.issueStructurePreview(request,'stamp');api.consumeStructurePreview({...request,token},'stamp');assert.throws(()=>api.consumeStructurePreview({...request,token},'stamp'),/预览已失效/);
  const cases=[];
  for(const [name, options] of [
    ['insensitive',{metric:()=>.094}],['out_of_range',{target:.12,metric:a=>a<50?.094:.108}],
    ['no_convergence',{metric:a=>a<50?.094:.108}],['final_check',{metric:(a,n,samples)=>n>samples?.108:a/400}],
  ]) {
    const native=await run(options);
    let nativeCalls;
    const stop=await api.listenStructureRequests(async r=>{
      const card=await run({...options,cardOptions:{silent:true}});nativeCalls=card.batch.mock.callCount();
      await api.finishStructureRequest(r,card.response);
    });
    const reply=await api.requestStructureReverse(request);stop();
    assert.equal(nativeCalls,0,`${name}: IPC card path writes ZERO times`);
    assert.equal(reply.message,native.alerts[0],`${name}: entire original error is unchanged`);
    const receipt=api.structureReceipt(project.name,reply);
    assert.ok(receipt.endsWith(native.alerts[0]));
    const context=receiptContext([{role:'assistant',content:receipt,appReceipt:true}]);
    assert.equal(context[0].content[0].result,receipt,'next model turn gets the complete application receipt');
    cases.push({name,message:reply.message,receipt,updateTaxItemsInclBatchCalls:nativeCalls});
  }
  const catalog=load('lib/ictSubjectCatalog.ts').ICT_SUBJECT_DEFINITIONS;
  const before={subject_funding_plans:{}},after={subject_funding_plans:{}};
  catalog.forEach((subject,i)=>{before[subject.subjectCode]={incl_tax:String(100+i),custom_subject_name:`业务${i}`,billing_subject_name:`开票${i}`};after[subject.subjectCode]={...before[subject.subjectCode]};});
  const s=catalog[0],id=`${s.side}:${s.groupId}:${s.key}`;
  before.subject_funding_plans[id]={id,enabled:true,mode:'custom',annualInclValues:[25,25,25,25,0,0,0,0,0,0]};
  after.subject_funding_plans[id]={...before.subject_funding_plans[id],annualInclValues:[30.01,30.01,30.01,29.99,0,0,0,0,0,0]};after[s.subjectCode].incl_tax='120.02';
  const changes=resultLib.structureSubjectChanges(before,after);
  assert.equal(changes.length,1);assert.equal(changes[0].code,s.subjectCode);assert.match(changes[0].name,/开票0/);
  const success={status:'success',message:'结构反算完成',metricType:'margin',target:.12,achieved:.12001,targetReached:true,changes,schemeName:'方案甲',stage:'pre_selection',snapshotVersion:1};
  const successReceipt=api.structureReceipt(project.name,success);
  for(const expected of ['目标值 | 实际达成值','12.0000% | 12.0010%','100.00 → 120.02','第4年：25.00 → 29.99','第10年：0.00 → 0.00'])assert.ok(successReceipt.includes(expected),expected);
  const evidence={prompt:intent.structureReversePrompt(true),cases,success:{...success,receipt:successReceipt}};
  const file=path.resolve(root,'../../src-tauri/src/agent_bridge/fixtures/structure-reverse-receipts.json');
  if(process.argv.includes('--write-fixtures'))fs.writeFileSync(file,JSON.stringify(evidence,null,2)+'\n');
  else assert.deepEqual(JSON.parse(JSON.stringify(evidence)),JSON.parse(fs.readFileSync(file,'utf8')));
  console.log('D: original four range errors → IPC → chat → appReceipt exactly preserved; every failure writes zero. Binding, selectors, single-use preview freshness, all-28 diff and annual tails passed.');
}
main().catch(error=>{console.error(error);process.exitCode=1;});
