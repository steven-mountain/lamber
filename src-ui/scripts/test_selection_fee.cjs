const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const root = path.join(__dirname, '../src');
const mocks = {};
const cache = new Map();
function load(relative) {
  const filename = path.join(root, relative);
  if (filename.endsWith('.json')) return JSON.parse(fs.readFileSync(filename,'utf8'));
  if (cache.has(filename)) return cache.get(filename);
  const module = {exports:{}};
  const source = ts.transpileModule(fs.readFileSync(filename,'utf8'), {compilerOptions:{target:ts.ScriptTarget.ES2020,module:ts.ModuleKind.CommonJS,esModuleInterop:true}}).outputText;
  vm.runInNewContext(source,{module,exports:module.exports,console,require:name=>{
    if(mocks[name]) return mocks[name];
    if(!name.startsWith('.')) return require(name);
    const next = path.relative(root,path.resolve(path.dirname(filename),name));
    return load(path.extname(next)?next:`${next}.ts`);
  }});
  cache.set(filename,module.exports); return module.exports;
}
// Minimal hook driver: executes production callbacks/effects and asynchronous requests.
let slots=[], cursor=0, effects=[], cleanups=[], dirty=false, pending=[];
const equal=(a,b)=>a&&b&&a.length===b.length&&a.every((v,i)=>Object.is(v,b[i]));
mocks.react={
  useState:init=>{const id=cursor++;if(!(id in slots))slots[id]=typeof init==='function'?init():init;return [slots[id],value=>{slots[id]=typeof value==='function'?value(slots[id]):value;dirty=true;}];},
  useRef:init=>{const id=cursor++;return slots[id]??(slots[id]={current:init});},
  useCallback:(fn,deps)=>{const id=cursor++;if(!equal(slots[id]?.deps,deps))slots[id]={deps,fn};return slots[id].fn;},
  useEffect:(fn,deps)=>{const id=cursor++;if(!equal(slots[id],deps)){slots[id]=deps;effects.push(()=>{cleanups[id]?.();cleanups[id]=fn();});}},
};
mocks['@tauri-apps/api/core']={invoke:(command,args)=>new Promise((resolve,reject)=>pending.push({command,args,resolve,reject}))};
const {useSelectionFeeCalculator}=load('hooks/useSelectionFeeCalculator.ts');
let hook;
function render(){do {dirty=false;cursor=0;hook=useSelectionFeeCalculator();const next=effects;effects=[];next.forEach(fn=>fn());}while(dirty);return hook;}
async function settle(){await Promise.resolve();await Promise.resolve();render();}
const fee = load('lib/selectionFee.ts');
const {normalizeTaxPairFromIncl}=load('lib/taxAmount.ts');
const catalog=load('lib/ictSubjectCatalog.ts');
function getApply(selectionOverrides={}, tax=6, serviceTax=6) {
  const source=fs.readFileSync(path.join(root,'hooks/useIctCalculations.ts'),'utf8');
  const ast=ts.createSourceFile('hook.ts',source,ts.ScriptTarget.Latest,true);
  let declaration;
  function visit(node){if(ts.isVariableDeclaration(node)&&node.name.getText(ast)==='applySelectionLimit')declaration=node;ts.forEachChild(node,visit);}visit(ast);
  const alerts=[],batches=[];
  const state={revIt:{},revCt:{},revNonItCt:{incl:0,tax:9,excl:0},costIt:{integration:{incl:'0',tax},device:{incl:'0',tax:13},bidding:{incl:'777',tax:serviceTax}},costCt:{},costMix:{},updateTaxItemsInclBatch:items=>batches.push(items)};
  const context={...load('lib/ictTaxItemBatch.ts'),MONEY_EPSILON:0.004,...fee,...catalog,normalizeTaxPairFromIncl,state,selection:{selectionFeeReady:true,selectionFeeError:''},selectionFeeTargetSubjectCode:'cost_it_integration',selectionFeeMergeService:false,selLimit:'302429.95',selFee:'2429.95',alert:message=>alerts.push(message),isTaxInclAutoFixEnabled:()=>false,...selectionOverrides};
  vm.createContext(context);
  const code=ts.transpileModule(`globalThis.apply = ${declaration.initializer.getText(ast)}`,{compilerOptions:{target:ts.ScriptTarget.ES2020}}).outputText;
  vm.runInContext(code,context);return {apply:context.apply,alerts,batches,state};
}
async function main(){
  assert.equal(fee.calculateSelectionFeeWriteAmounts('302429.95','2429.95').targetIncl,300000);
  assert.equal(fee.calculateSelectionFeeWriteAmounts('302429.95','2429.95',true).targetIncl,302429.95);
  assert.equal(fee.getSelectionFeeTargetTaxError('集成服务',6),null);
  assert.match(fee.getSelectionFeeTargetTaxError('集成服务',9),/当前为 9%.*既有/);
  assert.match(fee.getSelectionFeeTargetTaxError('设备',13),/6%.*13%/);
  for(const code of ['cost_it_device','cost_it_integration','cost_ct_construction']) assert.ok(fee.SELECTION_FEE_TARGET_SUBJECTS.some(x=>x.subjectCode===code));
  const split=getApply();split.apply();assert.equal(split.alerts.length,0);assert.equal(split.batches.length,1);assert.equal(split.batches[0].length,2);assert.equal(split.batches[0][0].incl,300000);assert.equal(split.batches[0][1].incl,2429.95);
  const merged=getApply({selectionFeeMergeService:true});merged.apply();assert.equal(merged.alerts.length,0);assert.equal(merged.batches[0].length,1);assert.equal(merged.batches[0][0].incl,302429.95);assert.equal(merged.state.costIt.bidding.incl,'777');
  for(const test of [getApply({selectionFeeTargetSubjectCode:'cost_it_device'}),getApply({},9)]){test.apply();assert.match(test.alerts[0],/6%/);assert.equal(test.batches.length,0);}
  for(const merge of [false,true]) { const test=getApply({selLimit:merge?'240':'346',selFee:'106',selectionFeeMergeService:merge});test.apply();assert.equal(test.batches.length,0);assert.match(test.alerts[0],/不可精确表示/); }
  for (const merge of [false,true]) {
    const test=getApply({selLimit:merge?'240':'346',selFee:'106',selectionFeeMergeService:merge,isTaxInclAutoFixEnabled:()=>true});
    test.apply();assert.equal(test.batches.length,0);assert.match(test.alerts[0],/归一后.*已停止写入/);
  }
  for (const merge of [false,true]) { const test=getApply({selectionFeeMergeService:merge,isTaxInclAutoFixEnabled:()=>true});test.apply();assert.equal(test.alerts.length,0);assert.equal(test.batches.length,1); }
  const stale=getApply({selection:{selectionFeeReady:false,selectionFeeError:'无对应报价'}});stale.apply();assert.equal(stale.batches.length,0);assert.equal(stale.alerts[0],'无对应报价');
  render();assert.equal(hook.selectionFeeMergeService,false);assert.equal(hook.selectionFeeReady,false);
  hook.handleSelFeeChange('quote','300000');render();const first=pending.shift();assert.equal(first.command,'calculate_selection_fee');assert.equal(hook.selectionFeePending,true);
  hook.handleSelFeeChange('limit','106600');render();const second=pending.shift();assert.equal(second.command,'reverse_calculate_selection_fee');assert.equal(hook.selQuote,'');
  first.resolve({quote:'300000.00',final_limit:'302479.95',actual_cost:'302429.95',selection_fee_incl:'2429.95',selection_fee_excl:'2292.41',quote_candidates:[]});await settle();assert.equal(hook.selectionFeeReady,false);
  second.reject('该限价落在资费表 10 万元档位跳变形成的空档，无对应报价');await settle();assert.match(hook.selectionFeeError,/无对应报价/);assert.equal(hook.selFee,'');assert.equal(hook.selectionFeeReady,false);
  hook.handleSelFeeChange('quote','300000');render();const fresh=pending.shift();fresh.resolve({quote:'300000.00',final_limit:'302479.95',actual_cost:'302429.95',selection_fee_incl:'2429.95',selection_fee_excl:'2292.41',quote_candidates:[]});await settle();assert.equal(hook.selFee,'2429.95');assert.equal(hook.selectionFeeError,'');
  hook.setSelectionFeeMergeService(true);render();const saved=JSON.parse(JSON.stringify(hook.buildSelectionFeePayload()));assert.equal(saved.selection_fee_merge_service,true);
  hook.restoreSelectionFeeState(null);render();assert.equal(hook.selectionFeeMergeService,false);
  hook.restoreSelectionFeeState({...saved,selection_fee_amount:'9999'});render();assert.equal(hook.selectionFeeMergeService,true);assert.equal(hook.selFee,'');assert.equal(hook.selectionFeePending,true);const restored=pending.shift();assert.equal(restored.args.quote,'300000');
  // Leaving the project invalidates its pending calculation even when another input is identical.
  hook.restoreSelectionFeeState(null);render();restored.resolve({selection_fee_incl:'9999',quote_candidates:[]});await settle();assert.equal(hook.selFee,'');assert.equal(hook.selectionFeeReady,false);
  hook.handleSelFeeChange('quote','10000');render();pending.shift().resolve({selection_fee:'100.00'});await settle();
  assert.match(hook.selectionFeeError,/版本不匹配/);assert.equal(hook.selectionFeeReady,false);
  const view=fs.readFileSync(path.join(root,'views/IctLifecycle.tsx'),'utf8');
  assert.match(view,/selectionFeeError && <p role="alert"/);assert.match(view,/disabled=\{!selectionFeeReady\}/);assert.match(view,/SELECTION_FEE_TARGET_SUBJECTS.filter\(subject => subject.groupId === group.groupId\)/);
  console.log('Selection fee: split/merge production writes, current-rate rejection, unchanged service item, exact-tax guard, stale result/error isolation, restore/recompute and persisted mode passed.');
}
main().catch(error=>{console.error(error);process.exitCode=1;});
