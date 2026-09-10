const assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path'), vm = require('node:vm');
const { mock } = require('node:test'), ts = require('typescript'), load = require('./load_ts.cjs');
const root = path.resolve(__dirname, '../src');
const frozen = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures/structure-commit-before-normalization.json'), 'utf8'));
const catalog = load(path.join(root, 'lib/ictSubjectCatalog.ts'));
const tax = load(path.join(root, 'lib/taxAmount.ts'));
const fund = load(path.join(root, 'lib/ictSubjectFundingPlan.ts'));
const batch = load(path.join(root, 'lib/ictTaxItemBatch.ts'));
const candidate = load(path.join(root, 'lib/ictStructureCandidate.ts'));
const clone = value => JSON.parse(JSON.stringify(value));
const clean = value => JSON.parse(JSON.stringify(value, (key, value) => ['updatedAt', 'lastChangedAt'].includes(key) ? undefined : value));
function declarations(source, names) {
  const ast = ts.createSourceFile('production.ts', source, ts.ScriptTarget.Latest, true), found = {};
  function visit(node) { if (ts.isVariableDeclaration(node) && names.includes(node.name.getText(ast))) found[node.name.getText(ast)] = node.initializer.getText(ast); ts.forEachChild(node, visit); }
  visit(ast); assert.equal(Object.keys(found).length, names.length); return found;
}
const compile = code => ts.transpileModule(code, { compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.CommonJS } }).outputText;
const current = fs.readFileSync(path.join(root, 'hooks/useIctCalculations.ts'), 'utf8');
const currentBatch = declarations(fs.readFileSync(path.join(root, 'hooks/useIctState.ts'), 'utf8'), ['updateTaxItemsInclBatch']);
const names = ['getPairedCostSubjectForRevenueSubject','buildSubjectFundingCoverageSubjects','buildInputDataPayload','getCurrentReverseSubjectState','buildCandidateSyncUpdates',
  'buildLockedTotalStructureCandidate','performLockedTotalStructureReverseCalculation','getMetricValue','METRIC_EPSILON','MONEY_EPSILON','roundMoney','formatCurrency','formatPercent'];
function initial() {
  const state = { revIt:{}, revCt:{}, revNonItCt:{}, costIt:{}, costCt:{}, costMix:{}, subjectFundingPlans:{}, cashflowSegments:[],
    cashflowModel:'model_a', distRev:[1,...Array(9).fill(0)],distCost:[1,...Array(9).fill(0)],ignoredDataHash:null,ignoredTailValue:null,
    balanceAllocation:{revenue:{enabled:false},investment:{enabled:false}},projectYears:4,discountRate:0.055 };
  for (const s of catalog.ICT_SUBJECT_DEFINITIONS) {
    const item={incl:0,excl:0,tax:s.defaultTaxRate,customSubjectName:`业务-${s.subjectCode}`,billingSubjectName:`开票-${s.subjectCode}`};
    if(s.groupId==='revNonItCt')state.revNonItCt=item;else state[s.groupId][s.key]=item;
  }
  for(const [key,amount,rate] of [['integration',40,13],['maintenance',60,6]]) {
    Object.assign(state.revIt[key],{incl:amount,excl:tax.exclFromIncl(amount,rate),tax:rate});
    const plan=fund.createDefaultSubjectFundingPlan({side:'revenue',groupId:'revIt',key},amount);
    state.subjectFundingPlans[plan.id]={...plan,mode:'custom',annualInclValues:[amount/4,amount/4,amount/4,amount/4,...Array(6).fill(0)]};
  }
  return state;
}
function host(autoFix, old=false, input=initial(), metric=payload=>Number(payload.rev_it_integration.incl_tax)/400) {
  const state=clone(input), alerts=[],evaluations=[];
  const ctx={...tax,...fund,...catalog,...batch,...candidate,...load(path.join(root,'lib/ictReverseCalculation.ts')),
    ...load(path.join(root,'lib/ictCalculationInput.ts')),...load(path.join(root,'lib/ictTaxItemEdit.ts')),
    state,...state,defaultTaxItem:rate=>({incl:0,excl:0,tax:rate}),isTaxInclAutoFixEnabled:()=>autoFix,
    effectiveDistRev:state.distRev,effectiveDistCost:state.distCost,buildSelectionFeePayload:()=>({}),
    buildDirectCashflowFromSegments:()=>({rev:[],cost:[]}),serializeBalanceAllocationRule:load(path.join(root,'lib/ictBalanceAllocation.ts')).serializeBalanceAllocationRule,
    revTargetType:'margin',revTargetValue:0.10325,alert:message=>alerts.push(message),setCashflowTable:()=>{},setMetrics:()=>{},updateData:()=>{},
    AI_CONTEXT_KEY:{ICT_CORE:'ict_core'},buildAiContextPayload:()=>({}),setCashflowCalculationSourceState:()=>{},
    invoke:async(command,{input})=>{assert.equal(command,'calculate_ict_benefit');evaluations.push(clone(input));return {margin_rate:metric(input),npv_rate:metric(input),cashflow:[]};},
  };
  for(const group of ['revIt','revCt','revNonItCt','costIt','costCt','costMix']) ctx['set'+group[0].toUpperCase()+group.slice(1)]=value=>{ctx[group]=value;state[group]=value;};
  ctx.setSubjectFundingPlansState=fn=>{state.subjectFundingPlans=fn(state.subjectFundingPlans);};
  state.setActiveTab=()=>{};state.setIgnoredDataHash=()=>{};state.setIgnoredTailValue=()=>{};
  const funcs=declarations(old?frozen.calculationHook:current,names);
  vm.createContext(ctx);
  vm.runInContext(compile((old?frozen.batchHelpers+frozen.batch:`const updateTaxItemsInclBatch=${currentBatch.updateTaxItemsInclBatch};`)
    +Object.entries(funcs).map(([k,v])=>`const ${k}=${v};`).join('\n')
    +'\nglobalThis.batchSetter=updateTaxItemsInclBatch;globalThis.build=buildLockedTotalStructureCandidate;globalThis.payload=buildInputDataPayload;globalThis.solve=performLockedTotalStructureReverseCalculation;'),ctx);
  state.updateTaxItemsInclBatch=mock.fn(ctx.batchSetter);
  return {ctx,state,alerts,evaluations,spy:state.updateTaxItemsInclBatch};
}
function structureFor(state,targetCode='rev_it_integration',balancingCode='rev_it_maintenance') {
  const targetSubject=catalog.ICT_SUBJECT_DEFINITIONS.find(s=>s.subjectCode===targetCode),balancingSubject=catalog.ICT_SUBJECT_DEFINITIONS.find(s=>s.subjectCode===balancingCode);
  const read=s=>s.groupId==='revNonItCt'?state.revNonItCt:state[s.groupId][s.key];
  return {side:targetSubject.side,sideLabel:'收入',totalInclAmount:100,fixedOtherInclAmount:0,reallocatablePoolInclAmount:100,
    targetSubject,balancingSubject,targetItem:read(targetSubject),balancingItem:read(balancingSubject),
    targetDisplayName:targetSubject.standardSubjectName,balancingDisplayName:balancingSubject.standardSubjectName,
    beforeTargetInclAmount:read(targetSubject).incl,beforeBalancingInclAmount:read(balancingSubject).incl};
}
async function solve(h,target=0.10325){const s=structureFor(h.state);return h.ctx.solve({ref:s.targetSubject},s,target,{silent:true});}
const evidence={method:'Production candidate + production payload + production solver + production batch setter; controlled metric engine. All 28 subjects and 10-year plans compared.',cases:[]};
async function main(){
  const old=host(true,true);await solve(old);assert.equal(old.spy.mock.callCount(),1);assert.equal(old.state.revIt.integration.incl+old.state.revIt.maintenance.incl,100.01);
  evidence.before={calls:1,total:100.01};
  for(const autoFix of [false,true]) {
    const h=host(autoFix),s=structureFor(h.state),point=h.ctx.build(s,41.33,autoFix);
    assert.equal(point.valid,true);assert.equal(point.targetAmount,autoFix?41.34:41.33);assert.equal(point.balancingAmount,autoFix?58.66:58.67);
    h.state.updateTaxItemsInclBatch([{groupId:'revIt',key:'integration',incl:point.targetAmount,reason:'reverse_calculation_sync'},
      {groupId:'revIt',key:'maintenance',incl:point.balancingAmount,reason:'balance_allocation_sync'}]);
    assert.deepEqual(clean(h.ctx.payload()),clean(point.payload),'actual full payload equals evaluated payload');
    const solver=host(autoFix),result=await solve(solver);
    assert.equal(result.status,autoFix?'error':'success');assert.equal(solver.spy.mock.callCount(),autoFix?0:1);
    assert.equal(Number((solver.state.revIt.integration.incl+solver.state.revIt.maintenance.incl).toFixed(2)),100);
    if(result.status==='success'){assert.deepEqual(clean(solver.ctx.payload()),clean(result.expectedInput));assert.ok(Math.abs(Number(solver.ctx.payload().rev_it_integration.incl_tax)/400-result.target)<=0.0001);}else assert.deepEqual(clean(solver.ctx.payload()),clean(host(autoFix).ctx.payload()));
    evidence.cases.push({name:'original witness',autoFix,calls:solver.spy.mock.callCount(),status:result.status,target:result.target,achieved:result.achieved,
      targetAmount:solver.state.revIt.integration.incl,balancingAmount:solver.state.revIt.maintenance.incl,plans:clean(solver.state.subjectFundingPlans)});
    // Before normalization only 41.33 reaches the goal. Its effective 41.34 no longer does.
    const beforeRounding=initial();beforeRounding.revIt.integration.incl=41.33;beforeRounding.revIt.maintenance.incl=58.67;
    const offTarget=host(autoFix,false,beforeRounding,p=>Number(p.rev_it_integration.incl_tax)===41.33?0.1:0.2);
    const before=clean(offTarget.ctx.payload()),rejected=await solve(offTarget,0.1);
    if(autoFix){assert.equal(rejected.status,'error');assert.equal(offTarget.spy.mock.callCount(),0);assert.deepEqual(clean(offTarget.ctx.payload()),before);}else{assert.equal(rejected.status,'success');assert.equal(offTarget.spy.mock.callCount(),1);}
    evidence.cases.push({name:'normalized metric misses target',autoFix,status:rejected.status,calls:offTarget.spy.mock.callCount()});
    // Final recomputation can fail even after search. Plans and all subject values stay untouched.
    const seen=new Map();const finalFail=host(autoFix,false,initial(),p=>{const amount=Number(p.rev_it_integration.incl_tax);const n=(seen.get(amount)||0)+1;seen.set(amount,n);return amount===40&&n>1?0.5:amount/400;});
    const untouched=clean(finalFail.ctx.payload()),failure=await solve(finalFail,0.1);
    assert.equal(failure.status,'error');assert.equal(finalFail.spy.mock.callCount(),0);assert.deepEqual(clean(finalFail.ctx.payload()),untouched);
    evidence.cases.push({name:'final metric drift',autoFix,status:failure.status,calls:0});
    const fixed=host(autoFix),previous=host(autoFix,true);const nowResult=await solve(fixed,0.1);await solve(previous,0.1);
    assert.equal(nowResult.status,'success');assert.deepEqual(clean(fixed.ctx.payload()),clean(previous.ctx.payload()),'fixed-point success unchanged');
  }
  // Enumerate the real batch transform against its frozen predecessor, including
  // CT linked items, all subject groups, tax normalization, split invalidation and all plan modes.
  let comparisons=0,closedFailures=0;
  for(const autoFix of [false,true]) for(const mode of ['upfront','equal','proportional','custom']) for(const subject of catalog.ICT_SUBJECT_DEFINITIONS) for(const amount of [0,41.33,1060]) {
    const input=initial();
    for(const item of catalog.ICT_SUBJECT_DEFINITIONS){const obj=item.groupId==='revNonItCt'?input.revNonItCt:input[item.groupId][item.key];Object.assign(obj,{incl:40,excl:tax.exclFromIncl(40,obj.tax),splitParts:[{incl:20,excl:tax.exclFromIncl(20,obj.tax)},{incl:20,excl:tax.exclFromIncl(20,obj.tax)}]});const ref={side:item.side,groupId:item.groupId,key:item.key},plan=fund.createDefaultSubjectFundingPlan(ref,40);plan.annualInclValues=[10,10,10,10,...Array(6).fill(0)];plan.annualPercentages=[25,25,25,25,...Array(6).fill(0)];input.subjectFundingPlans[plan.id]=fund.updateSubjectFundingPlanMode(plan,40,mode,4);}
    const now=host(autoFix,false,input),before=host(autoFix,true,input),updates=[{groupId:subject.groupId,key:subject.key,incl:amount,reason:'reverse_calculation_sync'}];
    now.state.updateTaxItemsInclBatch(updates);before.state.updateTaxItemsInclBatch(updates);
    assert.deepEqual(clean(now.ctx.payload()),clean(before.ctx.payload()),subject.subjectCode);comparisons++;
  }
  // Candidate/commit parity for both CT links where linked tax rates differ.
  for(const code of ['rev_ct_line','rev_ct_product']) for(const autoFix of [false,true]) {
    const h=host(autoFix);h.state.revCt[code==='rev_ct_line'?'line':'product'].tax=13;
    const s=structureFor(h.state,code),point=h.ctx.build(s,41.33,autoFix);assert.equal(point.valid,true);
    h.state.updateTaxItemsInclBatch([{groupId:s.targetSubject.groupId,key:s.targetSubject.key,incl:point.targetAmount,reason:'reverse_calculation_sync'},
      {groupId:s.balancingSubject.groupId,key:s.balancingSubject.key,incl:point.balancingAmount,reason:'balance_allocation_sync'}]);
    assert.deepEqual(clean(h.ctx.payload()),clean(point.payload),`CT parity ${code} ${autoFix}`);
  }
  for(let cents=1;cents<10000;cents++) {
    const h=host(true),s=structureFor(h.state);h.state.revIt.maintenance.tax=13;
    const point=h.ctx.build(s,cents/100,true);
    if(!point.valid){closedFailures++;assert.match(point.message,/仍无法保持/);assert.equal(h.spy.mock.callCount(),0);if(closedFailures===3)break;}
  }
  for(const autoFix of [false,true]) {
    const input=initial();input.revIt.integration.tax=100;input.revIt.maintenance.tax=100;
    const h=host(autoFix,false,input),s=structureFor(h.state);s.totalInclAmount=100.01;s.reallocatablePoolInclAmount=100.01;
    const before=clean(h.ctx.payload()),result=await h.ctx.solve({ref:s.targetSubject},s,0.1,{silent:true});
    if(autoFix){assert.equal(result.status,'error');assert.equal(h.spy.mock.callCount(),0);assert.deepEqual(clean(h.ctx.payload()),before);assert.match(result.message,/归一及承接补差/);}
    else{assert.equal(result.status,'success');assert.equal(h.spy.mock.callCount(),1);}
    evidence.cases.push({name:'unclosable pool',autoFix,status:result.status,calls:h.spy.mock.callCount()});
  }
  assert.ok(closedFailures>0,'normalization cannot always close after one correction');
  evidence.batchParity= comparisons;evidence.unclosableCandidates=closedFailures;
  fs.writeFileSync(path.resolve(root,'../../docs/verification/structure-commit-normalization-evidence.json'),JSON.stringify(evidence,null,2)+'\n');
  console.log(`Normalization: original bug reproduced then fixed; candidate/payload/actual commit parity; ${comparisons} frozen desktop batch comparisons; CT linkage; final metric failure spy=0; unclosable candidates rejected.`);
}
if(require.main===module)main().catch(error=>{console.error(error);process.exitCode=1;});
module.exports={host,structureFor,solve,declarations};
