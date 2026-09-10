const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),ts=require('typescript'),cp=require('node:child_process');
const root=path.resolve(__dirname,'../src'),load=require('./load_ts.cjs');
const fund=load(path.join(root,'lib/ictSubjectFundingPlan.ts')),tax=load(path.join(root,'lib/taxAmount.ts'));
const clean=x=>JSON.parse(JSON.stringify(x,(k,v)=>['updatedAt','lastChangedAt'].includes(k)?undefined:v));
const compile=source=>ts.transpileModule(source,{compilerOptions:{module:ts.ModuleKind.CommonJS,target:ts.ScriptTarget.ES2020,esModuleInterop:true}}).outputText;
const oldModule={exports:{}};
vm.runInNewContext(compile(cp.execFileSync('git',['show','HEAD:src-ui/src/lib/ictSubjectFundingPlan.ts'],{encoding:'utf8'})),{module:oldModule,exports:oldModule.exports,require:n=>load(path.join(root,'lib',n))});
const ast=ts.createSourceFile('hook.ts',fs.readFileSync(path.join(root,'hooks/useIctState.ts'),'utf8'),ts.ScriptTarget.Latest,true),declarations={};
function visit(n){if(ts.isVariableDeclaration(n)&&['updateTaxItemsInclBatch','collectPositiveFundingSubjects','syncFundingPlansAfterAmountChange'].includes(n.name.getText(ast)))declarations[n.name.getText(ast)]=n.initializer.getText(ast);ts.forEachChild(n,visit);}visit(ast);
assert.equal(Object.keys(declarations).length,1);
const frozen=JSON.parse(fs.readFileSync(path.join(__dirname,'fixtures/structure-commit-before-normalization.json'),'utf8'));
const calcAst=ts.createSourceFile('calc.ts',fs.readFileSync(path.join(root,'hooks/useIctCalculations.ts'),'utf8'),ts.ScriptTarget.Latest,true),calls=[];
function find(n){if(ts.isCallExpression(n)&&n.expression.getText(calcAst)==='state.updateTaxItemsInclBatch')calls.push(n.arguments[0].getText(calcAst));ts.forEachChild(n,find);}find(calcAst);assert.equal(calls.length,2);
function updates(index,context){return vm.runInNewContext(compile('globalThis.result='+calls[index])+'\nresult',context);}
function session(funding){
 const ctx={...load(path.join(root,'lib/ictTaxItemBatch.ts')),...tax,...funding,revIt:{},revCt:{},revNonItCt:{incl:0,tax:9,excl:0},costIt:{},costCt:{},costMix:{},plans:{},ignoredDataHash:null,isTaxInclAutoFixEnabled:()=>false,setCashflowCalculationSourceState:()=>{}};
 for(const group of ['revIt','revCt','revNonItCt','costIt','costCt','costMix'])ctx['set'+group[0].toUpperCase()+group.slice(1)]=v=>ctx[group]=v;
 ctx.setSubjectFundingPlansState=fn=>ctx.plans=fn(ctx.plans);
 for(const [key,amount] of [['integration',106000],['bidding',1060]]){
  ctx.costIt[key]={incl:amount,excl:tax.exclFromIncl(amount,6),tax:6};
  const ref={side:'cost',groupId:'costIt',key},plan=fund.createDefaultSubjectFundingPlan(ref,amount);
  ctx.plans[plan.id]={...plan,mode:'custom',annualInclValues:[amount*.25,amount*.25,amount*.25,amount*.25,...Array(6).fill(0)]};
 }
 vm.runInNewContext(compile((funding===fund ? Object.entries(declarations).map(([k,v])=>`const ${k}=${v};`).join('\n') : frozen.batchHelpers+frozen.batch)+'\nglobalThis.batch=updateTaxItemsInclBatch;'),ctx);
 return ctx;
}
const feeContext={targetSubject:{groupId:'costIt',key:'integration'},serviceFeeSubject:{groupId:'costIt',key:'bidding'},selectionFeeMergeService:false,writeAmounts:{targetIncl:300000,serviceFeeIncl:2429.95}};
const structureContext={structure:{targetSubject:feeContext.targetSubject,balancingSubject:feeContext.serviceFeeSubject},finalPoint:{targetAmount:300000,balancingAmount:2429.95}};
for(const [name,index,context] of [['selection limit',0,feeContext],['structure reverse',1,structureContext]]){
 const old=session(oldModule.exports),current=session(fund),positive=updates(index,context);
 old.batch(positive);current.batch(positive);
 assert.deepEqual(clean(current.costIt),clean(old.costIt));assert.deepEqual(clean(current.plans),clean(old.plans));
 const zeroContext=index===0?{...context,writeAmounts:{targetIncl:0,serviceFeeIncl:2429.95}}:{...context,finalPoint:{targetAmount:0,balancingAmount:2429.95}};
 current.batch(updates(index,zeroContext));const id='cost:costIt:integration';
 assert.equal(current.plans[id].enabled,false);assert.equal(current.plans[id].annualInclValues.reduce((a,b)=>a+b,0),0);
 current.plans=fund.normalizeSubjectFundingPlans(JSON.parse(JSON.stringify(current.plans)));
 current.batch(positive);
 assert.equal(current.plans[id].lastChangeReason,'restored_after_zero');
 assert.deepEqual(clean(current.plans[id].annualInclValues),[75000,75000,75000,75000,...Array(6).fill(0)]);
 console.log(`${name}: production call arguments → production batch setter; positive result equals pre-fix, zero/reload/restore preserves four-year plan`);
}
