const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),assert=require('node:assert/strict'),cp=require('node:child_process'),ts=require('typescript');
const root=path.resolve(__dirname,'../src'),repo=path.resolve(root,'../..');
const cache=new Map();
function load(file){file=path.resolve(root,file);if(file.endsWith('.json'))return JSON.parse(fs.readFileSync(file,'utf8'));if(cache.has(file))return cache.get(file);const m={exports:{}};const code=ts.transpileModule(fs.readFileSync(file,'utf8'),{compilerOptions:{module:ts.ModuleKind.CommonJS,target:ts.ScriptTarget.ES2020,esModuleInterop:true}}).outputText;vm.runInNewContext(code,{module:m,exports:m.exports,console,require:n=>n.startsWith('.')?load(path.resolve(path.dirname(file),n)+(path.extname(n)?'':'.ts')):require(n)});cache.set(file,m.exports);return m.exports;}
const plain=x=>JSON.parse(JSON.stringify(x)),clean=x=>JSON.parse(JSON.stringify(x,(key,v)=>['lastChangedAt','updatedAt'].includes(key)?undefined:v));
const catalog=load('lib/ictSubjectCatalog.ts'),fund=load('lib/ictSubjectFundingPlan.ts'),tax=load('lib/taxAmount.ts'),edits=load('lib/ictTaxItemEdit.ts'),calc=load('lib/ictCalculationInput.ts'),{prepareBenefitSimulation}=load('lib/ictBenefitSimulation.ts');
if(process.argv.includes('--prepare')) {const job=JSON.parse(fs.readFileSync(0,'utf8'));try{process.stdout.write(JSON.stringify(prepareBenefitSimulation(job.input,job.overrides,false)));}catch(e){process.stderr.write(e.message);process.exitCode=1;}return;}
const baselinePath=path.join(__dirname,'fixtures/ict-desktop-edit-baseline.json');
if(process.argv.includes('--capture-baseline')){
 const original=cp.execFileSync('git',['show','HEAD:src-ui/src/hooks/useIctState.ts'],{cwd:repo,encoding:'utf8'}),ast=ts.createSourceFile('baseline.ts',original,ts.ScriptTarget.Latest,true);const names=['updateTaxItem','collectPositiveFundingSubjects','syncFundingPlansAfterAmountChange'];const source={};function visit(n){if(ts.isVariableDeclaration(n)&&names.includes(n.name.getText(ast)))source[n.name.getText(ast)]=n.initializer.getText(ast);ts.forEachChild(n,visit);}visit(ast);assert.equal(Object.keys(source).length,3);fs.writeFileSync(baselinePath,JSON.stringify({commit:cp.execFileSync('git',['rev-parse','HEAD'],{cwd:repo,encoding:'utf8'}).trim(),source},null,2)+'\n');
}
const original=JSON.parse(fs.readFileSync(baselinePath,'utf8'));
function desktopEdit(state,plans,s,field,val,autoFix){
 const ctx={...state,...tax,...fund,plans,ignoredDataHash:null,isTaxInclAutoFixEnabled:()=>autoFix,setCashflowCalculationSourceState:()=>{},setSubjectFundingPlansState:fn=>{ctx.plans=fn(ctx.plans);}};
 for(const group of ['revIt','revCt','revNonItCt','costIt','costCt','costMix'])ctx['set'+group[0].toUpperCase()+group.slice(1)]=v=>ctx[group]=v;
 const source=Object.entries(original.source).map(([k,v])=>`const ${k} = ${v};`).join('\n')+'\nglobalThis.edit=updateTaxItem;';
 vm.runInNewContext(ts.transpileModule(source,{compilerOptions:{target:ts.ScriptTarget.ES2020}}).outputText,ctx);
 ctx.edit(s.groupId,s.key,field,val,'manual_amount_sync',{normalizeIncl:true});
 return {state:Object.fromEntries(['revIt','revCt','revNonItCt','costIt','costCt','costMix'].map(k=>[k,ctx[k]])),plans:ctx.plans};
}
function fixture(){
 const input={project_name:'AI测算28科目对照',customer_name:'合成验收客户',property_rights:'self',discount_rate:'0.055',project_years:4,cashflow_model:'model_e',cashflow_segment_value_mode:'amount',cashflow_calculation_source:'subject_funding_plans',rev_distribution:[.25,.25,.25,.25,0,0,0,0,0,0],cost_distribution:[.25,.25,.25,.25,0,0,0,0,0,0],cashflow_segments:[],ignore_tail_difference:false};
 const state={revIt:{},revCt:{},revNonItCt:{},costIt:{},costCt:{},costMix:{}};
 let plans={};
 catalog.ICT_SUBJECT_DEFINITIONS.forEach((s,i)=>{const incl=s.subjectCode==='rev_it_integration'?106000:s.subjectCode==='cost_it_integration'?53000:(i+1)*106;const item={incl,tax:s.defaultTaxRate,excl:tax.exclFromIncl(incl,s.defaultTaxRate),customSubjectName:`业务名称${i+1}`,billingSubjectName:`开票名称${i+1}`};if(s.groupId==='revNonItCt')state.revNonItCt=item;else state[s.groupId][s.key]=item;input[s.subjectCode]=edits.serializeTaxItemForPayload(item);});
 plans=fund.initializeMissingSubjectFundingPlans(plans,edits.positiveFundingSubjects(state));
 for(const plan of Object.values(plans)){const total=plan.annualInclValues.reduce((a,b)=>a+b,0);plan.mode='custom';plan.annualInclValues=[total*.25,total*.25,total*.25,total*.25,0,0,0,0,0,0];}
 return {state,plans,input:calc.finalizeIctInputWithFundingPlans(input,plans).input};
}
const base=fixture();let comparisons=0;const evidence=[];
for(const autoFix of [false,true])for(const s of catalog.ICT_SUBJECT_DEFINITIONS)for(const [field,val] of [['incl',1038],['tax',9],['excl',979.25]]){
 const old=desktopEdit(plain(base.state),plain(base.plans),s,field,val,autoFix),actual=edits.editIctTaxItem(plain(base.state),plain(base.plans),s.groupId,s.key,field,val,autoFix,'manual_amount_sync',{normalizeIncl:true});
 assert.deepEqual(clean(actual.state),clean(old.state),`${s.subjectCode}/${field}/${autoFix} subjects`);assert.deepEqual(clean(actual.plans),clean(old.plans),`${s.subjectCode}/${field}/${autoFix} plans`);comparisons++;
}
for(const overrides of [[],[{subject:'rev_it_integration',inclTax:'800000'}],[{subject:'rev_ct_product',inclTax:'80000'}],[{subject:'cost_it_device',inclTax:'30000',taxRate:'6'}]]){
 const before=JSON.stringify(base.input),result=prepareBenefitSimulation(base.input,overrides,false);assert.equal(JSON.stringify(base.input),before,'saved input unchanged');
 let desk={state:plain(base.state),plans:plain(base.plans)};
 for(const o of overrides){const s=catalog.ICT_SUBJECT_DEFINITIONS.find(s=>s.subjectCode===o.subject);desk=desktopEdit(desk.state,desk.plans,s,'incl',Number(o.inclTax),false);if(o.taxRate!==undefined)desk=desktopEdit(desk.state,desk.plans,s,'tax',Number(o.taxRate),false);}
 const desktop=plain(base.input);for(const s of catalog.ICT_SUBJECT_DEFINITIONS)desktop[s.subjectCode]=edits.serializeTaxItemForPayload(s.groupId==='revNonItCt'?desk.state.revNonItCt:desk.state[s.groupId][s.key]);const finalized=calc.finalizeIctInputWithFundingPlans(desktop,desk.plans).input;
 for(const s of catalog.ICT_SUBJECT_DEFINITIONS)assert.deepEqual(plain(result.input[s.subjectCode]),plain(finalized[s.subjectCode]),s.subjectCode);
 for(const key of ['rev_cashflow_excl','cost_cashflow_excl','it_rev_cashflow_excl','it_cost_cashflow_excl'])assert.deepEqual(plain(result.input[key]),plain(finalized[key]),key);
 if(overrides[0]?.subject==='rev_it_integration')assert.notDeepEqual(result.input.rev_cashflow_excl,base.input.rev_cashflow_excl,'Model E revenue cashflow responds');
 if(overrides[0]?.subject==='rev_ct_product')assert.equal(result.input.cost_ct_other.incl_tax,'80000');
 evidence.push({overrides,saved:base.input,prepared:result,desktop:finalized});
}
for(const overrides of [[{subject:'bad',inclTax:'1'}],[{subject:'rev_it_integration',inclTax:'十万'}],[{subject:'rev_it_integration',inclTax:'1.001'}],[{subject:'rev_ct_product',inclTax:'1'},{subject:'cost_ct_other',inclTax:'1'}],[{subject:'rev_it_integration',inclTax:'1'},{subject:'rev_it_integration',inclTax:'2'}]])assert.throws(()=>prepareBenefitSimulation(base.input,overrides,false));
for(const altered of [{...base.input,cashflow_calculation_source:null},{...base.input,rev_cashflow_excl:['1',...Array(9).fill('0')]},{...base.input,revenue_balance_rule:{enabled:true}},{...base.input,subject_funding_plans:{}}]){assert.throws(()=>prepareBenefitSimulation(altered,[{subject:'rev_it_integration',inclTax:'800000'}],false));assert.deepEqual(plain(prepareBenefitSimulation(altered,[],false).input),plain(altered));}
const split=plain(base.input);split.rev_it_integration={...split.rev_it_integration,split_parts:[{incl_tax:'53000',excl_tax:'50000'},{incl_tax:'53000',excl_tax:'50000'}]};const splitResult=prepareBenefitSimulation(split,[{subject:'rev_it_integration',inclTax:'800000'}],false);assert.equal(splitResult.input.rev_it_integration.split_parts,undefined);assert(splitResult.linkedChanges.some(c=>c.kind==='split_cleared'));assert.equal(split.rev_it_integration.split_parts.length,2);
const invalid=plain(split);invalid.rev_it_integration.split_parts[0].excl_tax='50001';assert.throws(()=>prepareBenefitSimulation(invalid,[{subject:'cost_it_device',inclTax:'100'}],false));
const evidencePath=path.join(repo,'src-tauri/src/agent_bridge/fixtures/benefit-simulation-desktop.json');
const generated={baselineCommit:original.commit,comparisons,cases:evidence};
if(process.argv.includes('--write-fixtures')) fs.writeFileSync(evidencePath,JSON.stringify(generated,null,2)+'\n');
else assert.deepEqual(clean(generated),clean(JSON.parse(fs.readFileSync(evidencePath,'utf8'))),'committed desktop fixture remains equivalent');
console.log(`${comparisons} old-desktop/new-shared edit comparisons passed; all 28 subjects + 4 annual arrays, CT, tax, split, Model E, invalid inputs, legacy refusal and immutable saved snapshots passed.`);
