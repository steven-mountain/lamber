const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),crypto=require('node:crypto'),ts=require('typescript');
const load=require('./load_ts.cjs'),root=path.resolve(__dirname,'../..');
const frozen=require('./fixtures/selection-page-before-catalog.json');
const catalogPath=path.join(root,'src-ui/src/lib/templateCompletion/catalog.json');
const current=JSON.parse(fs.readFileSync(catalogPath));
const clean=x=>JSON.parse(JSON.stringify(x));
function api(catalog){const mod={exports:{}};vm.runInNewContext(ts.transpileModule(fs.readFileSync(path.join(root,'src-ui/src/lib/templateCompletion/catalog.ts'),'utf8'),{compilerOptions:{module:ts.ModuleKind.CommonJS,esModuleInterop:true}}).outputText,{module:mod,exports:mod.exports,require:()=>catalog});return mod.exports;}
const oldApi=api(frozen.catalog),newApi=api(current);
const source=fs.readFileSync(path.join(root,'src-ui/src/views/TemplateForms.tsx'),'utf8');
const block=source.slice(source.indexOf('  const selectionResultCompletionItems ='),source.indexOf('  const meetingCompletion ='));
const oldBlock=ts.transpileModule(frozen.selectionBlock,{compilerOptions:{target:ts.ScriptTarget.ES2020}}).outputText;
const defaults=load(path.join(root,'src-ui/src/lib/templateGenerationState.ts')).TEMPLATE_FIELD_DEFAULTS;
function evaluate(block,state){return clean(vm.runInNewContext(block+'\nselectionResultCompletionItems',{...state,getCatalogCompletion:newApi.getCatalogCompletion,hasText:v=>String(v??'').trim().length>0,getFormValue:(key,fallback='')=>state.formData[key]??defaults[key]??fallback}));}
const oldLabels=[];let count=0;
for(const mode of ['single','batch'])for(let mask=0;mask<1024;mask++){
 const flag=n=>Boolean(mask&(1<<n));
 const state={selectionResultMode:mode,selectionBatchName:flag(0)?'合成批次':'',selectionBatchProjectIds:flag(1)?['a','b']:['a'],selectionBatchProjects:flag(2)?[{},{}]:[{}],currentSchemeStage:flag(3)?'post_selection':'pre_selection',projectBackground:flag(4)?'项目背景':'',formData:flag(5)?{gen_zx_winner_name:'合作伙伴',gen_zx_scope:'',gen_zx_method:'',gen_zx_rule:'',gen_zx_is_sme:''}:{},revCollection:flag(6)?'收款':'',expPayment:flag(6)?'付款':'',selectionBlockingConflicts:flag(7)?[{}]:[],selectionOverrideConflicts:flag(8)?[{}]:[],selectionConflictAcknowledged:flag(9),selectionRenewalDecisionsComplete:flag(9),selectionApprovalAmountPreview:flag(4)?{lt:()=>flag(3)}:null};
 const old=evaluate(oldBlock,state),next=evaluate(block,state);assert.equal(next.length,17);assert.ok(next.every(r=>r.evaluated));
 const original=next.filter(r=>old.some(o=>o.label===r.label));
 assert.deepEqual(original.map(({label,filled})=>({label,filled})),old,'all original 14 values AND order');
 if(!oldLabels.length)oldLabels.push(...old.map(r=>r.label));
 assert.equal(next.find(r=>r.key==='selection_batch_name').filled,mode==='single'||flag(0));count++;
}
// Other template records and their page calls are byte-for-byte unchanged.
for(const template of frozen.catalog.templates.filter(t=>t.id!=='selection'))assert.deepEqual(current.templates.find(t=>t.id===template.id),template);
assert.equal(source.slice(source.indexOf('  const meetingCompletionItems ='),source.indexOf('  const selectionResultCompletionItems =')),frozen.otherPageBlock);
const states=[{}, {formData:{},projectScale:'small'}, {formData:Object.fromEntries(current.templates.flatMap(t=>t.fields).filter(f=>f.kind==='text').map(f=>[f.key,'已填'])),projectScale:'large',hasMidThree:true,midThreeCode:'a',midThreeName:'b',projectBackground:'背景',itContent:'it',ctContent:'ct',revCollection:'收款',expPayment:'付款',completionValues:{gen_proj_bg:true},techItems:[{serviceName:'服务',amount:1}],inqVendors:[{vendorName:'供应商',quote:1}],hasSingleSource:false,hasSecurity:false,hasPublicUrl:false}];
for(const name of ['需求导入表.docx','立项签批表.docx','会审纪要.docx'])for(const state of states)assert.deepEqual(clean(newApi.getCatalogCompletion(name,state)),clean(oldApi.getCatalogCompletion(name,state)));
const approval=newApi.getCatalogCompletion('立项签批表.docx',{formData:{gen_sign_it_content:'IT',gen_sign_ct_content:'CT'},revCollection:'收',expPayment:'付',completionValues:{gen_proj_bg:true}});assert.equal(approval.filter(r=>r.filled).length,7);
const textState={selectionResultMode:'single',selectionBatchName:'',formData:{gen_zx_content_desc:'',gen_zx_industry:'',gen_zx_std_plan:''}};
for(const key of ['gen_zx_content_desc','gen_zx_industry','gen_zx_std_plan']){
 const before=newApi.getCatalogCompletion('甄选结果签批表.docx',textState);const after=newApi.getCatalogCompletion('甄选结果签批表.docx',{...textState,formData:{...textState.formData,[key]:'已填写'}});
 assert.equal(after.filter(r=>r.filled).length-before.filter(r=>r.filled).length,1);assert.equal(before.find(r=>r.key===key).filled,false);
}
const unknown=newApi.getCatalogCompletion('甄选结果签批表.docx',textState).filter(r=>!r.evaluated).map(r=>r.key);assert.deepEqual(clean(unknown),['post_selection_scheme','gen_proj_bg','public_fields_consistent','batch_overrides_acknowledged','renewal_costs_confirmed','approval_amount_below_500k']);
for(const [file,hash]of Object.entries(frozen.hashes))assert.equal(crypto.createHash('sha256').update(fs.readFileSync(path.join(root,file))).digest('hex'),hash,file);
const result={oldCasesCompared:count,oldItemsPerCase:14,newItems:17,oldLabels,newOrder:current.templates.find(t=>t.id==='selection').fields.map(f=>f.label),sixChatUnknown:unknown,approvalFilled:7,protectedHashes:frozen.hashes};
fs.writeFileSync(path.join(root,'docs/verification/selection-page-catalog-evidence.json'),JSON.stringify(result,null,2)+'\n');console.log('PASS:',count,'single/batch cases × 14 original predicates and order; 3 new fields; other templates unchanged; 6 unknown; approval 7/7; protected hashes unchanged.');
