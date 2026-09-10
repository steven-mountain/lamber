import assert from 'node:assert/strict';
import {readFileSync,writeFileSync} from 'node:fs';
import {createRequire} from 'node:module';
import vm from 'node:vm';
import {readTemplateFields} from '../dsh-tool-lamber/lib/index.js';
const require=createRequire(new URL('../../src-ui/package.json',import.meta.url));
const ts=require('typescript');
const frozen=JSON.parse(readFileSync(new URL('../../src-ui/scripts/fixtures/selection-page-before-catalog.json',import.meta.url)));
const catalogSource=readFileSync(new URL('../../src-ui/src/lib/templateCompletion/catalog.ts',import.meta.url),'utf8');
function api(catalog){const mod={exports:{}};vm.runInNewContext(ts.transpileModule(catalogSource,{compilerOptions:{module:ts.ModuleKind.CommonJS,esModuleInterop:true}}).outputText,{module:mod,exports:mod.exports,require:()=>catalog});return mod.exports;}
const current=api(JSON.parse(readFileSync(new URL('../../src-ui/src/lib/templateCompletion/catalog.json',import.meta.url))));
const old=api(frozen.catalog),clean=x=>JSON.parse(JSON.stringify(x));
const rows=JSON.parse(readFileSync(new URL('../../docs/verification/selection-page-projection-evidence.json',import.meta.url)));
process.env.LAMBER_BRIDGE_URL='http://127.0.0.1:1';process.env.LAMBER_BRIDGE_TOKEN='synthetic';
const previous=globalThis.fetch,results=[];
try{for(const row of rows){
 globalThis.fetch=async(_url,options)=>{assert.deepEqual(JSON.parse(options.body),{templateId:row.name,sessionId:'trusted'});return new Response(JSON.stringify(row.projection));};
 const output=await readTemplateFields.execute({templateId:row.name},{agent:{session:{id:'trusted'}},signal:new AbortController().signal});
 const ui=current.getCatalogCompletion(row.name,row.state,row.projection.attachments);
 assert.deepEqual(output.completion,clean(ui),'Rust saved projection → actual plugin equals UI for same known facts');
 if(row.name!=='甄选结果签批表.docx')assert.deepEqual(output.completion,clean(old.getCatalogCompletion(row.name,row.state,row.projection.attachments)),'other templates complete list/order frozen baseline');
 else {
  assert.equal(output.totalCount,17);assert.equal(output.unknownCount,6);
  assert.equal(output.completion.find(r=>r.key==='selection_batch_name').filled,row.mode==='single'||row.filled);
  assert.ok(output.missingFields.every(r=>r.evaluated));
  assert.ok(!output.fields.some(f=>f.key==='selection_batch_name'));
 }
 results.push({template:row.name,mode:row.mode,filled:row.filled,completion:output.completion,totalCount:output.totalCount,filledCount:output.filledCount,unknownCount:output.unknownCount});
}}finally{globalThis.fetch=previous;}
writeFileSync(new URL('../../docs/verification/selection-page-read-tool-evidence.json',import.meta.url),JSON.stringify(results,null,2)+'\n');
console.log('PASS: 16 real Rust projections → actual read tool → UI equality; frozen three-template baseline; selection 17 entries / 6 unknown; read-only batch name.');
