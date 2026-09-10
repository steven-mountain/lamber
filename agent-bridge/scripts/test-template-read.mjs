import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {readTemplateFields} from '../dsh-tool-lamber/lib/index.js';
const evidence=JSON.parse(readFileSync(new URL('../../docs/verification/template-read-projection-evidence.json',import.meta.url),'utf8'));
const oldFetch=globalThis.fetch;
process.env.LAMBER_BRIDGE_URL='http://127.0.0.1:1';process.env.LAMBER_BRIDGE_TOKEN='synthetic';
try {
  for(const projection of Object.values(evidence)){
    globalThis.fetch=async (_url,options)=>{
      assert.deepEqual(JSON.parse(options.body),{templateId:'需求导入表',sessionId:'trusted'});
      return new Response(JSON.stringify(projection));
    };
    const result=await readTemplateFields.execute({templateId:'需求导入表'}, {agent:{session:{id:'trusted'}},signal:new AbortController().signal});
    assert.equal(result.totalCount,11);
    assert.deepEqual(result.missingFields.map(f=>f.key),['gen_demand_env_require','attach1','attach2']);
    assert.equal('completionState' in result,false);
    assert.equal(result.fields[2].value,projection.fields[2].value);
    assert.equal(result.truncated,projection.truncated);
    assert.equal(result.filledCount,8);
  }
  const source=readFileSync(new URL('../../src-ui/src/lib/templateCompletion/catalog.ts',import.meta.url),'utf8');
  const generated=readFileSync(new URL('../dsh-tool-lamber/src/catalogCompletion.generated.ts',import.meta.url),'utf8');
  assert.equal(generated.slice(generated.indexOf('export interface')), source.slice(source.indexOf('export interface')), 'plugin compiles the actual shared pure function');
  console.log('PASS: production Rust projection → production plugin → shared 11-rule completion, full prose, explicit truncation, no approval');
}finally{globalThis.fetch=oldFetch;}
