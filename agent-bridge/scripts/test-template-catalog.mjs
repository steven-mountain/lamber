import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {getCatalogCompletion, getCatalogTemplate, templateCatalog} from '../dsh-tool-lamber/lib/catalogCompletion.generated.js';
import {templateTextFieldsByTemplate} from '../dsh-tool-lamber/lib/templateFields.generated.js';
import {readTemplateFields, fillTemplateFields} from '../dsh-tool-lamber/lib/index.js';
const evidence = JSON.parse(readFileSync(new URL('../../docs/verification/template-catalog-transaction-evidence.json',import.meta.url)));
const oldFetch = globalThis.fetch;
process.env.LAMBER_BRIDGE_URL='http://127.0.0.1:1';process.env.LAMBER_BRIDGE_TOKEN='synthetic';
try {
  for (const template of templateCatalog) {
    if (template.excludedReason) { assert.equal(templateTextFieldsByTemplate[template.id],undefined); continue; }
    const {before,after,savedState,approvedFields} = evidence[template.id];
    const textFields = template.fields.filter(f=>f.kind==='text');
    assert.deepEqual(Object.keys(templateTextFieldsByTemplate[template.id]),textFields.map(f=>f.key));
    const execute = async projection => {
      globalThis.fetch = async () => new Response(JSON.stringify(projection));
      return readTemplateFields.execute({templateId:template.name},{agent:{session:{id:'bound'}},signal:new AbortController().signal});
    };
    const a = await execute(before), b = await execute(after);
    const textGroups = new Set(textFields.map(field=>field.completionGroup || field.key));
    const textKeys = new Set(textFields.map(field=>field.key));
    const alreadyFilled = a.completion.filter(item=>textKeys.has(item.key) && item.filled);
    // The frozen desktop rule treats the optional SME field as satisfied even when empty.
    assert.deepEqual(alreadyFilled.map(item=>item.key), template.id==='selection' ? ['gen_zx_is_sme'] : []);
    assert.equal(b.filledCount-a.filledCount,textGroups.size-alreadyFilled.length);
    const textCompletion = b.completion.filter(item=>textKeys.has(item.key));
    assert.equal(textCompletion.length,textGroups.size);
    for (const item of textCompletion) assert.equal(item.filled,true,item.key);
    assert.deepEqual(b.completion.filter(item=>!textKeys.has(item.key)),a.completion.filter(item=>!textKeys.has(item.key)));
    assert.deepEqual(b.completion,getCatalogCompletion(after.templateId,savedState.filledDataJson,after.attachments));
    for(const f of b.fields) assert.equal(f.value,approvedFields[f.key]);
    assert.ok(!('completionState' in b));
    for(const f of template.fields.filter(f=>f.kind!=='text')) {
      await assert.rejects(fillTemplateFields.execute({projectId:after.projectId,templateId:after.templateId,fields:{[f.key]:'invalid'}},{}),/目录|not a declared property/);
    }
    assert.equal(getCatalogTemplate(after.templateId).id,template.id);
    console.log(`PASS ${template.name}: ${textFields.length} texts, read → edited approval → persisted → read/shared completion; nontext rejected`);
  }
  assert.equal(Object.keys(templateTextFieldsByTemplate.demand).length,8);
  for(const name of ['../需求导入表.docx','C:\\立项签批表.docx','立项签批表需求导入表.docx']) assert.equal(getCatalogTemplate(name),undefined);
  const source=readFileSync(new URL('../../src-ui/src/lib/templateCompletion/catalog.ts',import.meta.url),'utf8');
  const generated=readFileSync(new URL('../dsh-tool-lamber/src/catalogCompletion.generated.ts',import.meta.url),'utf8');
  assert.equal(generated.slice(generated.indexOf('export interface')),source.slice(source.indexOf('export interface')));
} finally {globalThis.fetch=oldFetch;}
