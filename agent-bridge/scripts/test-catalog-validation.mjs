import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { validateCatalog } from './validate-template-catalog.mjs';
const catalog=JSON.parse(readFileSync(new URL('../../src-ui/src/lib/templateCompletion/catalog.json',import.meta.url)));
validateCatalog(catalog);
const witness=JSON.parse(readFileSync(new URL('../../src-ui/scripts/fixtures/catalog-delivery-witness.json',import.meta.url)));
validateCatalog({templates:[witness]});
for(const patch of [{kind:'list.generated'},{requiredWhen:{field:'projectScale',equals:{nested:true}}},{requiredWhen:{field:'x',equals:true,extra:true}},{listType:'unknown'},{completionSources:[{field:'__proto__'}]},{validRow:{nonEmpty:['secret'],positive:[]}}]){
 const invalid=structuredClone(witness);Object.assign(invalid.fields[0],patch);assert.throws(()=>validateCatalog({templates:[invalid]}));
}
console.log('Catalog validation: independent delivery fixture accepted; unknown kind, invalid condition/source/group/list metadata rejected.');
