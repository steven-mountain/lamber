// Run `snapshot` before adding JSON data, then `verify` after all extension checks.
import {readFileSync,writeFileSync,readdirSync} from 'node:fs';
import {createHash} from 'node:crypto';
import {resolve,relative} from 'node:path';
const root = resolve(import.meta.dirname,'..');
const output=resolve(root,'docs/verification/template-catalog-extension-evidence.json');
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
function files(dir) {
  return readdirSync(dir,{withFileTypes:true}).flatMap(e => e.isDirectory() ? files(resolve(dir,e.name)) : [resolve(dir,e.name)]);
}
function snapshot() {
  const paths = ['src-ui/src','src-tauri/src','agent-bridge/dsh-tool-lamber/src','agent-bridge/scripts','src-ui/scripts','scripts']
    .flatMap(p=>files(resolve(root,p))).filter(p=>/\.(?:tsx?|rs|mjs|cjs)$/.test(p));
  paths.push(resolve(root,'src-tauri/build.rs'));
  const sources = Object.fromEntries(paths.sort().map(p=>[relative(root,p),sha(readFileSync(p))]));
  const catalog=JSON.parse(readFileSync(resolve(root,'src-ui/src/lib/templateCompletion/catalog.json')));
  return {sources,catalog,templates:catalog.templates.filter(t=>!t.excludedReason).map(t=>t.id)};
}
if(process.argv[2]==='snapshot') {
  const before=snapshot();
  if(before.templates.length!==2) throw new Error('Baseline must have exactly two supported templates');
  writeFileSync(output,JSON.stringify({before},null,2)+'\n');
  console.log(`Frozen ${Object.keys(before.sources).length} source hashes, including generated TypeScript; demand + approval only`);
} else if(process.argv[2]==='verify') {
  const record=JSON.parse(readFileSync(output));
  const after=snapshot();
  const changed=[...new Set([...Object.keys(record.before.sources),...Object.keys(after.sources)])].filter(k=>record.before.sources[k]!==after.sources[k]);
  if(changed.length) throw new Error(`Source changed during data-only extension: ${changed.join(', ')}`);
  if(after.templates.length!==3) throw new Error('Extension must have exactly three supported templates');
  record.after={templates:after.templates,catalog:after.catalog};
  record.unchangedSourceCount=Object.keys(after.sources).length;
  record.changedSourceFiles=changed;
  record.result='PASS: added third template using catalog JSON only; all .ts/.tsx/.rs and test/build scripts unchanged, including generated TypeScript';
  writeFileSync(output,JSON.stringify(record,null,2)+'\n');
  console.log(record.result);
} else throw new Error('Use snapshot or verify');
