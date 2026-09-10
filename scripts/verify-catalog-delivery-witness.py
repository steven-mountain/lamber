"""Temporarily register the independent fifth catalog fixture; restore data and plugin in finally.
Run from the repository root. Uses synthetic test workspaces only; do not edit catalog concurrently.
"""
from pathlib import Path
import json,hashlib,subprocess,tempfile
catalog=Path('src-ui/src/lib/templateCompletion/catalog.json');original=catalog.read_bytes()
evidence=Path('docs/verification/template-catalog-transaction-evidence.json');old_evidence=evidence.read_bytes()
roots=['src-ui/src','src-tauri/src','agent-bridge/dsh-tool-lamber/src','agent-bridge/scripts']
def hashes():return {str(p):hashlib.sha256(p.read_bytes()).hexdigest() for root in roots for p in Path(root).rglob('*') if p.suffix in ['.ts','.tsx','.rs','.mjs']}
before=hashes()
try:
 data=json.loads(original);data['templates'].append(json.loads(Path('src-ui/scripts/fixtures/catalog-delivery-witness.json').read_text()));catalog.write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n')
 with tempfile.TemporaryFile(mode='w+') as log:
  for command in [['npm','run','build','--prefix','agent-bridge/dsh-tool-lamber'],['cargo','test','--manifest-path','src-tauri/Cargo.toml','template_catalog_all'],['node','agent-bridge/scripts/test-template-catalog.mjs']]:
   subprocess.run(command,stdout=log,stderr=subprocess.STDOUT,check=True)
 after=hashes();changed=[p for p in before if before[p]!=after.get(p)]
 assert not changed,changed
 record=json.loads(evidence.read_text())['delivery_witness']
 Path('docs/verification/catalog-delivery-witness-evidence.json').write_text(json.dumps(dict(sourceCount=len(before),changedSources=changed,fixture='src-ui/scripts/fixtures/catalog-delivery-witness.json',transaction=record,scope='Synthetic independent fifth template; not a new production business template'),ensure_ascii=False,indent=2)+'\n')
 print('Fifth independent data-only witness passed:',len(before),'unchanged source hashes')
finally:
 catalog.write_bytes(original);evidence.write_bytes(old_evidence)
 subprocess.run(['npm','run','build','--prefix','agent-bridge/dsh-tool-lamber'],stdout=subprocess.DEVNULL,check=True)
