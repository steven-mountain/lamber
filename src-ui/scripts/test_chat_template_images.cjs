const assert = require('node:assert/strict');
const path = require('node:path');
const fs = require('node:fs'), vm = require('node:vm'), ts = require('typescript');
function load(file, mocks = {}) {
  if (file.endsWith('.json')) return JSON.parse(fs.readFileSync(file, 'utf8'));
  const mod = {exports:{}};
  vm.runInNewContext(ts.transpileModule(fs.readFileSync(file,'utf8'), {compilerOptions:{module:ts.ModuleKind.CommonJS,esModuleInterop:true,target:ts.ScriptTarget.ES2020}}).outputText,
    {module:mod,exports:mod.exports,require:name => mocks[name] || load(path.resolve(path.dirname(file), name)+(path.extname(name)?'':'.ts'),mocks)});
  return mod.exports;
}
const root = path.resolve(__dirname, '../src');
const { wantsTemplateImages, templateImagePrompt } = load(path.join(root, 'ai/templateImageIntent.ts'));
for (const text of ['展示已有图片', '请把附件1给我看', '修改图片', '替换客户确认材料', '查看需求导入表的图片', '打开asset_45c0b489687dc85f']) assert.equal(wantsTemplateImages(text), true, text);
for (const text of ['你好', '这张图是什么', '不要修改图片', '如何替换图片', '昨天让我展示图片', '> 展示已有图片', '```\n修改图片\n```', '客户说“修改图片”']) assert.equal(wantsTemplateImages(text), false, text);
assert.match(templateImagePrompt(true), /No write occurs/);

async function main() {
  let binding = { workspaceId: 'w', projectId: 'p', projectName: '测试项目' }, workspaceId = 'w';
  const calls = []; let duringRead; let failNotify = false;
  const target = {sessionId:'s',workspaceId:'w',projectId:'p',projectName:'测试项目',templateName:'ICT项目需求导入表.docx',assetId:'old',usage:'attach1',name:'old.png'};
  const api = load(path.join(root, 'services/chatTemplateImages.ts'), {
    '@tauri-apps/api/core': { invoke: async (command,args) => {
      if(command==='ai_get_session_binding') return binding;
      calls.push([command,args]); return 'new';
    }},
    './domainSaveService': {domainSaveService:{loadTemplateAssets: async () => [
      {id:'old',templateName:target.templateName,assetType:'image',usage:'attach1',originalFileName:'old.png'},
      {id:'sibling',templateName:target.templateName,assetType:'image',usage:'attach1'},
      {id:'vendor',templateName:'会审纪要.docx',assetType:'image',usage:'vendor_0'},
    ]}},
    '../utils/workspaceService': {workspaceService:{getState:async()=>({currentWorkspace:{workspaceId}})}},
    './aiProjectContextService': {loadAiTemplateAsset:async()=>{if(duringRead) duringRead();return {dataUrl:'data:image/png;base64,test'};}},
    './demandTemplateAssets': {publishDemandAssetsChanged:async()=>{if(failNotify)throw Error('notify');}},
  });
  const list = await api.listChatTemplateImages('s');
  for (const file of [{type:'image/gif',size:10},{type:'image/png',size:0},{type:'image/png',size:21*1024*1024}]) {
    await assert.rejects(api.prepareReplacement(file)); assert.equal(calls.length,0);
  }
  assert.equal(list.length,2,'existing images before template first save are listed');
  assert.equal(calls.length,0,'view/list never writes');
  await api.readChatTemplateImage(target);assert.equal(calls.length,0,'preview never writes');
  duringRead=()=>{workspaceId='other';};
  await assert.rejects(api.readChatTemplateImage(target),/切换/);workspaceId='w';duringRead=null;
  const next={name:'new.png',dataUrl:'data:image/png;base64,test',width:10,height:20};
  for(const change of [()=>{binding={...binding,projectId:'other'};},()=>{workspaceId='other';}]){
    change();await assert.rejects(api.replaceChatTemplateImage(target,next,()=>true),/切换/);
    assert.equal(calls.length,0);binding={workspaceId:'w',projectId:'p'};workspaceId='w';
  }
  await assert.rejects(api.replaceChatTemplateImage(target,next,()=>false),/失效/);assert.equal(calls.length,0);
  failNotify=true;
  const result=await api.replaceChatTemplateImage(target,next,()=>true);
  assert.equal(calls.length,1,'exactly one atomic replacement, no append/delete sequence');
  assert.equal(calls[0][0],'ai_replace_template_image');assert.equal(calls[0][1].request.assetId,'old');
  assert.equal(result.assetId,'new');assert.match(result.refreshWarning,/已保存/);
  console.log('PASS: explicit image intent, existing assets, view zero writes, stale reads rejected, binding/active failures zero writes, single replacement IPC and committed-notification failure.');
}
main().catch(e=>{console.error(e);process.exitCode=1;});
