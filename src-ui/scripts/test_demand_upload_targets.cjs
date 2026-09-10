const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const root = path.join(__dirname, '../src');
function load(relative, mocks = {}) {
  const filename = path.join(root, relative);
  if (filename.endsWith('.json')) return JSON.parse(fs.readFileSync(filename, 'utf8'));
  const mod = { exports: {} };
  const code = ts.transpileModule(fs.readFileSync(filename, 'utf8'), { compilerOptions: {
    module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, esModuleInterop: true,
  }}).outputText;
  vm.runInNewContext(code, { exports: mod.exports, module: mod, require: name => {
    if (mocks[name]) return mocks[name];
    const next = path.relative(root, path.resolve(path.dirname(filename), name));
    return load(path.extname(next) ? next : `${next}.ts`, mocks);
  }});
  return mod.exports;
}
const template = 'ICT项目需求导入表.docx';
const binding = { workspaceId: 'workspace-a', projectId: 'project-a', projectName: '合成项目A' };
function fixture(options = {}) {
  const f = { binding, workspaceId: binding.workspaceId, reads: [], assets: [], saved: [{ templateId: template, filledDataJson: { hasPublicUrl: true } }], ...options };
  const api = load('services/demandUploadTargets.ts', {
    '@tauri-apps/api/core': { invoke: async (command, args) => {
      f.reads.push([command, args]);
      if (command === 'ai_get_session_binding') return f.binding;
      if (command === 'get_available_templates') return f.available ?? [template];
      throw new Error(`Unexpected command ${command}`);
    }},
    '../utils/workspaceService': { workspaceService: { getState: async () => ({ currentWorkspace: { workspaceId: f.workspaceId } }) } },
    './domainSaveService': { domainSaveService: {
      loadTemplateStates: async projectId => { assert.equal(projectId, 'project-a'); f.reads.push(['states']); if (f.duringRead) await f.duringRead(); return f.saved; },
      loadTemplateState: async projectId => { assert.equal(projectId, 'project-a'); return null; },
      loadTemplateAssets: async (projectId, name) => { assert.equal(projectId, 'project-a'); assert.equal(name, template); return f.assets; },
    }},
    '../utils/projectService': { projectService: { getTemplateAssetPath: async id => {
      if (id === 'missing-on-disk') throw new Error('not found'); return '/synthetic/assets/picture.png';
    }}},
  });
  return { f, ...api };
}
async function main() {
  const { isDemandFormRequest, wantsDemandImageCompletion, demandImageCompletionPrompt } = load('ai/demandFormIntent.ts');
  for (const message of [
    '帮我生成需求导入表', '请生成一份ICT项目需求导入表', '我想填写需求导入表',
    '我会输入需求导入表', '我接下来会输入需求导入表的内容', '我准备填写需求导入表',
    '继续填写需求导入表', '请帮我补齐需求导入表附件1', '可以帮我生成需求导入表吗？',
    '能不能帮我生成需求导入表', '需求导入表帮我生成一下', '请填写《需求导入表》',
    '生成“需求导入表”', '你好，帮我生成需求导入表',
  ]) assert.equal(isDemandFormRequest(message), true, message);
  for (const message of [
    '', '你好', '我已经绑定项目了', '推荐合适产品', '分析项目收益', '需求导入表',
    '需求导入表是什么', '如何生成需求导入表', '生成需求导入表是什么意思',
    '生成需求导入表需要什么材料？', '你支持填写需求导入表吗', '读取需求导入表缺项',
    '看看需求导入表填了多少', '附件1还没传吧', '这张图片是什么',
    '不要生成需求导入表', '我不想填写需求导入表', '以后再生成需求导入表',
    '如果生成需求导入表，会发生什么', '昨天我让你生成需求导入表',
    '只有在我明确要求生成需求导入表的时候才提醒',
    '帮我生成需求导入表，但先不要提醒我上传图片',
    '客户说“帮我生成需求导入表”，这是什么意思',
    '> 帮我生成需求导入表\n解释上面的文字', '```\n帮我生成需求导入表\n```',
  ]) assert.equal(isDemandFormRequest(message), false, message);
  const user = content => ({ role: 'user', content });
  const assistant = content => ({ role: 'assistant', content });
  const greeting = [assistant('您好，我可以帮您生成需求导入表')];
  const filling = [...greeting, user('帮我生成需求导入表'), assistant('需要补图片')];
  assert.equal(wantsDemandImageCompletion(greeting), false, 'greeting/binding cannot activate');
  assert.equal(wantsDemandImageCompletion(filling), true, 'assistant streaming/receipts preserve current user intent');
  for (const text of ['先不填了', '推荐合适产品', '读取需求导入表', '']) {
    assert.equal(wantsDemandImageCompletion([...filling, user(text), assistant('请上传图片')]), false,
      'old requests, assistant responses and image-only turns cannot latch invitation');
  }
  assert.match(demandImageCompletionPrompt(false), /Do not proactively ask/);
  assert.match(demandImageCompletionPrompt(true), /actually missing/);

  const ordinary = fixture();
  const targets = await ordinary.loadDemandUploadTargets('session-a');
  assert.equal(targets.length, 2);
  assert.equal(targets[0].usage, 'attach1'); assert.equal(targets[1].usage, 'attach2');
  assert.equal(targets[0].projectId, 'project-a'); assert.equal(targets[0].sessionId, 'session-a');
  assert.ok(ordinary.f.reads.every(([command]) => !command.includes('context')), 'UI never requests model context');
  // The original no-template-page scenario: no page parameters or model invocation are available.
  ordinary.f.assets = [{ id: 'uploaded', usage: targets[0].usage }];
  const after = await ordinary.loadDemandUploadTargets('session-a');
  assert.equal(after.length, 1); assert.equal(after[0].usage, 'attach2');
  ordinary.f.assets.push({ id: 'uploaded-2', usage: after[0].usage });
  assert.equal((await ordinary.loadDemandUploadTargets('session-a')).length, 0);
  ordinary.f.assets[0].id = 'missing-on-disk';
  assert.equal((await ordinary.loadDemandUploadTargets('session-a'))[0].usage, 'attach1');
  for (const scope of [null, { ...binding, projectId: null }]) {
    const general = fixture({ binding: scope });
    assert.equal((await general.loadDemandUploadTargets('general')).length, 0);
    assert.equal(general.f.reads.length, 1, 'unbound/general must not read any business state');
  }
  const fresh = fixture({ saved: [] });
  assert.equal((await fresh.loadDemandUploadTargets('fresh')).length, 1, 'unique unsaved template can accept assets');
  fresh.f.available = [template, '另一个需求导入表.docx'];
  await assert.rejects(() => fresh.loadDemandUploadTargets('fresh'), /多张/);
  const switched = fixture();
  switched.f.duringRead = async () => { switched.f.workspaceId = 'workspace-b'; };
  await assert.rejects(() => switched.loadDemandUploadTargets('session-a'), /切换/);
  const reset = fixture();
  reset.f.duringRead = async () => { reset.f.binding = null; };
  await assert.rejects(() => reset.loadDemandUploadTargets('session-a'), /切换/);
  const guard = fixture();
  await guard.assertDemandUploadBinding(targets[0]);
  guard.f.binding = { ...binding, projectId: 'project-b' };
  await assert.rejects(() => guard.assertDemandUploadBinding(targets[0]), /绑定/);
  console.log('Demand intent: fresh binding, explicit fill/generate, ordinary/read-only/quoted/negative turns, topic/session switching, clear history and model guidance passed.');
  console.log('Demand card reads: no page/model dependency, catalog slots, upload +1, all-filled, missing file, general/unbound, unsaved/ambiguous template, workspace/reset races and write binding guard passed.');
}
main().catch(error => { console.error(error); process.exitCode = 1; });
