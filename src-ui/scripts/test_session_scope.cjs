const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const root = path.join(__dirname, '../src');
const calls = [];
let page = 'B';
const projects = [{ projectId: 'A', projectName: '甲项目' }, { projectId: 'B', projectName: '乙项目' }];
const mocks = {
  aiProjectContextService: {
    listAiWorkspaceProjects: async () => { calls.push('index'); return projects; },
    buildAiProjectContext: async request => {
      calls.push(request.projectId);
      return { projectId: request.projectId, projectName: projects.find(item => item.projectId === request.projectId).projectName,
        overview: { name: request.projectId, description: `secret-${request.projectId}` }, sources: [], warnings: [], templates: [] };
    },
  },
  useAiContextStore: { useAiContextStore: { getState: () => ({ activeModule: 'ict.core', businessData: { 'ict.core': { projectId: page, description: `draft-secret-${page}` } }, lastUpdated: {} }) } },
  useSaveStore: { useSaveStore: { getState: () => ({ dirtyScopes: ['lifecycle'] }) } },
  useProjectStore: { readStoredCurrentProject: () => ({ id: page }) },
  useNavigationStore: { readStoredNavigationState: () => ({ activeProjectId: page }) },
};
const cache = new Map();
function load(file) {
  file = path.resolve(file);
  if (file.endsWith(".json")) return JSON.parse(fs.readFileSync(file,"utf8"));
  const name = path.basename(file, '.ts');
  if (mocks[name]) return mocks[name];
  if (cache.has(file)) return cache.get(file).exports;
  const module = { exports: {} }; cache.set(file, module);
  const code = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { esModuleInterop: true, module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } }).outputText;
  vm.runInNewContext(code, { module, exports: module.exports,
    require: request => request === '@tauri-apps/api/core' ? { invoke: async () => [] }
      : request.startsWith('.') ? load(path.resolve(path.dirname(file), request.endsWith(".json") ? request : `${request}.ts`)) : require(request),
  });
  return module.exports;
}
async function run() {
  const policy = load(path.join(root, 'ai/sessionScopePolicy.ts'));
  assert.match(policy.GENERAL_SESSION_LABEL, /汇总与甄选费计算/);
  for (const bound of ['A', null]) {
    assert.match(policy.sessionScopePrompt(bound), /query_projects/);
    assert.match(policy.sessionScopePrompt(bound), /isBoundProject/);
  }
  assert.match(policy.sessionScopePrompt(null), /No single-project calculation/);
  assert.match(policy.sessionScopePrompt(null), /two pure selection-fee calculators/);
  assert.match(policy.sessionScopePrompt('A'), /projectId=A/);
  const { buildAiChatContext } = load(path.join(root, 'ai/context/buildAiChatContext.ts'));
  let result = await buildAiChatContext({ currentView: 'ict', userMessage: '请分析效益', boundProjectId: 'A' });
  assert.ok(calls.includes('A'));
  assert.ok(!calls.includes('B'));
  assert.ok(!JSON.stringify(result).includes('secret-B'));
  assert.equal(result.draftOverlay, undefined, 'another active page cannot overlay the binding');
  calls.length = 0;
  result = await buildAiChatContext({ currentView: 'ict', userMessage: '乙项目的成本是多少', boundProjectId: 'A' });
  assert.ok(!calls.includes('B'));
  assert.ok(!calls.includes('A'), 'must not substitute A for a named B request');
  calls.length = 0;
  result = await buildAiChatContext({ currentView: 'ict', userMessage: '全部项目的成本', boundProjectId: null });
  assert.equal(calls.length, 0, 'general chat cannot silently read the project index or state');
  assert.ok(!JSON.stringify(result).includes('secret-'));
  page = 'A'; calls.length = 0;
  result = await buildAiChatContext({ currentView: 'ict', userMessage: '请分析效益', boundProjectId: 'A' });
  assert.equal(result.draftOverlay.projectId, 'A');
  calls.length = 0;
  await assert.rejects(buildAiChatContext({ currentView: 'ict', userMessage: '乙项目的成本是多少' }), /缺少会话项目权限/);
  assert.equal(calls.length, 0, 'missing binding cannot activate unscoped routing');
  console.log('PASS: bound-project reads, cross-project named requests, switched-page draft, general chat, missing binding fails closed');
}
run().catch(error => { console.error(error); process.exitCode = 1; });
