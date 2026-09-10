import assert from 'node:assert/strict';
import test from 'node:test';
import LamberGateway, { permitted } from '../agent-bridge/webui/lamber-host/gateway.js';
import { TypertGatewayService } from '../agent-bridge/node_modules/@deepseek-ai/dsh-api-gateway/lib/index.js';
import { readFile } from 'node:fs/promises';
const source = (await readFile(new URL('../agent-bridge/webui/lamber-host/session-controller.js', import.meta.url), 'utf8'))
  .replace("'@deepseek-ai/dsh-api-session-controller'", JSON.stringify(new URL('../agent-bridge/node_modules/@deepseek-ai/dsh-api-session-controller/lib/index.js', import.meta.url).href))
  .replace("'dsh-tool-lamber/lib/bridge.js'", JSON.stringify(new URL('../agent-bridge/dsh-tool-lamber/lib/bridge.js', import.meta.url).href));
const { default: LamberSessionController } = await import('data:text/javascript;base64,' + Buffer.from(source).toString('base64'));
import { SessionController } from '../agent-bridge/node_modules/@deepseek-ai/dsh-api-session-controller/lib/index.js';

test('deployment gateway preserves presentation and compact but denies configuration escalation', () => {
  for (const namespace of ['fileReferences', 'directoryPicker', 'credentials', 'plugins', 'skills', 'subagents', 'workflows']) assert.equal(permitted({ namespace, method: 'read', args: {} }), false);
  for (const ns of ['permission', 'llm-deepseek', 'agent-default-model', 'credentials']) assert.equal(permitted({ namespace: 'settings', method: 'update', args: { ns, patch: {} } }), false);
  for (const ns of ['locale', 'ui-theme']) assert.equal(permitted({ namespace: 'settings', method: 'update', args: { ns, patch: {} } }), true);
  for (const line of ['/permission danger-full-access', '/goal run', '/compactly', '/bash echo x']) assert.equal(permitted({ namespace: 'commands', method: 'execute', args: { line } }), false);
  assert.equal(permitted({ namespace: 'commands', method: 'execute', args: { line: '/compact' } }), true);
  assert.equal(permitted({ namespace: 'session', method: 'prompt', args: {} }), true);
});

test('cold session lists and search never expose another workspace', async t => {
  t.mock.method(SessionController.prototype, 'list', async () => ({ items: [{ sessionId: 'own', cwd: process.cwd() }, { sessionId: 'foreign', cwd: '/another-workspace' }] }));
  t.mock.method(SessionController.prototype, 'search', async () => ({ items: [{ sessionId: 'foreign', snippet: 'foreign text' }, { sessionId: 'own', snippet: 'own text' }], hasMore: false }));
  const controller = Object.create(LamberSessionController.prototype);
  assert.deepEqual((await controller.list({})).items.map(x => x.sessionId), ['own']);
  assert.deepEqual((await controller.search({ query: 'text' })).items, [{ sessionId: 'own', snippet: 'own text' }]);
  t.mock.method(SessionController.prototype, 'inspect', async () => ({ meta: { cwd: '/another-workspace' } }));
  await assert.rejects(controller.inspect('foreign'), /不属于当前工作区/);
});

test('actual gateway invocation checks nested session requests before dispatch and narrows settings', async t => {
  const dispatched = [];
  t.mock.method(TypertGatewayService.prototype, 'invoke', async request => {
    dispatched.push(request);
    return { hasDocument: true, namespaces: [{ ns: 'ui-theme' }, { ns: 'llm-deepseek' }, { ns: 'locale' }] };
  });
  const gateway = Object.create(LamberGateway.prototype);
  gateway.ctx = { sessionController: { inspect: async id => { if (id !== 'own') throw new Error('foreign session'); } } };
  for (const args of [{ request: { sessionId: 'foreign' } }, { sessionId: 'foreign' }, { agent: 'foreign' }]) {
    await assert.rejects(gateway.invoke({ namespace: 'session', method: 'prompt', args }), /foreign session/);
  }
  assert.equal(dispatched.length, 0);
  const settings = await gateway.invoke({ namespace: 'settings', method: 'describe', args: {} });
  assert.equal(settings.hasDocument, false);
  assert.deepEqual(settings.namespaces.map(x => x.ns), ['ui-theme', 'locale']);
  await gateway.invoke({ namespace: 'session', method: 'prompt', args: { request: { sessionId: 'own' } } });
  assert.equal(dispatched.length, 2);
});

test('history follow uses the official address DTO and never opens foreign or subagent streams', async t => {
  const opened = [];
  t.mock.method(SessionController.prototype, 'inspect', async id => ({ meta: { cwd: id === 'own' ? process.cwd() : '/foreign' } }));
  t.mock.method(SessionController.prototype, 'follow', async function* (request) { opened.push(request.address.sessionId); yield { type: 'snapshot' }; });
  const controller = Object.create(LamberSessionController.prototype);
  for (const address of [{ kind: 'session', sessionId: 'foreign' }, { kind: 'subagent', parentSessionId: 'own', childSessionId: 'child' }]) {
    await assert.rejects(async () => { for await (const _ of controller.follow({ address })) {} });
  }
  assert.deepEqual(opened, []);
  const frames = [];
  for await (const frame of controller.follow({ address: { kind: 'session', sessionId: 'own' } })) frames.push(frame);
  assert.deepEqual(opened, ['own']);
  assert.equal(frames.length, 1);
  t.mock.restoreAll();
  // SRC dispatch derives named wire arguments from these methods. Preserve the
  // official reserved list name, including its leading underscore.
  const parameters = fn => fn.toString().match(/\(([^)]*)\)/)[1];
  for (const name of ['inspect', 'list', 'search', 'follow', 'control', 'create', 'prompt', 'fork']) {
    assert.equal(parameters(LamberSessionController.prototype[name]), parameters(SessionController.prototype[name]), name);
  }
});

test('the current claimed user message and saved/draft/receipt context enter this same model step', async t => {
  const environment = { LAMBER_BRIDGE_URL: process.env.LAMBER_BRIDGE_URL, LAMBER_BRIDGE_TOKEN: process.env.LAMBER_BRIDGE_TOKEN, LAMBER_PARENT_PID: process.env.LAMBER_PARENT_PID };
  process.env.LAMBER_BRIDGE_URL = 'http://127.0.0.1:1'; process.env.LAMBER_BRIDGE_TOKEN = 'synthetic-test'; process.env.LAMBER_PARENT_PID = String(process.pid);
  const responses = { bootstrap: { cwd: process.cwd(), workspaceName: 'fixture' }, admit: { projectId: 'p' }, business: {}, 'business-result': { result: { ok: true, value: { savedOfficial: ['saved'], draftOverlay: ['draft'] } } }, 'release-read': {}, receipts: ['committed result'], 'legacy-context': ['legacy result'] };
  let contextRequest;
  t.mock.method(globalThis, 'fetch', async (url, options) => {
    const method = new URL(url).pathname.split('/').at(-1);
    if (method === 'business') contextRequest = JSON.parse(options.body).payload;
    assert.ok(Object.hasOwn(responses, method), method);
    return new Response(JSON.stringify(responses[method]), { status: 200 });
  });
  const pluginSource = (await readFile(new URL('../agent-bridge/webui/lamber-host/host-policy.js', import.meta.url), 'utf8'))
    .replace("'@deepseek-ai/dsh-llm'", JSON.stringify(new URL('../agent-bridge/node_modules/@deepseek-ai/dsh-llm/lib/index.js', import.meta.url).href))
    .replace("'./business-presentation.generated.js'", JSON.stringify(new URL('../agent-bridge/webui/lamber-host/business-presentation.generated.js', import.meta.url).href))
    .replace("'dsh-tool-lamber/lib/bridge.js'", JSON.stringify(new URL('../agent-bridge/dsh-tool-lamber/lib/bridge.js', import.meta.url).href));
  const { apply } = await import('data:text/javascript;base64,' + Buffer.from(pluginSource).toString('base64'));
  const handlers = new Map(), dispose = [];
  const session = { id: 's', header: { cwd: process.cwd() }, snapshotEvents: () => [] };
  try {
    await apply({ typertGateway: { lamberWebDeployment: true }, sessionController: { lamberWebDeployment: true }, workspaceController: { lamberWebDeployment: true }, agentDefaultModel: { lamberWebDeployment: true }, sessions: { get: () => session }, workspaceRegistry: { create: async () => {} }, settings: { get: () => ({ preference: 'zh' }) }, on: (event, handler) => handlers.set(event, handler), effect: setup => dispose.push(setup()), connection: { rpc: { handle: () => () => {} } } });
    const user = { source: { kind: 'user' }, content: [{ type: 'text', text: '本轮明确要求' }] };
    const agent = { session };
    const result = await handlers.get('agent/pre-step')({ agent, messages: [user], signal: new AbortController().signal }, async () => ({ kind: 'enter', messages: [user] }));
    assert.equal(contextRequest.userMessage, '本轮明确要求');
    assert.equal(result.messages.at(-1), user);
    assert.equal(result.messages[0].source.kind, 'plugin');
    for (const marker of ['savedOfficial', 'draftOverlay', 'committed result', 'legacy result']) assert.ok(result.messages[0].content[0].text.includes(marker), marker);
  } finally {
    dispose.reverse().forEach(fn => fn?.());
    for (const [key, value] of Object.entries(environment)) { if (value === undefined) delete process.env[key]; else process.env[key] = value; }
  }
});


test('workspace follow reconnect and increments do not leak foreign names, paths or archive IDs', async () => {
  const source = (await readFile(new URL('../agent-bridge/webui/lamber-host/workspace-controller.js', import.meta.url), 'utf8'))
    .replace("'@deepseek-ai/dsh-api-workspace-controller'", JSON.stringify(new URL('../agent-bridge/node_modules/@deepseek-ai/dsh-api-workspace-controller/lib/index.js', import.meta.url).href));
  const { scopedWorkspaceFrame: scope } = await import('data:text/javascript;base64,' + Buffer.from(source).toString('base64'));
  const known = new Set(), sessions = new Set(['s']);
  const own = { workspaceId: 'w', path: '/own', sessionIds: ['s'] }, foreign = { workspaceId: 'other', path: '/private/other', sessionIds: ['foreign'] };
  const baseline = { type: 'baseline', value: { items: [own, foreign], archivedSessionIds: ['s', 'foreign'] } };
  assert.deepEqual(scope(baseline, '/own', known, sessions).value, { items: [own], archivedSessionIds: ['s'] });
  assert.equal(scope({ type: 'upsert', workspace: foreign }, '/own', known, sessions), null);
  assert.deepEqual(scope({ type: 'order', workspaceIds: ['other', 'w'] }, '/own', known, sessions).workspaceIds, ['w']);
  assert.deepEqual(scope({ type: 'archived', archivedSessionIds: ['foreign', 's'] }, '/own', known, sessions).archivedSessionIds, ['s']);
  assert.equal(scope({ type: 'remove', workspaceId: 'other' }, '/own', known, sessions), null);
  assert.deepEqual(scope({ type: 'remove', workspaceId: 'w' }, '/own', known, sessions), { type: 'remove', workspaceId: 'w' });
  assert.deepEqual(scope(baseline, '/own', known, sessions).value.items, [own]);
});
