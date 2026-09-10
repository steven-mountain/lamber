// Only mounted by the isolated Rust Host integration test; never a product plugin.
import assert from 'node:assert/strict';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
export const inject = ['typertGateway', 'sessionController', 'workspaceController', 'workspaceRegistry', 'sessions', 'settingsController'];
export async function apply(ctx) {
  const cwd = process.cwd(), foreign = join(dirname(cwd), 'foreign-workspace');
  mkdirSync(foreign, { recursive: true });
  await ctx.workspaceRegistry.create(cwd, 'own');
  await ctx.workspaceRegistry.create(foreign, 'foreign');
  for (const [id, path] of [['contract-own', cwd], ['contract-foreign', foreign]]) {
    const session = ctx.sessions.create(id, { meta: { cwd: path } });
    session.append('user/message', { id: 'message-' + id, source: { kind: 'user' }, content: [{ type: 'text', text: id }] }, { surfaceOp: 'append' });
  }
  const gateway = ctx.typertGateway;
  const invoke = (namespace, method, args = {}) => gateway.invoke({ namespace, method, args, signal: AbortSignal.timeout(5000) });
  const settings = await invoke('settings', 'describe');
  assert.equal(settings.hasDocument, false);
  assert.ok(settings.namespaces.every(row => ['locale', 'ui-theme'].includes(row.ns)));
  const sessions = await invoke('session', 'list', { _request: {} });
  assert.ok(sessions.items.some(row => row.sessionId === 'contract-own'));
  assert.ok(sessions.items.every(row => row.cwd === cwd));
  const inspected = await ctx.sessionController.inspect('contract-own');
  assert.equal(inspected.meta.cwd, cwd);
  await assert.rejects(ctx.sessionController.inspect('contract-foreign'), /不属于当前工作区/);
  await assert.rejects(invoke('session', 'rename', { request: { sessionId: 'contract-foreign', title: 'forbidden' } }), /不属于当前工作区/);
  await assert.rejects(invoke('settings', 'update', { ns: 'llm-deepseek', patch: {} }), /Lamber/);
  for (const [namespace, method, args] of [
    ['workspace', 'follow', {}],
    ['session', 'follow', { request: { address: { kind: 'session', sessionId: 'contract-own' } } }],
    ['session', 'control', {}],
  ]) {
    const controller = new AbortController();
    const stream = await gateway.stream({ namespace, method, args, signal: controller.signal });
    const iterator = stream[Symbol.asyncIterator]();
    try {
      const frame = (await iterator.next()).value;
      assert.ok(frame, namespace + '.' + method);
      assert.ok(!JSON.stringify(frame).includes('contract-foreign'));
      assert.ok(!JSON.stringify(frame).includes('foreign-workspace'));
      if (namespace === 'workspace') assert.ok(frame.value.items.some(row => row.path === cwd));
    } finally { controller.abort(); await iterator.return?.(); }
  }
  const denied = await gateway.stream({ namespace: 'session', method: 'follow', args: { request: { address: { kind: 'session', sessionId: 'contract-foreign' } } }, signal: AbortSignal.timeout(5000) });
  await assert.rejects(denied[Symbol.asyncIterator]().next(), /不属于当前工作区/);
  writeFileSync(join(dirname(cwd), 'host-contract-passed.json'), JSON.stringify({ settings: true, namedArguments: true, sessionBoundary: true, scopedStreams: true }));
}
