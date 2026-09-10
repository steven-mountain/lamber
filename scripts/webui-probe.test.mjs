import assert from 'node:assert/strict';
import test from 'node:test';
import { apply } from './fixtures/webui-probe/index.js';

test('probe denies all tools and restricts model admission to its synthetic cwd', async () => {
  let guard;
  let beforeStep;
  let created;
  await apply({
    tools: { guard: callback => { guard = callback; } },
    workspaceRegistry: { create: async (cwd, title) => { created = { cwd, title }; } },
    on: (event, callback) => {
      assert.equal(event, 'agent/pre-step');
      beforeStep = callback;
    },
  });
  assert.equal(created.cwd, process.cwd());
  for (const name of ['bash', 'fill_template_fields', 'run_benefit_calculation', 'unknown']) {
    assert.match(guard({ name }), /未开放业务工具/);
  }
  let modelAdmissions = 0;
  const next = async () => { modelAdmissions++; return { kind: 'proceed' }; };
  assert.deepEqual(await beforeStep({ agent: { session: { header: { cwd: process.cwd() } } } }, next), { kind: 'proceed' });
  assert.equal(modelAdmissions, 1);
  for (const cwd of ['', '/another-project', undefined]) {
    await assert.rejects(beforeStep({ agent: { session: { header: { cwd } } } }, next), /仅允许/);
  }
  assert.equal(modelAdmissions, 1, 'rejected directories never reach the model');
});
