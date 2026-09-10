// Historical defect witness: executes the frozen pre-repair production source.
// This is not an acceptance test. Current behavior is covered by test_structure_commit.cjs.
const assert = require('node:assert/strict');
const { host, solve } = require('./test_structure_commit.cjs');
(async () => {
  for (const autoFix of [false, true]) {
    const h = host(autoFix, true);
    const result = await solve(h);
    const total = Number((h.state.revIt.integration.incl + h.state.revIt.maintenance.incl).toFixed(2));
    assert.equal(result.status, 'success');
    assert.equal(h.spy.mock.callCount(), 1);
    assert.equal(total, autoFix ? 100.01 : 100);
    console.log({ source: 'frozen pre-repair production functions', autoFix, calls: h.spy.mock.callCount(), total, message: result.message });
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
