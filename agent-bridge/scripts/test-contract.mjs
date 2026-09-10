import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { once } from 'node:events';
import { apply } from '../dsh-tool-lamber/lib/index.js';
import { BRIDGE_CONTRACT } from '../dsh-tool-lamber/lib/contract.generated.js';
import { MISMATCH_MESSAGE, UNREACHABLE_MESSAGE } from '../dsh-tool-lamber/lib/contract.js';

let mode = 'match', requests = 0;
const server = createServer(async (req, res) => {
  requests++;
  let raw = ''; for await (const part of req) raw += part;
  assert.deepEqual(JSON.parse(raw), BRIDGE_CONTRACT);
  assert.equal(req.url, '/lamber-bridge/handshake');
  assert.equal(req.headers['x-lamber-bridge-token'], 'synthetic');
  res.setHeader('content-type', 'application/json');
  if (mode === 'old') { res.writeHead(404); res.end('{"error":"未知的 AI 桥接路由"}'); }
  else if (mode === 'mismatch') { res.writeHead(409); res.end('{"error":"version"}'); }
  else if (mode === 'token') { res.writeHead(401); res.end('{"error":"token"}'); }
  else if (mode === 'malformed') res.end('{}');
  else res.end(JSON.stringify(BRIDGE_CONTRACT));
});
server.listen(0, '127.0.0.1'); await once(server, 'listening');
process.env.LAMBER_BRIDGE_URL = `http://127.0.0.1:${server.address().port}`;
process.env.LAMBER_BRIDGE_TOKEN = 'synthetic';
const registered = [];
const ctx = { tools: { register: tool => registered.push(tool.name), guard: () => {} }, on: () => {} };
try {
  for (mode of ['old', 'mismatch', 'token', 'malformed']) {
    const before = requests;
    await assert.rejects(apply(ctx), { message: mode === 'token' ? UNREACHABLE_MESSAGE : MISMATCH_MESSAGE });
    assert.equal(registered.length, 0, 'no tool registered on failure');
    assert.equal(requests, before + 1, 'no retry');
  }
  mode = 'match'; await apply(ctx);
  assert.deepEqual(registered.sort(), ['calculate_selection_fee', 'fill_template_fields', 'query_projects', 'read_benefit_inputs', 'read_template_fields', 'reverse_calculate_selection_fee', 'run_benefit_calculation', 'simulate_benefit_calculation', 'write_test_marker']);
  server.close(); await once(server, 'close');
  registered.length = 0;
  await assert.rejects(apply(ctx), { message: UNREACHABLE_MESSAGE });
  assert.equal(registered.length, 0);
  for (const message of [MISMATCH_MESSAGE, UNREACHABLE_MESSAGE]) {
    assert.doesNotMatch(message, /404|401|lamber-bridge|workspace_handler/);
  }
  console.log('PASS: actual HTTP handshake, old route, mismatch, token, malformed reply, offline; registration fails closed without retries');
} finally { server.close(); }
