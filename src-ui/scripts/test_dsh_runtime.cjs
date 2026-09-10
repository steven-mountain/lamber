const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const snapshots = new Map();
const browser = { localStorage: { getItem: key => snapshots.get(key) ?? null, setItem: (key, value) => snapshots.set(key, value) }, addEventListener() {} };
const modules = new Map();
function load(file) {
  file = path.resolve(file);
  if (modules.has(file)) return modules.get(file).exports;
  const ref = { exports: {} }; modules.set(file, ref);
  const code = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } }).outputText;
  vm.runInNewContext(code, { module: ref, exports: ref.exports, console, setTimeout, clearTimeout, setInterval, clearInterval, TextDecoder, TextEncoder, Uint8Array, window: browser,
    crypto: require('node:crypto').webcrypto,
    require: name => name.startsWith('.') ? load(path.resolve(path.dirname(file), `${name}.ts`))
      : name.startsWith('@tauri-apps/') ? {} : require(name),
  }, { filename: file });
  return ref.exports;
}
const { DshRuntime, reduceDshUpdate, imageBlocks } = load(path.join(__dirname, '../src/ai/DshRuntime.ts'));
const { useAiSessionStore: store } = load(path.join(__dirname, '../src/store/useAiSessionStore.ts'));
const { PromptRenderer } = load(path.join(__dirname, '../src/ai/PromptRenderer.ts'));
const fixture = require('./fixtures/dsh-real-stream.json');
function frame(method, params, requestId = 'turn-a', sessionId = 'front-a') {
  return { method, params, requestId, lamberSessionId: sessionId };
}
async function run() {
  // Local history survives retirement; only stored ACP ids can be resumed.
  const old = store.getState().createSession();
  store.getState().appendMessages(old, [{ role: 'user', content: '旧版历史' }]);
  store.getState().flushPersistence();
  store.getState().hydrateFromStorage();
  assert.ok(store.getState().sessions.find(session => session.id === old).messages.some(message => message.content === '旧版历史'));
  const { DshMessageProjection } = load(path.join(__dirname, '../src/ai/DshMessageProjection.ts'));
  const delta = (step, seq, type, bytes, extra = {}) => ({ kind: 'delta', turn: 1, step, seq, chunk: { type, index: type === 'reasoning-delta' ? 0 : 1, bytes: Array.from(bytes), ...extra } });
  const commit = (step, seq, messageId) => ({ kind: 'commit', turn: 1, step, seq, messageId });
  const committedText = (messageId, content, kind = 'agent_message_chunk') => ({ messageId, sessionUpdate: kind, content: { type: 'text', text: content } });
  // Every possible UTF-8 byte cut, including Chinese and emoji.
  const bytes = new TextEncoder().encode('中文🙂1,234.56元');
  for (let cut = 1; cut < bytes.length; cut++) {
    const p = new DshMessageProjection();
    p.stream(delta(1, 1, 'text-delta', bytes.slice(0, cut)));
    const result = p.stream(delta(1, 2, 'text-delta', bytes.slice(cut)));
    assert.equal(result.content, '中文🙂1,234.56元');
  }
  // Both transport orders, multiple steps, sparse tool status and late duplicate.
  for (const acpFirst of [true, false]) {
    const p = new DshMessageProjection();
    p.stream(delta(1, 1, 'reasoning-delta', new TextEncoder().encode('预览思考')));
    p.stream(delta(1, 2, 'text-delta', new TextEncoder().encode('错误预览')));
    p.stream(delta(1, 3, 'tool-call-delta', new TextEncoder().encode('{"p":'), { index: 2, id: 'call', name: 'run_benefit_calculation' }));
    const applyAcp = () => { p.update(committedText('msg1', '正式思考', 'agent_thought_chunk')); p.update(committedText('msg1', '正式正文')); };
    if (acpFirst) applyAcp();
    p.stream(commit(1, 4, 'msg1'));
    if (!acpFirst) applyAcp();
    assert.equal(p.stream(delta(1, 5, 'text-delta', [65])).content, '正式正文');
    p.update({ sessionUpdate: 'tool_call', toolCallId: 'call', title: 'run_benefit_calculation', rawInput: { p: 1 }, status: 'in_progress' });
    p.update({ sessionUpdate: 'tool_call_update', toolCallId: 'call', status: 'completed' });
    assert.equal(p.stream(delta(2, 6, 'text-delta', new TextEncoder().encode('后续'))).content, '正式正文后续');
    p.update(committedText('msg2', '后续一'));
    p.update(committedText('msg2', '后续二')); // Separate committed blocks of one message.
    const result = p.stream(commit(2, 7, 'msg2'));
    assert.equal(result.content, '正式正文后续一后续二');
    assert.equal(result.think, '正式思考');
    assert.equal(result.toolCalls.length, 1);
    assert.equal(result.toolCalls[0].status, 'completed');
    assert.equal(p.finish().content, '正式正文后续一后续二');
  }
  // Terminal wins when HTTP metadata never catches up (including cancellation).
  const final = new DshMessageProjection();
  final.stream(delta(1, 1, 'text-delta', [65, 66, 67]));
  final.update(committedText('cancelled-prefix', 'A'));
  assert.equal(final.finish().content, 'A');
  // Replay actual plugin HTTP + ACP recordings with the production runtime.
  for (const file of process.argv.slice(2)) {
    const events = JSON.parse(fs.readFileSync(file, 'utf8'));
    let handler, last, progressive = 0, firstCommit = false;
    const authority = events.filter(e => e.method === 'session/update').reduce((message, e) => reduceDshUpdate(message, e.params.update), { role: 'assistant', content: '' });
    const runtime = new DshRuntime({ listen: async cb => { handler = cb; return () => {}; }, invoke: async () => {
      for (const event of events) {
        if (event.method === 'session/update' && ['agent_message_chunk', 'agent_thought_chunk'].includes(event.params.update.sessionUpdate)) firstCommit = true;
        handler(event);
      }
      return events.at(-1).params.sessionId;
    } });
    await runtime.execute({ sessionId: events[0].lamberSessionId, requestId: events[0].requestId, text: '', signal: new AbortController().signal,
      onSession() {}, onUpdate: msg => { if (!firstCommit && (msg.content || msg.think)) progressive++; last = msg; } });
    assert.ok(progressive >= 3, 'Actual subscription must display incrementally before ACP commits');
    assert.equal(last.content, authority.content);
    assert.equal(last.think, authority.think);
    assert.ok(!last.content.includes('�'));
    console.log(`stream recording ${path.basename(file)}: ${progressive} pre-commit updates; final ACP text identical`);
  }

  // Actual observed ACP payloads, including UTF-8 text and the final-event order.
  for (const [name, events] of Object.entries(fixture)) {
    let handler, unlistened = false, last;
    const runtime = new DshRuntime({
      listen: async cb => { handler = cb; return () => { unlistened = true; }; },
      invoke: async () => { for (const [method, params] of events) handler(frame(method, params)); return 'acp-a'; },
    });
    await runtime.execute({ sessionId: 'front-a', requestId: 'turn-a', text: '', signal: new AbortController().signal,
      onUpdate: msg => { last = msg; }, onSession: id => assert.equal(id, 'acp-a') });
    assert.ok(unlistened);
    if (name === 'completed') assert.equal(last.content, '停止后可继续。');
  }
  // Switching the visible session cannot redirect output, even before invoke resolves.
  const a = store.getState().createSession();
  const b = store.getState().createSession();
  store.getState().appendMessages(a, [{ role: 'assistant', content: '' }]);
  let handler;
  const runtime = new DshRuntime({ listen: async cb => { handler = cb; return () => {}; }, invoke: async () => {
    handler(frame('session/update', { update: { sessionUpdate: 'agent_message_chunk', content: { type: 'text', text: '错误会话' } } }, 'turn-a', b));
    for (const [method, params] of fixture.completed) handler(frame(method, params, 'turn-a', a));
    handler(frame('session/update', { update: { sessionUpdate: 'agent_message_chunk', content: { type: 'text', text: '迟到片段' } } }, 'turn-a', a));
    return 'acp-a';
  } });
  await runtime.execute({ sessionId: a, requestId: 'turn-a', text: '', signal: new AbortController().signal,
    onUpdate: msg => store.getState().updateLastAssistantMessage(a, msg), onSession: id => store.getState().setHarnessSessionId(a, id) });
  assert.equal(store.getState().currentSessionId, b);
  assert.equal(store.getState().sessions.find(s => s.id === a).messages.at(-1).content, '停止后可继续。');
  assert.ok(!store.getState().sessions.find(s => s.id === b).messages.some(m => m.content.includes('停止后')));
  store.getState().flushPersistence(); store.getState().hydrateFromStorage();
  assert.equal(store.getState().sessions.find(s => s.id === a).harnessSessionId, 'acp-a');
  store.getState().resetSessionMessages(a);
  assert.equal(store.getState().sessions.find(s => s.id === a).harnessSessionId, undefined);
  // Sparse tool updates must merge into the call, never erase its title/input.
  let msg = { role: 'assistant', content: '' };
  for (const update of [
    { sessionUpdate: 'agent_message_chunk', content: { type: 'text', text: '中文🙂' } },
    { sessionUpdate: 'agent_thought_chunk', content: { type: 'text', text: '思考' } },
    { sessionUpdate: 'agent_message_chunk', content: { type: 'text', text: '尾块' } },
    { sessionUpdate: 'tool_call', toolCallId: 'tool-1', title: 'run_benefit_calculation', rawInput: { projectId: 1 } },
    { sessionUpdate: 'tool_call_update', toolCallId: 'tool-1', status: 'completed', content: [{ type: 'content', content: { type: 'text', text: 'NPV' } }] },
  ]) msg = reduceDshUpdate(msg, update);
  assert.equal(msg.content, '中文🙂尾块'); assert.equal(msg.think, '思考');
  assert.equal(msg.toolCalls.length, 1); assert.equal(msg.toolCalls[0].input.projectId, 1);
  assert.equal(msg.toolCalls[0].title, 'run_benefit_calculation'); assert.equal(msg.toolCalls[0].status, 'completed');
  // Abort during startup sends cancel only after prompt admission and awaits terminal.
  const abort = new AbortController(); let calls = [], release, unlistened = false;
  const cancelled = new DshRuntime({ listen: async cb => { handler = cb; return () => { unlistened = true; }; }, invoke: async command => {
    calls.push(command);
    if (command === 'ai_send_prompt') { abort.abort(); await new Promise(resolve => { release = resolve; }); return 'acp-a'; }
    handler(frame('session/turn-ended', { stopReason: 'Cancelled' }));
  } });
  const task = cancelled.execute({ sessionId: 'front-a', requestId: 'turn-a', text: '', signal: abort.signal, onUpdate() {}, onSession() {} });
  await new Promise(resolve => setImmediate(resolve));
  assert.deepEqual(calls, ['ai_send_prompt']); assert.equal(unlistened, false);
  release(); await task;
  assert.deepEqual(calls, ['ai_send_prompt', 'ai_cancel_prompt']); assert.equal(unlistened, true);
  // Terminal errors and failed invokes release listeners, including early terminal errors.
  for (const earlyEvent of [false, true]) {
    let removed = false;
    const failing = new DshRuntime({ listen: async cb => { handler = cb; return () => { removed = true; }; }, invoke: async () => {
      if (earlyEvent) { handler(frame('session/turn-ended', { error: 'upstream unavailable' })); return 'acp-a'; }
      throw new Error('startup failed');
    } });
    await assert.rejects(failing.execute({ sessionId: 'front-a', requestId: 'turn-a', text: '', signal: new AbortController().signal, onUpdate() {}, onSession() {} }));
    assert.equal(removed, true);
  }
  assert.equal(imageBlocks([{ name: 'test', dataUrl: 'data:image/png;base64,aGVsbG8=' }])[0].mimeType, 'image/png');
  assert.throws(() => imageBlocks([{ name: 'lost attachment' }]));
  const text = new PromptRenderer().render({ systemRules: [{ id: 'units', priority: 1, content: '金额单位为元' }],
    dynamicState: { layer1Core: [{ type: 'json', title: '已保存', content: { amount: 1200 } }], layer2Active: [{ type: 'json', title: '未保存', content: { amount: 2400 } }], layer3Context: [] }, userIntent: { raw: '比较两者' } });
  assert.ok(text.includes('1200') && text.includes('2400') && text.includes('金额单位为元') && text.includes('# USER_INTENT\n比较两者'));
  console.log('dsh adapter: real-event replay, isolation, persistence, sparse tools, cancel races, errors, images and context passed');
}
run().catch(error => { console.error(error); process.exitCode = 1; });
