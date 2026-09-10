const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const ts = require('typescript');
const root = path.join(__dirname, '../src');
function load(relative, mocks = {}) {
  const mod = { exports: {} };
  const filename = path.join(root, relative);
  const code = ts.transpileModule(fs.readFileSync(filename, 'utf8'), { compilerOptions: {
    module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, esModuleInterop: true,
  }}).outputText;
  vm.runInNewContext(code, { module: mod, exports: mod.exports, ...mocks.globals, require: name => {
    if (mocks[name]) return mocks[name];
    return load(path.relative(root, path.resolve(path.dirname(filename), name)) + '.ts', mocks);
  }});
  return mod.exports;
}
const geometry = load('lib/aiWindowPlacement.ts');
const main = { workArea: { position: { x: 0, y: 50 }, size: { width: 2940, height: 1810 } } };
const upper = { workArea: { position: { x: 0, y: -2160 }, size: { width: 3840, height: 2100 } } };
const size = { width: 1560, height: 1360 };
const plain = value => JSON.parse(JSON.stringify(value));
function fixture({ existing = true, position = { x: 838, y: -1604 }, saved = null, fail = false } = {}) {
  const calls = [], storage = new Map(saved ? [[geometry.AI_WINDOW_POSITION_KEY, JSON.stringify(saved)]] : []);
  const handlers = new Map();
  let exists = existing, rejectShow = fail, creates = 0;
  const target = {
    once: async (event, handler) => { handlers.set(event, handler); return () => handlers.delete(event); },
    unminimize: async () => calls.push('unminimize'),
    outerPosition: async () => position, outerSize: async () => size, scaleFactor: async () => 2,
    setSize: async value => calls.push(['size', plain(value)]),
    setPosition: async value => { calls.push(['position', plain(value)]); position = value; },
    show: async () => { if (rejectShow) throw new Error('native show failure'); calls.push('show'); },
    setFocus: async () => calls.push('focus'),
  };
  class WebviewWindow {
    static async getByLabel() { return exists ? target : null; }
    constructor(label, options) { creates++; calls.push(['create', options]); exists = true; return target; }
  }
  class Point { constructor(x,y) { this.x=x; this.y=y; } }
  class Size { constructor(width,height) { this.width=width; this.height=height; } }
  const service = load('services/aiAssistantWindow.ts', {
    globals: { window: { __TAURI_INTERNALS__: {} }, localStorage: {
      getItem: key => storage.get(key) ?? null, setItem: (key,value) => storage.set(key,value),
    }},
    '@tauri-apps/api/event': { emit: async () => {}, emitTo: async () => {} },
    '@tauri-apps/api/webviewWindow': { WebviewWindow },
    '@tauri-apps/api/window': { availableMonitors: async () => [main], currentMonitor: async () => main, PhysicalPosition: Point, PhysicalSize: Size },
    '../store/useAiContextStore': { AI_CONTEXT_REFRESH_REQUEST_EVENT: 'refresh' },
  });
  return { ...service, calls, storage, handlers, creates: () => creates, recover: () => { rejectShow = false; } };
}
async function mainTest() {
  const visible = { x: 400, y: 150 };
  assert.deepEqual(plain(geometry.placeAiWindow(visible, size, [main], main).position), visible);
  for (const y of [-802, -957]) {
    const placed = geometry.placeAiWindow({ x: 419*2, y: y*2 }, size, [main], main);
    assert.ok(placed.position.y >= main.workArea.position.y, 'reported stale coordinates recover on internal screen');
    assert.ok(placed.position.y + size.height <= 1860);
  }
  const negative = { x: 500, y: -2000 };
  assert.deepEqual(plain(geometry.placeAiWindow(negative, size, [main,upper], main).position), negative, 'connected negative-coordinate monitor preserved');
  const oversized = geometry.placeAiWindow({x:99999,y:99999},{width:5000,height:3000},[main],main);
  assert.deepEqual(plain(oversized.size),main.workArea.size);
  assert.equal(geometry.parseWindowPosition('{"x":1,"y":1e999}'),null);
  assert.equal(geometry.parseWindowPosition('{bad'),null);
  assert.equal(geometry.parseWindowPosition('{"x":1,"y":2,"version":99}'),null);
  const restore = fixture();
  await restore.openAiAssistantWindow('hub');
  assert.equal(restore.calls[0], 'unminimize');
  assert.equal(restore.calls.at(-1), 'focus');
  assert.ok(restore.calls.find(call => call[0] === 'position')[1].y >= 50);
  assert.equal(JSON.parse(restore.storage.get(geometry.AI_WINDOW_POSITION_KEY)).version,2);
  const create = fixture({ existing:false, saved:{x:539,y:-957} });
  const a = create.openAiAssistantWindow('hub'), b = create.openAiAssistantWindow('project_board');
  assert.equal(a,b,'concurrent opens share native creation lifecycle');
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(create.creates(),1);
  assert.equal(create.calls[0][1].visible,false);
  assert.equal(create.calls[0][1].x,undefined,'never pass unchecked saved coordinates into native creation');
  create.handlers.get('tauri://created')({payload:null});
  await a;
  assert.ok(create.calls.find(call => call[0] === 'position')[1].y >= 50);
  const physical = fixture({ existing:false, saved:{version:2,x:500,y:100} });
  const physicalOpen = physical.openAiAssistantWindow('hub');
  await new Promise(resolve => setImmediate(resolve));
  physical.handlers.get('tauri://created')({payload:null});
  await physicalOpen;
  assert.deepEqual(physical.calls.find(call => call[0] === 'position')[1], {x:500,y:100}, 'physical cache must not be scaled twice');
  assert.equal(physical.handlers.size,0,'native lifecycle listeners cleaned up');
  const fail = fixture({fail:true});
  await assert.rejects(() => fail.openAiAssistantWindow('hub'), /native show failure/);
  fail.recover(); await fail.openAiAssistantWindow('hub');
  const failCreate = fixture({existing:false});
  const pending = failCreate.openAiAssistantWindow('hub');
  await new Promise(resolve => setImmediate(resolve));
  failCreate.handlers.get('tauri://error')({payload:'native creation failure'});
  await assert.rejects(() => pending,/native creation failure/);
  console.log('AI window: reported negative coordinates, connected upper monitor, visible-position preservation, oversize, invalid cache, minimized reuse, creation lifecycle, concurrent clicks and retry after native failure passed.');
}
mainTest().catch(error => { console.error(error); process.exitCode=1; });
