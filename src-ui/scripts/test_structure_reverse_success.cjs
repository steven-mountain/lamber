const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const { mock } = require('node:test');
const load = require('./load_ts.cjs');

const root = path.resolve(__dirname, '..');
const sourcePath = path.join(root, 'src/hooks/useIctCalculations.ts');
const source = fs.readFileSync(sourcePath, 'utf8');
const ast = ts.createSourceFile(sourcePath, source, ts.ScriptTarget.Latest, true);
const names = ['performLockedTotalStructureReverseCalculation', 'performReverseCalculation', 'getMetricValue',
  'METRIC_EPSILON', 'MONEY_EPSILON', 'roundMoney', 'formatCurrency', 'formatPercent'];
const declarations = {};
function visit(node) {
  if (ts.isVariableDeclaration(node) && names.includes(node.name.getText(ast))) {
    declarations[node.name.getText(ast)] = node.initializer.getText(ast);
  }
  ts.forEachChild(node, visit);
}
visit(ast);
assert.equal(Object.keys(declarations).length, names.length);
const baseline = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures/structure-reverse-before-success-check.json'), 'utf8')).declarations;
const beforeCardSource = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures/structure-reverse-before-card.json'), 'utf8')).function;
const beforeCardAst = ts.createSourceFile('before-card.ts', beforeCardSource, ts.ScriptTarget.Latest, true);
const beforeCardFunction = beforeCardAst.statements[0].declarationList.declarations[0].initializer.getText(beforeCardAst);

for (const name of ['METRIC_EPSILON', 'MONEY_EPSILON', 'roundMoney', 'formatCurrency', 'formatPercent']) {
  assert.equal(declarations[name], baseline[name], `${name} must remain unchanged`);
}
assert.equal(Number(declarations.METRIC_EPSILON), 0.0001);
assert.equal(Number(declarations.MONEY_EPSILON), 0.004);
assert.equal((source.match(/const modelEAmountMode = false;/g) || []).length, 3);
const { buildLockedTotalStructureSamplePoints } = load(path.join(root, 'src/lib/ictReverseCalculation.ts'));
const plain = value => JSON.parse(JSON.stringify(value));
const compile = code => ts.transpileModule(code, { compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.CommonJS } }).outputText;
const baseStructure = {
  side: 'revenue', sideLabel: '收入', totalInclAmount: 100,
  targetSubject: { groupId: 'revIt', key: 'integration', subjectCode: 'rev_it_integration' },
  balancingSubject: { groupId: 'revIt', key: 'maintenance', subjectCode: 'rev_it_maintenance' },
  targetDisplayName: 'IT集成收入', balancingDisplayName: 'IT维护收入',
  beforeTargetInclAmount: 40, beforeBalancingInclAmount: 60,
  fixedOtherInclAmount: 0, reallocatablePoolInclAmount: 100,
};
const linear = amount => amount / 400;
const jump = amount => amount < 50 ? 0.094 : 0.108;
const evidence = { kind: 'Production function control-flow tests with synthetic candidates/engine; no real project writes', failures: [], successes: [] };

async function run(options = {}, before = false) {
  const structure = { ...plain(baseStructure), ...options.structure };
  const originalTarget = options.target ?? 0.1;
  const selected = { ref: { side: structure.side, ...structure.targetSubject } };
  const events = [], alerts = [], evaluated = [];
  const amounts = { target: structure.beforeTargetInclAmount, balancing: structure.beforeBalancingInclAmount };
  const initialAmounts = { ...amounts };
  const sampleCount = buildLockedTotalStructureSamplePoints(structure.reallocatablePoolInclAmount, structure.beforeTargetInclAmount).length;
  let candidateCount = 0;
  const batch = mock.fn(updates => {
    events.push('write');
    amounts.target = updates[0].incl;
    amounts.balancing = updates[1].incl;
    options.batchImplementation?.(updates);
  });
  const setTarget = mock.fn();
  const sideEffects = [];
  const context = {
    isTaxInclAutoFixEnabled: () => options.autoFix ?? process.argv.includes('--auto-fix'),
    revTargetType: options.metricType || 'margin', revTargetValue: originalTarget,
    state: {
      ignoredDataHash: 'preserved-until-success',
      setIgnoredDataHash: value => sideEffects.push(['ignoredHash', value]),
      setIgnoredTailValue: value => sideEffects.push(['ignoredTail', value]),
      updateTaxItemsInclBatch: batch,
      setCashflowSegments: () => assert.fail('model_e must stay disabled'),
      setActiveTab: value => sideEffects.push(['activeTab', value]),
      ...options.stateOverrides,
    },
    setRevTargetValue: setTarget,
    buildLockedTotalStructureSamplePoints,
    buildLockedTotalStructureCandidate: (s, amount, autoFix) => {
      const index = ++candidateCount;
      const candidate = options.buildCandidate ? options.buildCandidate(s, amount, autoFix, index, sampleCount) : {
        valid: true, targetAmount: amount, balancingAmount: s.reallocatablePoolInclAmount - amount,
        payload: { amount, index }, nextSegments: [], modelETransfers: [],
      };
      return options.candidate ? options.candidate(candidate, index, sampleCount) : candidate;
    },
    invoke: async (command, { input }) => {
      assert.equal(command, 'calculate_ict_benefit');
      events.push('evaluate');
      if (options.engineFailure?.(input.index, sampleCount)) throw new Error('engine failure');
      const metric = (options.metric || linear)(input.amount, input.index, sampleCount);
      evaluated.push({ amount: input.amount, metric });
      if (options.emptyResult?.(input.index, sampleCount)) return 0;
      return { margin_rate: metric, npv_rate: metric, cashflow: [{ year: 1, metric }] };
    },
    setCashflowTable: value => sideEffects.push(['cashflow', plain(value)]),
    setMetrics: value => sideEffects.push(['metrics', plain(value)]),
    updateData: (key, value) => sideEffects.push(['context', key, plain(value)]),
    AI_CONTEXT_KEY: { ICT_CORE: 'ict_core' },
    buildAiContextPayload: (enabled, value) => {
      events.push('prepare-context');
      if (options.contextFailure) throw new Error('context preparation failure');
      return { enabled, ...value };
    },
    alert: message => alerts.push(message),
  };
  const activeDeclarations = before ? { ...declarations, ...baseline } : { ...declarations, ...options.declarationsOverride };
  vm.createContext(context);
  vm.runInContext(compile(Object.entries(activeDeclarations).map(([name, value]) => `const ${name} = ${value};`).join('\n') + '\nglobalThis.run = performReverseCalculation;'), context);
  const reverseContext = options.blocked
    ? { mode: 'blocked', message: '测试：承接规则阻止反算' }
    : { mode: 'locked_total_structure', structure };
  const response = await context.run(options.noSubject ? null : selected, reverseContext, options.cardOptions);
  assert.equal(setTarget.mock.callCount(), 0, 'never rewrite the target via setter');
  assert.equal(context.revTargetValue, originalTarget, 'never assign a new target');
  return { alerts, batch, amounts, initialAmounts, sideEffects, events, evaluated, structure, response };
}

async function failure(name, options, expected, extra) {
  const actual = await run(options);
  assert.equal(actual.batch.mock.callCount(), 0, `${name}: updateTaxItemsInclBatch MUST have zero calls`);
  assert.deepEqual(actual.amounts, actual.initialAmounts, `${name}: amounts unchanged`);
  assert.deepEqual(actual.sideEffects, [], `${name}: no result/context/ignore-flag mutation`);
  assert.equal(actual.alerts.length, 1, name);
  assert.doesNotMatch(actual.alerts[0], /结构反算完成|已为你调整/);
  assert.match(actual.alerts[0], expected, name);
  if (extra) extra(actual);
  const frozenCard = await run({ ...options, declarationsOverride: { performLockedTotalStructureReverseCalculation: beforeCardFunction } });
  assert.deepEqual(actual.alerts, frozenCard.alerts, `${name}: exact pre-card error preserved`);
  const card = await run({ ...options, cardOptions: { silent: true } });
  assert.equal(card.batch.mock.callCount(), 0, `${name}: card also MUST make zero writes`);
  assert.deepEqual(card.alerts, [], 'card does not open a native alert');
  assert.equal(card.response.status, 'error');
  assert.equal(card.response.message, actual.alerts[0], `${name}: native error returned verbatim to card`);
  evidence.failures.push({ name, updateTaxItemsInclBatchCalls: actual.batch.mock.callCount(), targetUnchanged: true, message: actual.alerts[0] });
}
async function success(name, options, expectedAmount, compareBefore = true) {
  const actual = await run(options);
  assert.equal(actual.batch.mock.callCount(), 1, name);
  assert.equal(actual.alerts.length, 1, name);
  assert.match(actual.alerts[0], /^结构反算完成：/);
  const writes = plain(actual.batch.mock.calls[0].arguments[0]);
  assert.equal(writes[0].incl, expectedAmount, name);
  assert.equal(writes[0].incl + writes[1].incl + actual.structure.fixedOtherInclAmount, actual.structure.totalInclAmount);
  assert.ok(Math.abs(actual.evaluated.at(-1).metric - (options.target ?? 0.1)) <= 0.0001);
  assert.ok(actual.events.indexOf('prepare-context') < actual.events.indexOf('write'));
  assert.equal(actual.events.at(-1), 'write', 'all engine evaluations and preparation precede amount write');
  if (compareBefore) {
    const old = await run(options, true);
    assert.equal(old.batch.mock.callCount(), 1);
    assert.deepEqual(writes, plain(old.batch.mock.calls[0].arguments[0]), `${name}: same exact before/after amounts and reasons`);
    assert.deepEqual(actual.alerts, old.alerts, `${name}: full success message unchanged`);
    assert.deepEqual(actual.sideEffects, old.sideEffects, `${name}: same final context, funding entry and displayed results`);
  }
  evidence.successes.push({ name, updateTaxItemsInclBatchCalls: 1, comparedWithPreFix: compareBefore, writes, message: actual.alerts[0] });
}

async function main() {
  const originalBug = await run({ metric: jump, structure: { beforeTargetInclAmount: 30, beforeBalancingInclAmount: 70 } }, true);
  assert.equal(originalBug.batch.mock.callCount(), 1, 'pre-fix witness MUST reproduce an actual write call');
  assert.match(originalBug.alerts[0], /结构反算完成/);
  assert.match(originalBug.alerts[0], /目标 10.00%，当前 9.40%/);
  assert.notDeepEqual(originalBug.amounts, originalBug.initialAmounts, 'pre-fix witness changes actual mocked amounts');
  evidence.preFixWitness = { updateTaxItemsInclBatchCalls: 1, message: originalBug.alerts[0], writes: plain(originalBug.batch.mock.calls[0].arguments[0]) };

  await failure('invalid target', { target: 'invalid' }, /有效的目标值/);
  await failure('missing subject', { noSubject: true }, /请选择/);
  await failure('blocked reverse context', { blocked: true }, /承接规则阻止/);
  await failure('same target and balancing subject', { structure: { balancingSubject: baseStructure.targetSubject } }, /不能.*相同/);
  await failure('negative reallocatable pool', { structure: { reallocatablePoolInclAmount: -1 } }, /锁定总额小于/);
  await failure('all candidates invalid with reason', { candidate: c => ({ ...c, valid: false, message: '候选校验失败' }) }, /候选校验失败/);
  await failure('all candidates invalid fallback', { candidate: c => ({ ...c, valid: false }) }, /无法支持任何/);
  await failure('all candidate payloads missing', { candidate: c => ({ ...c, payload: null }) }, /无法支持任何/);
  await failure('all engine results empty', { emptyResult: () => true }, /无法支持任何/);
  await failure('candidate preparation exception', { candidate: () => { throw new Error('candidate preparation failure'); } }, /candidate preparation failure/);
  await failure('sample engine exception', { engineFailure: () => true }, /结构反算失败: Error: engine failure/);
  await failure('insensitive range', { metric: () => 0.094 }, /不敏感.*目标 10.00%.*最接近值 9.40%.*9.40% - 9.40%/);
  await failure('target below sample range', { metric: jump, target: 0.09 }, /无法达到.*目标 9.00%.*最接近值 9.40%.*9.40% - 10.80%/);
  await failure('target above sample range', { metric: jump, target: 0.12 }, /无法达到.*目标 12.00%.*最接近值 10.80%.*9.40% - 10.80%/);
  for (const metricType of ['margin', 'npv_rate']) {
    for (const increasing of [true, false]) {
      await failure(`${metricType} discontinuous ${increasing ? 'increasing' : 'decreasing'}`, {
        metricType, metric: a => jump(increasing ? a : 100 - a),
      }, /没有找到.*目标 10.00%.*最接近值 9.40%.*9.40% - 10.80%/);
    }
  }
  await failure('search engine exception', { metric: jump, engineFailure: (n, samples) => n > samples }, /engine failure/);
  await failure('search candidate invalid', { metric: jump, candidate: (c, n, samples) => ({ ...c, valid: n <= samples }) }, /没有找到.*9.40% - 10.80%/);
  await failure('search candidate payload missing', { metric: jump, candidate: (c, n, samples) => ({ ...c, payload: n <= samples ? c.payload : null }) }, /没有找到/);
  await failure('search engine result empty', { metric: jump, emptyResult: (n, samples) => n > samples }, /没有找到/);
  for (const [name, change, expected] of [
    ['invalid', c => ({ ...c, valid: false, message: '最终候选拒绝' }), /最终候选拒绝/],
    ['invalid fallback', c => ({ ...c, valid: false }), /最终结构反算候选/],
    ['missing payload', c => ({ ...c, payload: null }), /最终结构反算候选/],
    ['wrong total', c => ({ ...c, balancingAmount: c.balancingAmount + 0.01 }), /未能保持同侧/],
    ['negative target', c => ({ ...c, targetAmount: -0.01, balancingAmount: 100.01 }), /出现负金额/],
    ['negative balancing', c => ({ ...c, targetAmount: 100.01, balancingAmount: -0.01 }), /出现负金额/],
  ]) {
    await failure(`final candidate ${name}`, { candidate: (c, n, samples) => n > samples ? change(c) : c }, expected);
  }
  await failure('final engine exception', { engineFailure: (n, samples) => n > samples }, /engine failure/);
  await failure('final engine result empty', { emptyResult: (n, samples) => n > samples }, /最终结构反算候选/);
  await failure('final recomputation off target', { metric: (a, n, samples) => n > samples ? 0.108 : linear(a) }, /复算值 10.80% 未达到目标.*目标 10.00%.*最接近值 10.00%.*0.00% - 25.00%/);
  await failure('NPV final recomputation just outside tolerance', { metricType: 'npv_rate', metric: (a, n, samples) => n > samples ? 0.10010001 : linear(a) }, /复算值.*未达到目标.*目标净现值率/);
  await failure('result context preparation exception', { contextFailure: true }, /context preparation failure/);
  await failure('metric tolerance cannot be widened', { metric: a => a < 50 ? 0.1 - 0.00010001 : 0.1 + 0.0002 }, /没有找到/);
  // Exercise the defensive collapse branch by collapsing candidate amounts in the stub,
  // not by changing production sample amounts, rounding or either tolerance.
  await failure('collapsed candidate interval keeps nearest metric', {
    metric: jump, candidate: c => ({ ...c, targetAmount: 50, balancingAmount: 50 }),
  }, /最接近值 9.40%/, result => assert.equal(result.evaluated.at(-1).metric, 0.108));

  // Inspect the actual production loop's retained best, independently of diagnostic tracking.
  let loop;
  function findLoop(node) {
    if (ts.isForStatement(node) && node.initializer?.getText(ast) === 'let step = 0' && node.condition?.getText(ast) === 'step < 45') loop = node.getText(ast);
    ts.forEachChild(node, findLoop);
  }
  findLoop(ast);
  assert.ok(loop);
  const loopContext = { METRIC_EPSILON: 0.0001, MONEY_EPSILON: 0.004, target: 0.1,
    low: { targetAmount: 50, metricValue: 0.094 }, high: { targetAmount: 50, metricValue: 0.108 },
    best: { targetAmount: 50, metricValue: 0.094 }, increasing: true,
    roundMoney: n => Number(n.toFixed(2)),
    evaluate: async () => ({ valid: true, payload: {}, result: {}, targetAmount: 50, metricValue: 0.108 }),
  };
  await vm.runInNewContext(compile(`(async () => { ${loop} })()`), loopContext);
  assert.equal(loopContext.best.metricValue, 0.094, 'collapse must not overwrite best with worse last midpoint');
  evidence.collapseBest = loopContext.best.metricValue;

  await success('linear increasing exact sample', {}, 40);
  await success('linear decreasing exact sample', { metric: a => linear(100 - a) }, 60);
  await success('linear interior search', { target: 0.10325 }, 41.33);
  await success('NPV rate interior search', { target: 0.10325, metricType: 'npv_rate' }, 41.33);
  await success('minimum amount change among multiple roots', { metric: a => Math.abs(a - 50) / 100, target: 0.2 }, 30);
  await success('minimum sample endpoint', { metric: jump, target: 0.094 }, 40);
  await success('maximum sample endpoint', { metric: jump, target: 0.108 }, 50);
  await success('zero amount endpoint', { target: 0 }, 0);
  await success('whole pool endpoint', { target: 0.25 }, 100);
  await success('nonzero fixed same-side amount', { structure: { fixedOtherInclAmount: 25, totalInclAmount: 125 } }, 40);
  await success('cost-side amount and context', { structure: {
    side: 'cost', sideLabel: '投入',
    targetSubject: { groupId: 'costIt', key: 'integration', subjectCode: 'cost_it_integration' },
    balancingSubject: { groupId: 'costIt', key: 'maintenance', subjectCode: 'cost_it_maintenance' },
    targetDisplayName: 'IT集成投入', balancingDisplayName: 'IT维护投入',
  } }, 40);
  await success('within unchanged metric tolerance', { metric: a => a < 50 ? 0.1 - 0.00009999 : 0.108 }, 40);
  await success('qualify before minimum amount preference', { metric: a => a < 50 ? 0.094 : a < 80 ? 0.108 : a === 90 ? 0.1 : 0.12 }, 90, false);
  const preview = await run({ metric: jump, cardOptions: { silent: true, previewOnly: true } });
  assert.equal(preview.batch.mock.callCount(), 0);
  assert.deepEqual(preview.sideEffects, []);
  assert.deepEqual(plain(preview.response), { status: 'range', minMetric: 0.094, maxMetric: 0.108 });
  const changed = await run({ cardOptions: { silent: true, beforeCommit: async () => { throw Error('stale input'); } } });
  assert.equal(changed.batch.mock.callCount(), 0);
  assert.match(changed.response.message, /stale input/);
  const cardSuccess = await run({ cardOptions: { silent: true } });
  assert.equal(cardSuccess.batch.mock.callCount(), 1);
  assert.equal(cardSuccess.response.target, 0.1);
  assert.equal(cardSuccess.response.achieved, 0.1);
  assert.equal(cardSuccess.response.targetReached, true);
  evidence.cardAdapter = { nativeErrorsReturnedVerbatim: evidence.failures.length, previewWrites: 0, staleWrites: 0, successWrites: 1 };
  const output = path.resolve(root, `../docs/verification/structure-reverse-success-check${process.argv.includes('--auto-fix') ? '-autofix' : ''}-evidence.json`);
  fs.writeFileSync(output, JSON.stringify(evidence, null, 2) + '\n');
  console.log(`${evidence.failures.length} failure scenarios: updateTaxItemsInclBatch spy count = 0 EACH; target unchanged.`);
  console.log(`${evidence.successes.length} success scenarios; pre-fix exact comparisons, endpoints, minimum change, collapse and unchanged epsilon passed.`);
}
if (require.main === module) main().catch(error => { console.error(error); process.exitCode = 1; });
module.exports = { run };
