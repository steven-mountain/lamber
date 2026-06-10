const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const ts = require("typescript");

const sourcePath = path.join(__dirname, "../src/lib/ictReverseSearch.ts");
const source = fs.readFileSync(sourcePath, "utf8");
const transpiled = ts.transpileModule(source, {
  compilerOptions: {
    module: ts.ModuleKind.CommonJS,
    target: ts.ScriptTarget.ES2020,
  },
});

const moduleRef = { exports: {} };
vm.runInNewContext(transpiled.outputText, {
  module: moduleRef,
  exports: moduleRef.exports,
  require,
}, { filename: sourcePath });

const {
  buildCostReverseFeasibilityProbeAmounts,
  selectHighestMetricProbe,
} = moduleRef.exports;

assert.deepEqual(
  Array.from(buildCostReverseFeasibilityProbeAmounts("margin")),
  [0],
);
assert.deepEqual(
  Array.from(buildCostReverseFeasibilityProbeAmounts("npv_rate")),
  [0, 0.01],
);

const zeroOutflowBoundary = selectHighestMetricProbe([
  { amount: 0, metricValue: 0 },
  { amount: 0.01, metricValue: 7358490 },
]);
assert.equal(zeroOutflowBoundary.amount, 0.01);
assert.ok(zeroOutflowBoundary.metricValue > 0.1);

const ordinaryBoundary = selectHighestMetricProbe([
  { amount: 0, metricValue: 1.25 },
  { amount: 0.01, metricValue: 1.2499 },
]);
assert.equal(ordinaryBoundary.amount, 0);

console.log("ICT reverse search tests passed.");
