const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const ts = require("typescript");

const sourcePath = path.join(__dirname, "../src/lib/ictSubjectFundingPlan.ts");
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
  buildAnnualCashflowFromSubjectFundingPlans,
  createSubjectFundingPlanId,
  normalizeCashflowCalculationSource,
  validateSubjectFundingPlanCoverage,
} = moduleRef.exports;

const annual = (...values) => Array.from({ length: 10 }, (_, index) => values[index] || 0);
const row = (side, groupId, key, displayName, amount, taxRate, isItScope = false) => ({
  subjectRef: { side, groupId, key },
  displayName,
  subjectAmountIncl: amount,
  taxRate,
  isItScope,
});
const planFor = (subject, values) => ({
  id: createSubjectFundingPlanId(subject.subjectRef),
  subjectRef: subject.subjectRef,
  mode: "custom",
  annualInclValues: values,
  enabled: true,
  source: "manual",
});

assert.equal(normalizeCashflowCalculationSource("legacy_model"), "subject_funding_plans");
assert.equal(normalizeCashflowCalculationSource(null), "subject_funding_plans");

const revIt = row("revenue", "revIt", "integration", "IT集成收入", 1060, 6, true);
const revCt = row("revenue", "revCt", "product", "CT产品收入", 1090, 9, false);
const costIt = row("cost", "costIt", "device", "IT设备投入", 1130, 13, true);
const subjects = [revIt, revCt, costIt];
const plans = {
  [createSubjectFundingPlanId(revIt.subjectRef)]: planFor(revIt, annual(530, 530)),
  [createSubjectFundingPlanId(revCt.subjectRef)]: planFor(revCt, annual(0, 1090)),
  [createSubjectFundingPlanId(costIt.subjectRef)]: planFor(costIt, annual(1130)),
};

const coverage = validateSubjectFundingPlanCoverage(subjects, plans);
assert.equal(coverage.valid, true);

const cashflow = buildAnnualCashflowFromSubjectFundingPlans(subjects, plans);
assert.deepEqual(Array.from(cashflow.annualRevenueIncl.slice(0, 3)), [530, 1620, 0]);
assert.deepEqual(Array.from(cashflow.annualCostIncl.slice(0, 3)), [1130, 0, 0]);
assert.deepEqual(Array.from(cashflow.annualRevenueExcl.slice(0, 3)), [500, 1500, 0]);
assert.deepEqual(Array.from(cashflow.annualCostExcl.slice(0, 3)), [1000, 0, 0]);
assert.deepEqual(Array.from(cashflow.annualItRevenueExcl.slice(0, 3)), [500, 500, 0]);
assert.deepEqual(Array.from(cashflow.annualItCostExcl.slice(0, 3)), [1000, 0, 0]);

const legacySegmentCashflow = {
  rev: annual(999999),
  cost: annual(888888),
  itRev: annual(777777),
  itCost: annual(666666),
};

assert.notDeepEqual(cashflow.annualRevenueExcl, legacySegmentCashflow.rev);
assert.notDeepEqual(cashflow.annualCostExcl, legacySegmentCashflow.cost);
assert.notDeepEqual(cashflow.annualItRevenueExcl, legacySegmentCashflow.itRev);
assert.notDeepEqual(cashflow.annualItCostExcl, legacySegmentCashflow.itCost);

console.log("Subject funding final-source tests passed.");
