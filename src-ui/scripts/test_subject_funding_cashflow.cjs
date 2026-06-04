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

const row = (side, groupId, key, displayName, subjectAmountIncl, taxRate, isItScope = false) => ({
  subjectRef: { side, groupId, key },
  displayName,
  subjectAmountIncl,
  taxRate,
  isItScope,
});

const planFor = (subjectRow, values, enabled = true) => ({
  id: createSubjectFundingPlanId(subjectRow.subjectRef),
  subjectRef: subjectRow.subjectRef,
  mode: "custom",
  annualInclValues: values,
  enabled,
  source: "manual",
});

const revIt = row("revenue", "revIt", "integration", "系统集成服务收入", 106, 6, true);
const revCt = row("revenue", "revCt", "product", "产品收入", 109, 9, false);
const costIt = row("cost", "costIt", "device", "主要设备", 113, 13, true);
const zeroCost = row("cost", "costMix", "marketing", "融合营销成本", 0, 6, false);
const subjects = [revIt, revCt, costIt, zeroCost];

const validPlans = {
  [createSubjectFundingPlanId(revIt.subjectRef)]: planFor(revIt, annual(106)),
  [createSubjectFundingPlanId(revCt.subjectRef)]: planFor(revCt, annual(109)),
  [createSubjectFundingPlanId(costIt.subjectRef)]: planFor(costIt, annual(56.5, 56.5)),
};

const coverage = validateSubjectFundingPlanCoverage(subjects, validPlans);
assert.equal(coverage.valid, true);
assert.equal(coverage.counts.revenueSubjectCount, 2);
assert.equal(coverage.counts.costSubjectCount, 1);
assert.equal(coverage.counts.revenuePlannedCount, 2);
assert.equal(coverage.counts.costPlannedCount, 1);
assert.equal(coverage.counts.issueCount, 0);

const cashflow = buildAnnualCashflowFromSubjectFundingPlans(subjects, validPlans);
assert.deepEqual(Array.from(cashflow.annualRevenueIncl.slice(0, 3)), [215, 0, 0]);
assert.deepEqual(Array.from(cashflow.annualRevenueExcl.slice(0, 3)), [200, 0, 0]);
assert.deepEqual(Array.from(cashflow.annualCostExcl.slice(0, 3)), [50, 50, 0]);
assert.deepEqual(Array.from(cashflow.annualItRevenueExcl.slice(0, 3)), [100, 0, 0]);
assert.deepEqual(Array.from(cashflow.annualItCostExcl.slice(0, 3)), [50, 50, 0]);
assert.deepEqual(Array.from(cashflow.annualNetExcl.slice(0, 3)), [150, -50, 0]);

const missingPlanCoverage = validateSubjectFundingPlanCoverage(subjects, {
  [createSubjectFundingPlanId(revIt.subjectRef)]: validPlans[createSubjectFundingPlanId(revIt.subjectRef)],
  [createSubjectFundingPlanId(costIt.subjectRef)]: validPlans[createSubjectFundingPlanId(costIt.subjectRef)],
});
assert.equal(missingPlanCoverage.valid, false);
assert.equal(missingPlanCoverage.issues.some(issue => issue.type === "missing_plan"), true);

const disabledCoverage = validateSubjectFundingPlanCoverage(subjects, {
  ...validPlans,
  [createSubjectFundingPlanId(revCt.subjectRef)]: planFor(revCt, annual(109), false),
});
assert.equal(disabledCoverage.valid, false);
assert.equal(disabledCoverage.issues.some(issue => issue.type === "disabled_plan"), true);

const invalidCoverage = validateSubjectFundingPlanCoverage(subjects, {
  ...validPlans,
  [createSubjectFundingPlanId(revIt.subjectRef)]: planFor(revIt, [106]),
  [createSubjectFundingPlanId(revCt.subjectRef)]: planFor(revCt, annual(108)),
  [createSubjectFundingPlanId(costIt.subjectRef)]: planFor(costIt, annual(114, -1)),
  [createSubjectFundingPlanId(zeroCost.subjectRef)]: planFor(zeroCost, annual(1)),
});
assert.equal(invalidCoverage.valid, false);
assert.equal(invalidCoverage.issues.some(issue => issue.type === "invalid_length"), true);
assert.equal(invalidCoverage.issues.some(issue => issue.type === "negative_annual_value"), true);
assert.equal(invalidCoverage.issues.some(issue => issue.type === "amount_mismatch"), true);
assert.equal(invalidCoverage.issues.some(issue => issue.type === "zero_subject_with_nonzero_plan"), true);

assert.equal(normalizeCashflowCalculationSource("subject_funding_plans"), "subject_funding_plans");
assert.equal(normalizeCashflowCalculationSource("legacy_model"), "subject_funding_plans");
assert.equal(normalizeCashflowCalculationSource("unexpected"), "subject_funding_plans");
assert.equal(normalizeCashflowCalculationSource(undefined), "subject_funding_plans");

console.log("Subject funding cashflow tests passed.");
