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
  SUBJECT_FUNDING_PLAN_MIGRATION_VERSION,
  buildAnnualCashflowFromSubjectFundingPlans,
  createSubjectFundingPlanId,
  migrateLegacySubjectFundingPlans,
  validateSubjectFundingPlanCoverage,
} = moduleRef.exports;

const annual = (...values) => Array.from({ length: 10 }, (_, index) => values[index] || 0);
const row = (side, groupId, key, amount, taxRate, isItScope = false) => ({
  subjectRef: { side, groupId, key },
  displayName: `${side}-${groupId}-${key}`,
  subjectAmountIncl: amount,
  taxRate,
  isItScope,
});
const planFor = (subject, values, mode = "custom") => ({
  id: createSubjectFundingPlanId(subject.subjectRef),
  subjectRef: subject.subjectRef,
  mode,
  annualInclValues: values,
  enabled: true,
  source: "manual",
});

const rev = row("revenue", "revIt", "integration", 106, 6, true);
const cost = row("cost", "costIt", "device", 113, 13, true);
const zero = row("cost", "costMix", "marketing", 0, 6, false);
const subjects = [rev, cost, zero];

{
  const migrated = migrateLegacySubjectFundingPlans(subjects, {}, undefined);
  assert.equal(migrated.changed, true);
  assert.equal(migrated.completed, true);
  assert.equal(migrated.migrationVersion, SUBJECT_FUNDING_PLAN_MIGRATION_VERSION);

  const revPlan = migrated.plans[createSubjectFundingPlanId(rev.subjectRef)];
  const costPlan = migrated.plans[createSubjectFundingPlanId(cost.subjectRef)];
  const zeroPlan = migrated.plans[createSubjectFundingPlanId(zero.subjectRef)];
  assert.equal(revPlan.source, "migration");
  assert.equal(costPlan.source, "migration");
  assert.deepEqual(Array.from(revPlan.annualInclValues), annual(106));
  assert.deepEqual(Array.from(costPlan.annualInclValues), annual(113));
  assert.equal(zeroPlan, undefined);
}

{
  const existing = planFor(rev, annual(53, 53), "equal");
  const migrated = migrateLegacySubjectFundingPlans(subjects, { [existing.id]: existing }, undefined);
  assert.equal(migrated.changed, true);
  assert.deepEqual(Array.from(migrated.plans[existing.id].annualInclValues), annual(53, 53));
  assert.equal(migrated.plans[existing.id].mode, "equal");
  assert.ok(migrated.plans[createSubjectFundingPlanId(cost.subjectRef)]);
}

{
  const bad = planFor(rev, annual(10));
  const migrated = migrateLegacySubjectFundingPlans(subjects, { [bad.id]: bad }, undefined);
  assert.equal(migrated.completed, false);
  assert.equal(migrated.migrationVersion, undefined);
  assert.equal(migrated.coverage.valid, false);
  assert.equal(migrated.plans[bad.id].annualInclValues[0], 10);
  assert.ok(migrated.plans[createSubjectFundingPlanId(cost.subjectRef)]);
}

{
  const alreadyMigrated = migrateLegacySubjectFundingPlans(subjects, {}, SUBJECT_FUNDING_PLAN_MIGRATION_VERSION);
  assert.equal(alreadyMigrated.changed, false);
  assert.equal(alreadyMigrated.completed, true);
  assert.equal(Object.keys(alreadyMigrated.plans).length, 0);
}

{
  const migrated = migrateLegacySubjectFundingPlans(subjects, {}, undefined);
  const coverage = validateSubjectFundingPlanCoverage(subjects, migrated.plans);
  assert.equal(coverage.valid, true);
  const cashflow = buildAnnualCashflowFromSubjectFundingPlans(subjects, migrated.plans);
  assert.deepEqual(Array.from(cashflow.annualRevenueExcl.slice(0, 2)), [100, 0]);
  assert.deepEqual(Array.from(cashflow.annualCostExcl.slice(0, 2)), [100, 0]);
  assert.deepEqual(Array.from(cashflow.annualItRevenueExcl.slice(0, 2)), [100, 0]);
  assert.deepEqual(Array.from(cashflow.annualItCostExcl.slice(0, 2)), [100, 0]);
}

console.log("Subject funding migration tests passed.");
