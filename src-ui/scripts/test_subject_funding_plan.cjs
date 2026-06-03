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
  buildEqualAnnualInclValues,
  createDefaultSubjectFundingPlan,
  createSubjectFundingPlanId,
  normalizeSubjectFundingPlans,
  removeSubjectFundingPlan,
  updateSubjectFundingPlanAnnualValue,
  upsertSubjectFundingPlan,
  validateSubjectFundingPlan,
} = moduleRef.exports;

const revenueRef = { side: "revenue", groupId: "revIt", key: "integration" };
const costRef = { side: "cost", groupId: "costIt", key: "device" };

assert.equal(createSubjectFundingPlanId(revenueRef), "revenue:revIt:integration");
assert.equal(createSubjectFundingPlanId(costRef), "cost:costIt:device");

const revenuePlan = createDefaultSubjectFundingPlan(revenueRef, 1000000);
assert.equal(revenuePlan.mode, "upfront");
assert.equal(revenuePlan.source, "manual");
assert.equal(revenuePlan.enabled, true);
assert.deepEqual(Array.from(revenuePlan.annualInclValues), [1000000, 0, 0, 0, 0, 0, 0, 0, 0, 0]);

const otherRevenuePlan = createDefaultSubjectFundingPlan({ side: "revenue", groupId: "revIt", key: "maintenance" }, 800);
let plans = upsertSubjectFundingPlan({}, revenuePlan);
plans = upsertSubjectFundingPlan(plans, otherRevenuePlan);
assert.equal(Object.keys(plans).length, 2);
assert.notEqual(revenuePlan.id, otherRevenuePlan.id);

const equalValues = buildEqualAnnualInclValues(1000000, 3);
assert.deepEqual(Array.from(equalValues.slice(0, 4)), [333333.33, 333333.33, 333333.34, 0]);
assert.equal(equalValues.reduce((sum, value) => sum + value, 0), 1000000);

const customPlan = updateSubjectFundingPlanAnnualValue(revenuePlan, 1, 700000);
assert.equal(customPlan.mode, "custom");
assert.equal(customPlan.annualInclValues[1], 700000);

const nonNegativePlan = updateSubjectFundingPlanAnnualValue(revenuePlan, 2, -10);
assert.equal(nonNegativePlan.annualInclValues[2], 0);

const validResult = validateSubjectFundingPlan(revenuePlan, 1000000);
assert.equal(validResult.valid, true);
assert.equal(validResult.subjectAmountIncl, 1000000);
assert.equal(validResult.plannedAmountIncl, 1000000);
assert.equal(validResult.difference, 0);
assert.equal(validateSubjectFundingPlan(customPlan, 1000000).difference, -700000);
assert.equal(validateSubjectFundingPlan(updateSubjectFundingPlanAnnualValue(revenuePlan, 0, 300000), 1000000).difference, 700000);

plans = removeSubjectFundingPlan(plans, revenueRef);
assert.equal(Boolean(plans[revenuePlan.id]), false);
assert.equal(Object.keys(normalizeSubjectFundingPlans(undefined)).length, 0);

console.log("Subject funding plan tests passed.");
