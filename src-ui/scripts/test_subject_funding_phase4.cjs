const assert = require("assert/strict");
const fs = require("fs");
const path = require("path");
const ts = require("typescript");

const subjectPlanPath = path.resolve(__dirname, "../src/lib/ictSubjectFundingPlan.ts");
const code = fs.readFileSync(subjectPlanPath, "utf-8");

const jsCode = ts.transpileModule(code, { compilerOptions: { module: ts.ModuleKind.CommonJS } }).outputText;

const fakeModule = `
${jsCode}

Object.assign(exports, {
  syncSubjectFundingPlanToAmount,
  syncSubjectFundingPlansToAmounts,
  initializeMissingSubjectFundingPlans,
  buildAnnualCashflowSubjectContributions,
  updateSubjectFundingPlanMode,
  updateSubjectFundingPlanAnnualValue,
  createDefaultSubjectFundingPlan: (ref, amount) => ({
    id: exports.createSubjectFundingPlanId(ref),
    subjectRef: ref,
    mode: "upfront",
    annualInclValues: exports.buildUpfrontAnnualInclValues(amount),
    enabled: true,
    source: "manual"
  })
});
`;

const exportsObj = {};
require("vm").runInNewContext(fakeModule, { exports: exportsObj, Math, Number, Array, Date, console, require });

const {
  syncSubjectFundingPlanToAmount,
  initializeMissingSubjectFundingPlans,
  buildAnnualCashflowSubjectContributions,
  updateSubjectFundingPlanMode,
  updateSubjectFundingPlanAnnualValue,
  createDefaultSubjectFundingPlan
} = exportsObj;

const revRef = { side: "revenue", groupId: "revIt", key: "integration" };
const costRef = { side: "cost", groupId: "costIt", key: "hardware" };

console.log("=== Running Phase 4 Tests ===");

// 1. Zero amount reset
{
  const plan = createDefaultSubjectFundingPlan(revRef, 1000);
  const p1 = updateSubjectFundingPlanAnnualValue(plan, 0, 300);
  const p2 = updateSubjectFundingPlanAnnualValue(p1, 1, 700);
  const plans = { [p2.id]: p2 };

  const zeroResult = syncSubjectFundingPlanToAmount(plans, revRef, 0);
  const zeroPlan = zeroResult[p2.id];
  assert.equal(zeroPlan, undefined);

  const restoredResult = syncSubjectFundingPlanToAmount(zeroResult, revRef, 1200);
  const restoredPlan = restoredResult[p2.id];
  assert.equal(restoredPlan.enabled, true);
  assert.equal(restoredPlan.mode, "upfront");
  assert.equal(restoredPlan.annualInclValues[0], 1200);
  assert.equal(restoredPlan.annualInclValues[1], 0);
  console.log("  ✓ Zero amount reset (returns to unmaintained, restores as upfront)");
}

// 2. Sync reasons tracking
{
  const plan = createDefaultSubjectFundingPlan(revRef, 1000);
  const p1 = updateSubjectFundingPlanMode(plan, 1000, "equal", 2);
  assert.equal(p1.lastChangeReason, "manual_plan_edit");

  const p2 = syncSubjectFundingPlanToAmount({ [p1.id]: p1 }, revRef, 2000, "ct_linkage_sync");
  assert.equal(p2[p1.id].lastChangeReason, "ct_linkage_sync");
  console.log("  ✓ Sync reason tracking");
}

// 3. Batch initialization
{
  const plans = {};
  const subjects = [
    { subjectRef: revRef, amountIncl: 1000 },
    { subjectRef: costRef, amountIncl: 0 } // should be skipped
  ];
  const result = initializeMissingSubjectFundingPlans(plans, subjects);
  const revId = [revRef.side, revRef.groupId, revRef.key].join(":");
  const costId = [costRef.side, costRef.groupId, costRef.key].join(":");
  assert.ok(result[revId]);
  assert.equal(result[revId].lastChangeReason, "auto_created_upfront");
  assert.equal(result[revId].annualInclValues[0], 1000);
  assert.equal(result[costId], undefined);
  console.log("  ✓ Batch initialization (skips zero amount)");
}

// 4. Annual drill-down generator
{
  const pRev = createDefaultSubjectFundingPlan(revRef, 1060);
  const pCost = createDefaultSubjectFundingPlan(costRef, 500); // 500 upfront
  const plans = { [pRev.id]: pRev, [pCost.id]: pCost };

  const subjects = [
    { subjectRef: revRef, displayName: "Rev IT", subjectAmountIncl: 1060, taxRate: 6 },
    { subjectRef: costRef, displayName: "Cost HW", subjectAmountIncl: 500, taxRate: 0 }
  ];

  const drilldown = buildAnnualCashflowSubjectContributions(plans, subjects);
  assert.equal(drilldown.length, 10);
  // Year 0
  assert.equal(drilldown[0].length, 2);
  assert.equal(drilldown[0][0].annualInclAmount, 1060);
  assert.equal(drilldown[0][0].annualExclAmount, 1000);
  assert.equal(drilldown[0][1].annualInclAmount, 500);
  assert.equal(drilldown[0][1].annualExclAmount, 500);
  
  console.log("  ✓ Annual cashflow drill-down generator");
}

console.log("All Phase 4 tests passed.");
