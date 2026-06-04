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
  buildUpfrontAnnualInclValues,
  createDefaultSubjectFundingPlan,
  createSubjectFundingPlanId,
  normalizeAnnualInclValues,
  syncSubjectFundingPlanToAmount,
  syncSubjectFundingPlansToAmounts,
  validateSubjectFundingPlan,
} = moduleRef.exports;

const PLAN_YEARS = 10;
const revRef = { side: "revenue", groupId: "revIt", key: "integration" };
const costRef = { side: "cost", groupId: "costIt", key: "device" };
const revCtProductRef = { side: "revenue", groupId: "revCt", key: "product" };
const costCtOtherRef = { side: "cost", groupId: "costCt", key: "other" };

const sumAnnual = (values) => values.reduce((sum, v) => sum + v, 0);
const roundCents = (v) => Math.round(v * 100) / 100;

// ─────────────────────────────────────────────────────────
// Test 1: Proportional scaling preserves distribution shape
// ─────────────────────────────────────────────────────────
{
  const plan = {
    ...createDefaultSubjectFundingPlan(revRef, 1000),
    mode: "custom",
    annualInclValues: [600, 300, 100, 0, 0, 0, 0, 0, 0, 0],
  };
  const plans = { [plan.id]: plan };

  const result = syncSubjectFundingPlanToAmount(plans, revRef, 2000);
  const synced = result[plan.id];
  assert.equal(synced.annualInclValues.length, PLAN_YEARS);
  assert.equal(synced.annualInclValues[0], 1200, "Year 1 should scale 600 → 1200");
  assert.equal(synced.annualInclValues[1], 600, "Year 2 should scale 300 → 600");
  // Year 3 gets tail correction: 2000 - 1200 - 600 = 200
  assert.equal(synced.annualInclValues[2], 200, "Year 3 should scale 100 → 200 (with tail)");
  assert.equal(synced.mode, "custom", "Mode should be preserved");
  assert.equal(synced.enabled, true, "Plan should be enabled");
  console.log("  ✓ Test 1: Proportional scaling");
}

// ─────────────────────────────────────────────────────────
// Test 2: Tail-difference correction ensures zero tolerance
// ─────────────────────────────────────────────────────────
{
  // Use values that produce rounding drift
  const plan = {
    ...createDefaultSubjectFundingPlan(revRef, 1000),
    mode: "custom",
    annualInclValues: [333.33, 333.33, 333.34, 0, 0, 0, 0, 0, 0, 0],
  };
  const plans = { [plan.id]: plan };

  const result = syncSubjectFundingPlanToAmount(plans, revRef, 777.77);
  const synced = result[plan.id];
  const total = roundCents(sumAnnual(synced.annualInclValues));
  assert.equal(total, 777.77, `Total should be exactly 777.77, got ${total}`);
  console.log("  ✓ Test 2: Tail-difference correction (zero tolerance)");
}

// ─────────────────────────────────────────────────────────
// Test 3: Auto-create when plan is missing and amount > 0
// ─────────────────────────────────────────────────────────
{
  const result = syncSubjectFundingPlanToAmount({}, costRef, 5000);
  const planId = createSubjectFundingPlanId(costRef);
  const synced = result[planId];
  assert.ok(synced, "Plan should be auto-created");
  assert.equal(synced.mode, "upfront", "Auto-created plan should be upfront");
  assert.equal(synced.annualInclValues[0], 5000, "Year 1 should have full amount");
  assert.equal(synced.enabled, true, "Plan should be enabled");
  assert.equal(synced.source, "manual", "Source should be manual");
  const validation = validateSubjectFundingPlan(synced, 5000);
  assert.equal(validation.valid, true, "Validation should pass");
  console.log("  ✓ Test 3: Auto-create upfront plan");
}

// ─────────────────────────────────────────────────────────
// Test 4: Clear to zero removes plan
// ─────────────────────────────────────────────────────────
{
  const plan = createDefaultSubjectFundingPlan(revRef, 1000);
  const plans = { [plan.id]: plan };

  const result = syncSubjectFundingPlanToAmount(plans, revRef, 0);
  const synced = result[plan.id];
  assert.equal(synced, undefined, "Plan should be removed");
  console.log("  ✓ Test 4: Clear to zero removes plan");
}

// ─────────────────────────────────────────────────────────
// Test 5: No-op when amount <= 0 and no plan exists
// ─────────────────────────────────────────────────────────
{
  const result = syncSubjectFundingPlanToAmount({}, revRef, 0);
  assert.equal(Object.keys(result).length, 0, "Should not create a plan for zero amount");
  console.log("  ✓ Test 5: No-op for zero amount without existing plan");
}

// ─────────────────────────────────────────────────────────
// Test 6: Zero-total fallback to upfront
// ─────────────────────────────────────────────────────────
{
  // Simulate a previously disabled plan (all zeros)
  const plan = {
    ...createDefaultSubjectFundingPlan(revRef, 0),
    annualInclValues: Array(PLAN_YEARS).fill(0),
    enabled: false,
  };
  const plans = { [plan.id]: plan };

  const result = syncSubjectFundingPlanToAmount(plans, revRef, 3000);
  const synced = result[plan.id];
  assert.equal(synced.mode, "upfront", "Should fall back to upfront");
  assert.equal(synced.annualInclValues[0], 3000, "Year 1 should have full amount");
  assert.equal(synced.enabled, true, "Plan should be re-enabled");
  console.log("  ✓ Test 6: Zero-total fallback to upfront");
}

// ─────────────────────────────────────────────────────────
// Test 7: Batch sync — multiple subjects, no interference
// ─────────────────────────────────────────────────────────
{
  const revPlan = {
    ...createDefaultSubjectFundingPlan(revRef, 1000),
    mode: "equal",
    annualInclValues: [500, 500, 0, 0, 0, 0, 0, 0, 0, 0],
  };
  const costPlan = createDefaultSubjectFundingPlan(costRef, 2000);
  const plans = { [revPlan.id]: revPlan, [costPlan.id]: costPlan };

  const result = syncSubjectFundingPlansToAmounts(plans, [
    { subjectRef: revRef, newAmountIncl: 2000 },
    { subjectRef: costRef, newAmountIncl: 500 },
  ]);

  const syncedRev = result[revPlan.id];
  const syncedCost = result[costPlan.id];

  assert.equal(roundCents(sumAnnual(syncedRev.annualInclValues)), 2000, "Revenue sum should be 2000");
  assert.equal(syncedRev.annualInclValues[0], 1000, "Revenue year 1 scaled from 500 → 1000");
  assert.equal(syncedRev.annualInclValues[1], 1000, "Revenue year 2 scaled from 500 → 1000");

  assert.equal(roundCents(sumAnnual(syncedCost.annualInclValues)), 500, "Cost sum should be 500");
  assert.equal(syncedCost.annualInclValues[0], 500, "Cost year 1 = 500 (upfront scaled)");
  console.log("  ✓ Test 7: Batch sync — no interference");
}

// ─────────────────────────────────────────────────────────
// Test 8: CT linkage scenario — revCt.product + costCt.other
// ─────────────────────────────────────────────────────────
{
  const revPlan = createDefaultSubjectFundingPlan(revCtProductRef, 1000);
  const costPlan = createDefaultSubjectFundingPlan(costCtOtherRef, 1000);
  const plans = { [revPlan.id]: revPlan, [costPlan.id]: costPlan };

  // Simulate CT linkage: both sides change to same amount
  const result = syncSubjectFundingPlansToAmounts(plans, [
    { subjectRef: revCtProductRef, newAmountIncl: 2500 },
    { subjectRef: costCtOtherRef, newAmountIncl: 2500 },
  ]);

  const syncedRev = result[revPlan.id];
  const syncedCost = result[costPlan.id];

  assert.equal(roundCents(sumAnnual(syncedRev.annualInclValues)), 2500, "Revenue plan should sum to 2500");
  assert.equal(roundCents(sumAnnual(syncedCost.annualInclValues)), 2500, "Cost plan should sum to 2500");
  assert.equal(syncedRev.subjectRef.side, "revenue", "Revenue plan side should be revenue");
  assert.equal(syncedCost.subjectRef.side, "cost", "Cost plan side should be cost");
  console.log("  ✓ Test 8: CT linkage sync");
}

// ─────────────────────────────────────────────────────────
// Test 9: Scaling preserves "equal" mode
// ─────────────────────────────────────────────────────────
{
  const plan = {
    ...createDefaultSubjectFundingPlan(revRef, 1000),
    mode: "equal",
    equalYears: 5,
    annualInclValues: [200, 200, 200, 200, 200, 0, 0, 0, 0, 0],
  };
  const plans = { [plan.id]: plan };

  const result = syncSubjectFundingPlanToAmount(plans, revRef, 3000);
  const synced = result[plan.id];
  assert.equal(synced.mode, "equal", "Mode should remain equal");
  assert.equal(synced.annualInclValues[0], 600, "Year 1 should scale from 200 → 600");
  assert.equal(synced.annualInclValues[4], 600, "Year 5 should scale from 200 → 600");
  assert.equal(roundCents(sumAnnual(synced.annualInclValues)), 3000, "Total should be 3000");
  console.log("  ✓ Test 9: Equal mode preserved");
}

// ─────────────────────────────────────────────────────────
// Test 10: Scaling down (amount decreases)
// ─────────────────────────────────────────────────────────
{
  const plan = {
    ...createDefaultSubjectFundingPlan(revRef, 1000),
    mode: "custom",
    annualInclValues: [600, 200, 200, 0, 0, 0, 0, 0, 0, 0],
  };
  const plans = { [plan.id]: plan };

  const result = syncSubjectFundingPlanToAmount(plans, revRef, 500);
  const synced = result[plan.id];
  assert.equal(roundCents(sumAnnual(synced.annualInclValues)), 500, "Total should be 500");
  assert.equal(synced.annualInclValues[0], 300, "Year 1: 600 * 500/1000 = 300");
  assert.equal(synced.annualInclValues[1], 100, "Year 2: 200 * 500/1000 = 100");
  console.log("  ✓ Test 10: Scale down");
}

// ─────────────────────────────────────────────────────────
// Test 11: Negative amount removes plan
// ─────────────────────────────────────────────────────────
{
  const plan = createDefaultSubjectFundingPlan(revRef, 1000);
  const plans = { [plan.id]: plan };

  const result = syncSubjectFundingPlanToAmount(plans, revRef, -500);
  const synced = result[plan.id];
  assert.equal(synced, undefined, "Plan should be removed for negative amount");
  console.log("  ✓ Test 11: Negative amount removes plan");
}

// ─────────────────────────────────────────────────────────
// Test 12: Immutability — original plans not mutated
// ─────────────────────────────────────────────────────────
{
  const plan = createDefaultSubjectFundingPlan(revRef, 1000);
  const plans = { [plan.id]: plan };
  const originalValues = [...plan.annualInclValues];

  syncSubjectFundingPlanToAmount(plans, revRef, 2000);
  assert.equal(JSON.stringify(plan.annualInclValues), JSON.stringify(originalValues), "Original plan should not be mutated");
  console.log("  ✓ Test 12: Immutability");
}

console.log("\nAll subject funding sync tests passed.");
