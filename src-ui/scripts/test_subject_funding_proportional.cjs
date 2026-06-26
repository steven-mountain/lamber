const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const ts = require("typescript");

function load(p) {
  const source = fs.readFileSync(p, "utf8");
  const transpiled = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
    },
  });
  const moduleRef = { exports: {} };
  vm.runInNewContext(
    transpiled.outputText,
    { module: moduleRef, exports: moduleRef.exports, require },
    { filename: p },
  );
  return moduleRef.exports;
}

const fp = load(path.join(__dirname, "../src/lib/ictSubjectFundingPlan.ts"));

const {
  buildProportionalAnnualInclValues,
  normalizeAnnualPercentages,
  sumAnnualPercentages,
  isAnnualPercentagesComplete,
  updateSubjectFundingPlanMode,
  updateSubjectFundingPlanPercentage,
  createDefaultSubjectFundingPlan,
  syncSubjectFundingPlanToAmount,
  buildAnnualCashflowFromSubjectFundingPlans,
  validateSubjectFundingPlan,
  createSubjectFundingPlanId,
} = fp;

const ref = (side, groupId, key) => ({ side, groupId, key });

// 1) Proportional distribution: 95% year1, 5% year6 on a 1000 amount
{
  const pct = Array(10).fill(0);
  pct[0] = 95;
  pct[5] = 5;
  const values = buildProportionalAnnualInclValues(1000, pct);
  assert.equal(values[0], 950, "year1 should be 950");
  assert.equal(values[5], 50, "year6 should be 50");
  const total = values.reduce((s, v) => s + v, 0);
  assert.equal(Math.round(total * 100), 100000, "total must equal subject amount");
  console.log("  ✓ 95/5 split distributes correctly and sums to amount");
}

// 2) Rounding tail folds into last non-zero year (33.33 / 33.33 / 33.34 style)
{
  const pct = [33.3333, 33.3333, 33.3334, 0, 0, 0, 0, 0, 0, 0];
  const values = buildProportionalAnnualInclValues(100, pct);
  const total = values.reduce((s, v) => s + v, 0);
  assert.equal(Math.round(total * 100), 10000, "thirds must still total exactly the amount");
  console.log("  ✓ rounding tail keeps integer-cents total exact");
}

// 3) Percentage helpers
{
  assert.equal(sumAnnualPercentages([95, 0, 0, 0, 0, 5]), 100);
  assert.equal(isAnnualPercentagesComplete([95, 0, 0, 0, 0, 5]), true);
  assert.equal(isAnnualPercentagesComplete([95, 0, 0, 0, 0, 4]), false);
  const norm = normalizeAnnualPercentages([95, -3, "x", 5]);
  assert.equal(norm[0], 95);
  assert.equal(norm[1], 0, "negative clamped to 0");
  assert.equal(norm[2], 0, "non-numeric -> 0");
  console.log("  ✓ percentage helpers (sum / complete / normalize)");
}

// 4) updateSubjectFundingPlanMode -> proportional defaults to 100% year1
{
  const base = createDefaultSubjectFundingPlan(ref("cost", "costIt", "device"), 1000);
  const prop = updateSubjectFundingPlanMode(base, 1000, "proportional", 10);
  assert.equal(prop.mode, "proportional");
  assert.equal(prop.annualPercentages[0], 100);
  assert.equal(prop.annualInclValues[0], 1000);
  console.log("  ✓ switching to proportional defaults to 100% in year 1");
}

// 5) updateSubjectFundingPlanPercentage edits and recomputes amounts
{
  let plan = createDefaultSubjectFundingPlan(ref("revenue", "revIt", "integration"), 1000);
  plan = updateSubjectFundingPlanPercentage(plan, 1000, 0, 95);
  plan = updateSubjectFundingPlanPercentage(plan, 1000, 5, 5);
  assert.equal(plan.mode, "proportional");
  assert.equal(plan.annualInclValues[0], 950);
  assert.equal(plan.annualInclValues[5], 50);
  const v = validateSubjectFundingPlan(plan, 1000);
  assert.equal(v.valid, true, "proportional plan should validate against amount");
  console.log("  ✓ editing percentages recomputes amounts and validates");
}

// 6) Amount change re-derives proportional amounts from stored percentages
{
  let plan = createDefaultSubjectFundingPlan(ref("cost", "costIt", "device"), 1000);
  plan = updateSubjectFundingPlanPercentage(plan, 1000, 0, 95);
  plan = updateSubjectFundingPlanPercentage(plan, 1000, 5, 5);
  const id = createSubjectFundingPlanId(plan.subjectRef);
  const plans = { [id]: plan };
  const synced = syncSubjectFundingPlanToAmount(plans, plan.subjectRef, 2000);
  assert.equal(synced[id].annualInclValues[0], 1900, "95% of 2000");
  assert.equal(synced[id].annualInclValues[5], 100, "5% of 2000");
  console.log("  ✓ amount change keeps percentage distribution");
}

// 7) Unmaintained-subject fallback keeps IT breakdown intact
{
  const subjects = [
    { subjectRef: ref("cost", "costIt", "device"), displayName: "设备", subjectAmountIncl: 1130, taxRate: 13, isItScope: true },
    { subjectRef: ref("cost", "costIt", "integration"), displayName: "集成", subjectAmountIncl: 106, taxRate: 6, isItScope: true },
  ];
  // device has a maintained 2-year plan; integration has NO plan (unmaintained)
  const devId = createSubjectFundingPlanId(subjects[0].subjectRef);
  const plans = {
    [devId]: {
      id: devId,
      subjectRef: subjects[0].subjectRef,
      mode: "custom",
      annualInclValues: [565, 0, 565, 0, 0, 0, 0, 0, 0, 0],
      enabled: true,
      source: "manual",
    },
  };

  // Without fallback: integration is dropped, device spreads
  const noFallback = buildAnnualCashflowFromSubjectFundingPlans(subjects, plans);
  assert.equal(Math.round(noFallback.annualItCostExcl[0] * 100), 50000, "device y1 excl = 500");

  // With fallback: integration (unmaintained) lands upfront in year 1, device still spread
  const withFallback = buildAnnualCashflowFromSubjectFundingPlans(subjects, plans, {
    fallbackUnmaintainedToUpfront: true,
  });
  // y1 = device 500 + integration 100 = 600 ; y3 = device 500
  assert.equal(Math.round(withFallback.annualItCostExcl[0] * 100), 60000, "y1 excl = 600 with fallback");
  assert.equal(Math.round(withFallback.annualItCostExcl[2] * 100), 50000, "y3 excl = 500 (device spread preserved)");
  console.log("  ✓ unmaintained subject falls back to upfront, maintained IT plan still spreads");
}

console.log("All proportional + fallback tests passed.");
