const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const ts = require("typescript");

const sourcePath = path.join(__dirname, "../src/lib/presetFieldKeys.ts");
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
  PRESET_FIELD_KEYS,
  PRESET_FIELD_DEFINITIONS,
  getPresetFieldDefinition,
  getPresetFieldCategories,
  presetAppliesToField,
} = moduleRef.exports;

assert.equal(PRESET_FIELD_KEYS.projectBackground, "project_basic.background");
assert.equal(PRESET_FIELD_KEYS.approvalReviewers, "approval.reviewers");
assert.equal(PRESET_FIELD_KEYS.demandUnit, "demand.unit");
assert.equal(PRESET_FIELD_KEYS.meetingItConstructionContent, "meeting.it_construction_content");
assert.equal(PRESET_FIELD_KEYS.paymentRevenueCollectionMethod, "payment.revenue_collection_method");

const background = getPresetFieldDefinition("project_basic.background");
assert.equal(background.kind, "text_snippet");
assert.equal(background.category, "项目背景");

const reviewer = getPresetFieldDefinition("approval.reviewers");
assert.equal(reviewer.kind, "short_value");
assert.equal(reviewer.category, "审核人员");

assert.ok(getPresetFieldCategories("short_value").includes("审核人员"));
assert.ok(getPresetFieldCategories("text_snippet").includes("项目方案"));
assert.ok(getPresetFieldCategories("short_value").includes("项目需求单位"));
assert.ok(getPresetFieldCategories("text_snippet").includes("收付款方式"));
assert.equal(getPresetFieldDefinition("demand.deployment_environment").kind, "text_snippet");
assert.equal(getPresetFieldDefinition("meeting.onsite_support").kind, "short_value");
assert.equal(getPresetFieldDefinition("approval.it_service_content").category, "IT服务内容");

const allFieldKeys = PRESET_FIELD_DEFINITIONS.map(field => field.fieldKey);
assert.equal(new Set(allFieldKeys).size, allFieldKeys.length);
assert.equal(presetAppliesToField([], "approval.reviewers"), true);
assert.equal(presetAppliesToField(["approval.reviewers"], "approval.reviewers"), true);
assert.equal(presetAppliesToField(["approval.reviewers"], "project_basic.customer_name"), false);

console.log("Common preset field key tests passed.");
