const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const ts = require("typescript");

const sourcePath = path.join(__dirname, "../src/lib/presetFieldKeys.ts");
const source = fs.readFileSync(sourcePath, "utf8");
const quickFillSource = fs.readFileSync(
  path.join(__dirname, "../src/components/common-presets/CommonPresetQuickFill.tsx"),
  "utf8",
);
const iconMapSource = fs.readFileSync(
  path.join(__dirname, "../src/components/icons/iconMap.ts"),
  "utf8",
);
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
  console: { warn() {} },
}, { filename: sourcePath });

const {
  PRESET_FIELD_KEYS,
  PRESET_FIELD_DEFINITIONS,
  PRESET_FIELD_REGISTRY,
  getPresetFieldDefinition,
  getPresetFieldDisplay,
  getPresetFieldCategories,
  isPresetFieldEligible,
  presetAppliesToField,
} = moduleRef.exports;

assert.equal(PRESET_FIELD_KEYS.projectBackground, "project_basic.background");
assert.equal(PRESET_FIELD_KEYS.approvalReviewers, "approval.reviewers");
assert.equal(PRESET_FIELD_KEYS.demandUnit, "demand.unit");
assert.equal(PRESET_FIELD_KEYS.meetingItConstructionContent, "meeting.it_construction_content");
assert.equal(PRESET_FIELD_KEYS.paymentRevenueCollectionMethod, "payment.revenue_collection_method");
assert.equal(PRESET_FIELD_KEYS.projectPropertyRights, "project_basic.property_rights");

const background = getPresetFieldDefinition("project_basic.background");
assert.equal(background.kind, "text_snippet");
assert.equal(background.category, "项目背景");
assert.deepEqual(
  Array.from(background.templates),
  ["ICT生命周期测算", "立项签批表", "会审纪要"],
);
assert.deepEqual(Array.from(background.groups), ["项目概况"]);
assert.equal(background.fieldType, "long_text");
assert.equal(background.presetEligible, true);

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

const incomeTerms = getPresetFieldDefinition("payment.revenue_collection_method");
assert.equal(incomeTerms.label, "收入条款");
assert.deepEqual(Array.from(incomeTerms.templates), ["立项签批表", "会审纪要"]);
assert.deepEqual(Array.from(incomeTerms.groups), ["商务条款"]);
assert.equal(getPresetFieldDefinition("contract.income_terms").fieldKey, incomeTerms.fieldKey);

const propertyRights = getPresetFieldDefinition("project_basic.property_rights");
assert.equal(propertyRights.defaultEnabled, false);
assert.equal(propertyRights.presetEligible, true);

const revenueAmount = getPresetFieldDefinition("finance.revenue_amount");
assert.equal(revenueAmount.fieldType, "amount");
assert.equal(revenueAmount.presetEligible, false);
assert.equal(isPresetFieldEligible("finance.revenue_amount"), false);
assert.equal(isPresetFieldEligible("project_basic.background"), true);
assert.equal(PRESET_FIELD_DEFINITIONS.includes(revenueAmount), false);

const unnamed = getPresetFieldDisplay("legacy.unknown_field");
assert.equal(unnamed.label, "未命名字段");
assert.deepEqual(Array.from(unnamed.templates), ["暂未配置"]);
assert.equal(unnamed.presetEligible, false);

const allFieldKeys = PRESET_FIELD_REGISTRY.map(field => field.fieldKey);
assert.equal(new Set(allFieldKeys).size, allFieldKeys.length);
assert.equal(presetAppliesToField([], "approval.reviewers"), true);
assert.equal(presetAppliesToField(["approval.reviewers"], "approval.reviewers"), true);
assert.equal(presetAppliesToField(["approval.reviewers"], "project_basic.customer_name"), false);

assert.match(quickFillSource, /name="presetLibrary"/);
assert.doesNotMatch(quickFillSource, /name="quickAction"/);
assert.match(quickFillSource, /关闭预设/);
assert.match(quickFillSource, /当前字段内容和常用资料库中的资料都会保留/);
assert.match(iconMapSource, /presetLibrary:\s*Bookmark/);
assert.match(iconMapSource, /more:\s*MoreHorizontal/);

console.log("Common preset field key tests passed.");
