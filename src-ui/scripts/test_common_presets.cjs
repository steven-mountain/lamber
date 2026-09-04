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
assert.equal(PRESET_FIELD_KEYS.approvalBranchAttendees, "approval.branch_attendees");
assert.equal(PRESET_FIELD_KEYS.demandUnit, "demand.unit");
assert.equal(PRESET_FIELD_KEYS.meetingItConstructionContent, "meeting.it_construction_content");
assert.equal(PRESET_FIELD_KEYS.paymentRevenueCollectionMethod, "payment.revenue_collection_method");
assert.equal(PRESET_FIELD_KEYS.projectPropertyRights, "project_basic.property_rights");
assert.equal(
  PRESET_FIELD_KEYS.implementationConstructionInterface,
  "implementation.construction_interface",
);

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

const constructionInterface = getPresetFieldDefinition(
  "implementation.construction_interface",
);
assert.equal(constructionInterface.fieldType, "long_text");
assert.equal(constructionInterface.presetEligible, true);
assert.equal(constructionInterface.dictionaryKey, null);

const procurementMethod = getPresetFieldDefinition("procurement.method");
assert.equal(procurementMethod.fieldType, "select");
assert.equal(procurementMethod.presetEligible, false);
assert.equal(procurementMethod.dictionaryKey, "procurement_method");
assert.equal(PRESET_FIELD_DEFINITIONS.includes(procurementMethod), false);

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
assert.doesNotMatch(quickFillSource, /moreOpen/);
assert.match(quickFillSource, /document\.addEventListener\("pointerdown"/);
assert.match(quickFillSource, /title="关闭预设"/);
assert.match(quickFillSource, /关闭预设/);
assert.match(quickFillSource, /当前字段内容和常用资料库中的资料都会保留/);
assert.match(quickFillSource, /选择已有常用内容/);
assert.match(quickFillSource, /暂无可用于该字段的常用内容/);
assert.match(quickFillSource, /保存当前内容为第一条常用内容/);
assert.match(quickFillSource, /最近使用/);
assert.doesNotMatch(quickFillSource, /绑定字段：/);
assert.match(quickFillSource, /type PresetPanelView = "select" \| "save" \| "edit"/);
assert.match(quickFillSource, /setPanelView\("save"\)/);
assert.match(quickFillSource, /setPanelView\("edit"\)/);
assert.match(quickFillSource, /panelView === "select"/);
assert.match(quickFillSource, /panelView === "save"/);
assert.match(quickFillSource, /panelView === "edit"/);
assert.doesNotMatch(quickFillSource, /createPortal/);
assert.doesNotMatch(quickFillSource, /fixed inset-0/);
assert.doesNotMatch(quickFillSource, /role="dialog"/);
assert.match(quickFillSource, /将保存的内容/);
assert.match(quickFillSource, /常用名称/);
assert.match(quickFillSource, /所属模板：/);
assert.match(quickFillSource, /标签（可选）/);
assert.match(quickFillSource, /没有合适的常用内容？/);
assert.match(quickFillSource, /buildDefaultPresetName/);
assert.doesNotMatch(quickFillSource, /setSaveOpen\(current => !current\)/);
assert.match(quickFillSource, /kind === "text_snippet" \? "\\n" : " "/);
assert.match(quickFillSource, /commonPresetService\.delete\(preset\.id\)/);
assert.match(quickFillSource, /删除只影响常用资料库，当前字段内容不会被清空/);
assert.match(quickFillSource, /await loadPresets\(\)/);
assert.match(quickFillSource, />\s*替换\s*</);
assert.doesNotMatch(quickFillSource, />\s*插入\s*</);
assert.match(quickFillSource, />\s*追加\s*</);
assert.match(quickFillSource, />\s*编辑\s*</);
assert.match(quickFillSource, /编辑常用内容/);
assert.match(quickFillSource, /预设名称/);
assert.match(quickFillSource, /预设内容/);
assert.match(quickFillSource, /id: editingPreset\.id/);
assert.match(quickFillSource, /applicableFieldKeys: editingPreset\.applicableFieldKeys/);
assert.match(quickFillSource, /await commonPresetService\.save\(\{/);
assert.match(quickFillSource, /await loadPresets\(\);\s*setPanelView\("select"\)/);
assert.match(quickFillSource, /await loadPresets\(\);\s*setEditingPreset\(null\);\s*setPanelView\("select"\)/);
assert.match(quickFillSource, /role="button"/);
assert.match(quickFillSource, /onClick=\{\(\) => void applyPreset\(preset, "replace"\)\}/);
assert.match(quickFillSource, /event\.stopPropagation\(\)/);
assert.match(quickFillSource, /bg-destructive-soft/);
assert.doesNotMatch(quickFillSource, /presetMenuId/);
assert.doesNotMatch(quickFillSource, /更多操作：/);
assert.match(iconMapSource, /presetLibrary:\s*Bookmark/);
assert.match(iconMapSource, /more:\s*MoreHorizontal/);

console.log("Common preset field key tests passed.");
