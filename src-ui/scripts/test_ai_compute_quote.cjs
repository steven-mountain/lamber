const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const ts = require("typescript");

const moduleCache = new Map();
function loadTsFile(sourcePath) {
  const normalizedPath = path.normalize(sourcePath);
  if (moduleCache.has(normalizedPath)) return moduleCache.get(normalizedPath).exports;
  const source = fs.readFileSync(normalizedPath, "utf8");
  const transpiled = ts.transpileModule(source, {
    compilerOptions: {
      esModuleInterop: true,
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
    },
  });
  const moduleRef = { exports: {} };
  moduleCache.set(normalizedPath, moduleRef);
  const localRequire = request => {
    if (request.startsWith(".")) {
      const resolved = path.resolve(path.dirname(normalizedPath), request);
      return loadTsFile(path.extname(resolved) ? resolved : `${resolved}.ts`);
    }
    return require(request);
  };
  vm.runInNewContext(transpiled.outputText, {
    module: moduleRef,
    exports: moduleRef.exports,
    require: localRequire,
  }, { filename: normalizedPath });
  return moduleRef.exports;
}

function loadTs(relativePath) {
  return loadTsFile(path.join(__dirname, "../src/features/ai-compute-quote", relativePath));
}

const {
  buildAiComputeQuoteOutput,
  buildAiComputeQuoteOutputFundingPlans,
  calculateQuoteBlueprint,
  evaluateQuoteFormula,
  runAiComputeQuoteSensitivity,
  summarizeQuote,
} = loadTs("calculations.ts");
const {
  buildAiComputeEvenAmounts,
  buildAiComputeFirstYearAmounts,
  createDefaultAiComputeFundingPlan,
  getAiComputeDiscountRateDecimal,
  getAiComputeDiscountRatePercent,
  getAiComputeProjectCycleYears,
  ictDiscountRateToAiComputePercent,
  normalizeAiComputeFundingPlan,
  normalizeAiComputeDiscountRatePercent,
  normalizeAiComputeProjectCycleValue,
  sumAiComputeFundingPlan,
  updateAiComputeFundingPlanYear,
  validateAiComputeFundingPlan,
} = loadTs("fundingPlans.ts");
const { createH200Blueprint } = loadTs("presets.ts");
const {
  DEFAULT_PARAMETER_GROUPS,
  PARAMETER_GROUP_IDS,
  canDeleteAiComputeParameterGroup,
  initializeAiComputeDiscountRate,
  migrateAiComputeDiscountRateOwnership,
  moveAiComputeParameter,
  moveAiComputeParameterByOffset,
  normalizeAiComputeParameterLayout,
  reorderAiComputeParameterGroup,
} = loadTs("parameterLayout.ts");
const {
  buildAiComputeAutoSyncPreview,
  buildAiComputeIctExportPayloads,
  buildAiComputeIctExportPreview,
  buildIntelligentComputeAggregatePreview,
  validateIntelligentComputeSources,
} = loadTs("ictExport.ts");
const {
  applySuccessfulAiComputeSync,
  buildIntelligentComputeSyncLock,
  clearAiComputeControlForMappingChange,
  getAiComputeSyncFingerprint,
  projectStateSyncIncludesAmountSource,
  reconcileAiComputeBlueprintFromIct,
  restoreIctResultFromProjectState,
  restoreAiComputeFormulaControl,
} = loadTs("ictSync.ts");
const {
  canDeleteIntelligentAmountSource,
  getDefaultCreateAmountSourceBaseMode,
  isH200BaselineAmountSource,
} = loadTs("amountSources.ts");
const {
  AMOUNT_SOURCE_PACKAGE_KIND,
  buildAmountSourcePackage,
  buildBlueprintFromAmountSourcePackage,
  getDefaultImportedAmountSourceName,
  normalizeAmountSourcePackage,
} = loadTs("amountSourceExchange.ts");
const {
  finalizeIctInputWithFundingPlans,
} = loadTsFile(path.join(__dirname, "../src/lib/ictCalculationInput.ts"));
const {
  ICT_SUBJECT_DEFINITIONS,
} = loadTsFile(path.join(__dirname, "../src/lib/ictSubjectCatalog.ts"));
const {
  clampFormulaCursor,
  insertFormulaTokensAt,
  removeFormulaTokenAt,
  removeFormulaTokenBeforeCursor,
} = loadTs("formulaTokenEditing.ts");
const {
  useNavigationStore,
} = loadTsFile(path.join(__dirname, "../src/store/useNavigationStore.ts"));

const parameters = [
  { id: "a", name: "参数A", key: "a", value: 5 },
  { id: "b", name: "参数B", key: "b", value: 4 },
  { id: "rate", name: "资金成本率", key: "rate", value: 10, unit: "%" },
];
const expression = tokens => ({ version: 2, tokens });
const item = (id, name, formula, enabled = true) => ({
  id,
  side: "cost",
  name,
  formula,
  amountInclTax: 0,
  amountExclTax: 0,
  taxRate: 0,
  enabled,
  outputEnabled: true,
});
const blueprintWithItems = costItems => ({
  id: "test",
  name: "测试蓝图",
  parameters,
  revenueItems: [],
  costItems,
  mappings: [],
});

// 智算 → ICT → 智算来源闭环与上下文清理。
const navigation = useNavigationStore;
navigation.getState().openIctFromIntelligentCompute({
  type: "intelligent_compute",
  workspaceId: "workspace-a",
  projectId: "project-intelligent",
  projectName: "智算项目 A",
  amountSourceId: "amount-source-a",
});
assert.equal(navigation.getState().currentView, "ict_lifecycle");
assert.equal(navigation.getState().ictOrigin.projectId, "project-intelligent");
navigation.getState().navigateTo(
  "ai_compute_quote",
  "project-intelligent",
  null,
  "amount-source-a",
);
assert.equal(navigation.getState().ictOrigin.amountSourceId, "amount-source-a");
navigation.getState().openIctFromIntelligentCompute({
  type: "intelligent_compute",
  workspaceId: "workspace-a",
  projectId: "project-intelligent",
  projectName: "智算项目 A",
  amountSourceId: "amount-source-a",
});
navigation.getState().navigateTo("ict_lifecycle", "project-other");
assert.equal(navigation.getState().ictOrigin, null);
navigation.getState().navigateTo("hub");
navigation.getState().navigateTo("ict_lifecycle", "project-direct");
assert.equal(navigation.getState().ictOrigin, null);
navigation.getState().openIctFromIntelligentCompute({
  type: "intelligent_compute",
  workspaceId: "workspace-a",
  projectId: "project-intelligent",
  projectName: "智算项目 A",
  amountSourceId: null,
});
navigation.getState().clearContext();
assert.equal(navigation.getState().activeProjectId, null);
assert.equal(navigation.getState().ictOrigin, null);

// 金额来源管理：H200 标准作为默认基底和受保护基准，普通来源可删除。
const h200BaselineSource = {
  id: "source-h200",
  description: null,
  metadata: { sourceRole: "h200_baseline" },
  sourceVersion: 1,
  createdAt: "2026-06-01T00:00:00Z",
};
const legacyH200BaselineSource = {
  id: "source-h200-legacy",
  description: "智算项目默认金额来源",
  metadata: {},
  sourceVersion: 1,
  createdAt: "2026-06-01T00:00:00Z",
};
const normalAmountSource = {
  id: "source-quote",
  description: "普通报价来源",
  metadata: {},
  sourceVersion: 4,
  createdAt: "2026-06-02T00:00:00Z",
};
const h200PresetCopySource = {
  id: "source-h200-copy",
  description: "64 台 H200、5 年服务期的标准报价预设。金额口径为元、含税。",
  metadata: {},
  sourceVersion: 2,
  createdAt: "2026-06-03T00:00:00Z",
};
assert.equal(getDefaultCreateAmountSourceBaseMode(), "h200");
assert.equal(isH200BaselineAmountSource(h200BaselineSource), true);
assert.equal(isH200BaselineAmountSource(legacyH200BaselineSource), true);
assert.equal(isH200BaselineAmountSource(normalAmountSource), false);
assert.equal(canDeleteIntelligentAmountSource([normalAmountSource], normalAmountSource.id), false);
assert.equal(
  canDeleteIntelligentAmountSource([h200BaselineSource, normalAmountSource], h200BaselineSource.id),
  false,
);
assert.equal(
  canDeleteIntelligentAmountSource([h200BaselineSource, normalAmountSource], normalAmountSource.id),
  true,
);
assert.equal(
  canDeleteIntelligentAmountSource([h200BaselineSource, h200PresetCopySource], h200PresetCopySource.id),
  true,
);
const staleSyncLock = buildIntelligentComputeSyncLock(
  { projectId: "project-1", syncRevision: 2 },
  [h200BaselineSource, normalAmountSource],
);
assert.equal(staleSyncLock.expectedSyncRevision, 2);
assert.equal(
  JSON.stringify(staleSyncLock.sourceVersions),
  JSON.stringify({ "source-h200": 1, "source-quote": 4 }),
);
const returnedSyncLock = buildIntelligentComputeSyncLock(
  { projectId: "project-1", syncRevision: 3 },
  [h200BaselineSource, normalAmountSource],
);
assert.equal(returnedSyncLock.expectedSyncRevision, 3);
const persistedIctResult = restoreIctResultFromProjectState({
  syncRevision: 4,
  lastResult: {
    npv: 900,
    npv_rate: "0.12",
    margin_rate: "0.28",
    dynamic_payback: ">10",
    irr: "--",
    it_npv: "800",
    it_npv_rate: "0.1",
    it_margin_rate: "0.26",
    cashflow: [{ year: 1, cash_in: 100, cash_out: "20", net_cash: "80", cum_net_cash: "80", pv: "76", cum_pv: "76" }],
  },
});
assert.equal(persistedIctResult.npv, "900");
assert.equal(persistedIctResult.cashflow[0].cash_in, "100");
assert.equal(restoreIctResultFromProjectState({ syncRevision: 0, lastResult: { npv: "1" } }), null);
assert.equal(restoreIctResultFromProjectState({ syncRevision: 1, lastResult: { npv: "1" } }), null);
assert.equal(
  projectStateSyncIncludesAmountSource({
    controlledSubjects: {
      "revenue:rev_it_cloud": {
        side: "revenue",
        ictSubjectCode: "rev_it_cloud",
        amountInclTax: 100,
        taxRate: 0.06,
        yearlyAmounts: [100],
        sourceLineItemIds: ["source-quote:item-1"],
      },
    },
  }, "source-quote"),
  true,
);
assert.equal(
  projectStateSyncIncludesAmountSource({
    controlledSubjects: {
      "revenue:rev_it_cloud": {
        side: "revenue",
        ictSubjectCode: "rev_it_cloud",
        amountInclTax: 100,
        taxRate: 0.06,
        yearlyAmounts: [100],
        sourceLineItemIds: ["source-other:item-1"],
      },
    },
  }, "source-quote"),
  false,
);

// 金额来源交换包：只导出当前来源业务结构，不携带项目身份、版本或同步控制状态。
const h200ForExchange = createH200Blueprint();
const exchangeBlueprint = {
  ...h200ForExchange,
  revenueItems: h200ForExchange.revenueItems.map((lineItem, index) => index === 0
    ? {
        ...lineItem,
        formulaControlStatus: "ict_override",
        ictOverride: {
          ictSubjectCode: "revenue_it_service",
          amountInclTax: 123,
          taxRate: 0.06,
          yearlyAmounts: [123, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          modifiedAt: "2026-06-16T00:00:00Z",
        },
        ictControlMessage: "历史 ICT 覆盖",
      }
    : lineItem),
  syncState: { revision: 9, status: "synced", subjects: {} },
};
const sourceForExchange = {
  id: "source-h200",
  projectId: "project-source",
  name: "H200 标准",
  description: "来源描述",
  enabled: true,
  sourceVersion: 7,
  metadata: { sourceRole: "h200_baseline", scenarioId: "old-scenario", customTag: "keep" },
  parameterGroups: exchangeBlueprint.parameterGroups,
  parameters: exchangeBlueprint.parameters,
  revenueItems: exchangeBlueprint.revenueItems,
  costItems: exchangeBlueprint.costItems,
  mappings: exchangeBlueprint.mappings,
  calculationSnapshot: {
    syncState: { revision: 9 },
    formalResult: { npv: 1 },
    summary: { totalRevenue: 1 },
  },
  createdAt: "2026-06-01T00:00:00Z",
  updatedAt: "2026-06-02T00:00:00Z",
};
const exportedPackage = buildAmountSourcePackage(sourceForExchange, exchangeBlueprint, {
  projectId: "project-source",
  projectYears: 5,
  discountRate: 0.05,
});
assert.equal(exportedPackage.kind, AMOUNT_SOURCE_PACKAGE_KIND);
assert.equal(Object.prototype.hasOwnProperty.call(exportedPackage.source, "projectId"), false);
assert.equal(Object.prototype.hasOwnProperty.call(exportedPackage.source, "sourceVersion"), false);
assert.equal(exportedPackage.source.metadata.sourceRole, undefined);
assert.equal(exportedPackage.source.metadata.scenarioId, undefined);
assert.equal(exportedPackage.source.metadata.customTag, "keep");
assert.equal(exportedPackage.source.calculationSnapshot.syncState, undefined);
assert.equal(exportedPackage.source.calculationSnapshot.formalResult, undefined);
assert.equal(exportedPackage.source.revenueItems[0].formulaControlStatus, undefined);
assert.equal(exportedPackage.source.revenueItems[0].ictOverride, undefined);
assert.equal(exportedPackage.source.revenueItems[0].ictControlMessage, undefined);
const normalizedPackage = normalizeAmountSourcePackage(exportedPackage);
assert.equal(getDefaultImportedAmountSourceName(normalizedPackage), `${exportedPackage.source.name}（导入）`);
const importedBlueprint = buildBlueprintFromAmountSourcePackage(normalizedPackage, {
  sourceId: "source-imported",
  name: "导入副本",
  projectYears: 3,
  discountRate: 0.08,
});
assert.equal(importedBlueprint.id, "source-imported");
assert.equal(importedBlueprint.scenarioId, "source-imported");
assert.equal(importedBlueprint.name, "导入副本");
assert.equal(getAiComputeProjectCycleYears(importedBlueprint.parameters), 3);
assert.equal(getAiComputeDiscountRatePercent(importedBlueprint.parameters), 8);
assert.equal(importedBlueprint.syncState, undefined);
const importedWithCurrentProjectSettings = buildBlueprintFromAmountSourcePackage(normalizedPackage, {
  sourceId: "source-imported-current-project",
  name: "保留当前项目参数",
  projectYears: 4,
  discountRate: 0.06,
});
assert.equal(getAiComputeProjectCycleYears(importedWithCurrentProjectSettings.parameters), 4);
assert.equal(getAiComputeDiscountRatePercent(importedWithCurrentProjectSettings.parameters), 6);

// 1. 参数 × 参数 × 固定值。
const multiplyResult = evaluateQuoteFormula(expression([
  { type: "parameter", id: "a", name: "参数A" },
  { type: "operator", operator: "*" },
  { type: "parameter", id: "b", name: "参数B" },
  { type: "operator", operator: "*" },
  { type: "constant", value: 3 },
]), parameters);
assert.equal(multiplyResult.value, 60);
assert.equal(multiplyResult.status, "valid");

const parenthesisResult = evaluateQuoteFormula(expression([
  { type: "left_parenthesis" },
  { type: "parameter", id: "a", name: "参数A" },
  { type: "operator", operator: "+" },
  { type: "parameter", id: "b", name: "参数B" },
  { type: "right_parenthesis" },
  { type: "operator", operator: "*" },
  { type: "constant", value: 2 },
]), parameters);
assert.equal(parenthesisResult.value, 18);

// 光标可在 Token 中间插入，并在编辑后保持稳定位置。
const cursorTokens = [
  { type: "parameter", id: "a", name: "参数A" },
  { type: "operator", operator: "*" },
  { type: "constant", value: 2 },
];
const middleInsert = insertFormulaTokensAt(cursorTokens, 1, [
  { type: "operator", operator: "+" },
  { type: "parameter", id: "b", name: "参数B" },
]);
assert.deepEqual(Array.from(middleInsert.tokens, token => token.type), [
  "parameter", "operator", "parameter", "operator", "constant",
]);
assert.equal(middleInsert.cursor, 3);
const deleteBeforeCursor = removeFormulaTokenBeforeCursor(middleInsert.tokens, middleInsert.cursor);
assert.equal(deleteBeforeCursor.tokens[1].type, "operator");
assert.equal(deleteBeforeCursor.cursor, 2);
const deleteAfterCursor = removeFormulaTokenAt(deleteBeforeCursor.tokens, 3, 1);
assert.equal(deleteAfterCursor.cursor, 1);
assert.equal(clampFormulaCursor(99, deleteAfterCursor.tokens.length), deleteAfterCursor.tokens.length);

// 兼容旧版结构化公式。
assert.equal(evaluateQuoteFormula({ type: "multiply", operands: [
  { type: "parameter", parameterId: "a" },
  { type: "constant", value: 3 },
] }, parameters).value, 15);

// 2. SUM(机器成本, 维保成本, 组网成本) × 资金成本率。
const fundingItems = [
  item("machine", "机器成本", expression([{ type: "constant", value: 100 }])),
  item("maintenance", "维保成本", expression([{ type: "constant", value: 20 }])),
  item("network", "组网成本", expression([{ type: "constant", value: 30 }])),
  item("capital", "资金成本", expression([
    { type: "function", name: "SUM" },
    { type: "line_item", id: "machine", name: "机器成本" },
    { type: "comma" },
    { type: "line_item", id: "maintenance", name: "维保成本" },
    { type: "comma" },
    { type: "line_item", id: "network", name: "组网成本" },
    { type: "right_parenthesis" },
    { type: "operator", operator: "*" },
    { type: "parameter", id: "rate", name: "资金成本率" },
    { type: "operator", operator: "*" },
    { type: "constant", value: 0.01 },
  ])),
];
const fundingCalculated = calculateQuoteBlueprint(blueprintWithItems(fundingItems));
assert.equal(fundingCalculated.costItems.find(current => current.id === "capital").amountInclTax, 15);

// 3. 修改被引用计算项后，下游自动更新。
const updatedFunding = calculateQuoteBlueprint(blueprintWithItems(fundingItems.map(current =>
  current.id === "machine"
    ? { ...current, formula: expression([{ type: "constant", value: 200 }]) }
    : current
)));
assert.equal(updatedFunding.costItems.find(current => current.id === "capital").amountInclTax, 25);

// 4. 循环引用只标记相关项，不抛异常。
const circular = calculateQuoteBlueprint(blueprintWithItems([
  item("cycle-a", "循环 A", expression([{ type: "line_item", id: "cycle-b", name: "循环 B" }])),
  item("cycle-b", "循环 B", expression([{ type: "line_item", id: "cycle-a", name: "循环 A" }])),
  item("independent", "独立项", expression([{ type: "constant", value: 9 }])),
]));
assert.equal(circular.costItems.find(current => current.id === "cycle-a").calculationStatus, "error");
assert.match(circular.costItems.find(current => current.id === "cycle-b").calculationError, /循环引用/);
assert.equal(circular.costItems.find(current => current.id === "independent").amountInclTax, 9);

// 5. 除零不会崩溃。
const divideByZero = evaluateQuoteFormula(expression([
  { type: "constant", value: 10 },
  { type: "operator", operator: "/" },
  { type: "constant", value: 0 },
]), parameters);
assert.equal(divideByZero.status, "error");
assert.match(divideByZero.errors[0], /除数不能为 0/);

// 6. 不存在的参数和计算项返回明确错误。
const missingParameter = evaluateQuoteFormula(expression([
  { type: "parameter", id: "missing", name: "不存在参数" },
]), parameters);
assert.match(missingParameter.errors[0], /不存在/);
const missingItem = calculateQuoteBlueprint(blueprintWithItems([
  item("broken", "缺失引用", expression([{ type: "line_item", id: "missing-item", name: "不存在计算项" }])),
]));
assert.match(missingItem.costItems[0].calculationError, /不存在/);

// 禁用的被引用项按 0 计算并返回提示。
const disabledReference = calculateQuoteBlueprint(blueprintWithItems([
  item("disabled", "禁用项", expression([{ type: "constant", value: 100 }]), false),
  item("consumer", "引用项", expression([
    { type: "line_item", id: "disabled", name: "禁用项" },
    { type: "operator", operator: "+" },
    { type: "constant", value: 5 },
  ])),
]));
assert.equal(disabledReference.costItems.find(current => current.id === "consumer").amountInclTax, 5);
assert.match(disabledReference.costItems.find(current => current.id === "consumer").calculationWarnings[0], /按 0/);

// 7. 折叠状态是组件本地 UI 状态，不进入蓝图，计算结果不受影响。
const collapsedResult = calculateQuoteBlueprint(blueprintWithItems(fundingItems));
const expandedResult = calculateQuoteBlueprint(blueprintWithItems(fundingItems));
assert.equal(collapsedResult.costItems.find(current => current.id === "capital").amountInclTax,
  expandedResult.costItems.find(current => current.id === "capital").amountInclTax);
assert.equal(Object.hasOwn(collapsedResult.costItems[0], "expanded"), false);

// H200 默认蓝图和依赖公式。
const h200 = calculateQuoteBlueprint(createH200Blueprint());
assert.equal(h200.revenueItems.find(current => current.id === "revenue-gpu-service").amountInclTax, 345600000);
assert.equal(h200.costItems.find(current => current.id === "cost-machine").amountInclTax, 211200000);
assert.equal(h200.costItems.find(current => current.id === "cost-capital").amountInclTax, 24000000);
assert.equal(h200.revenueItems.find(current => current.id === "revenue-bandwidth").amountInclTax, 76800000);
assert.equal(h200.costItems.find(current => current.id === "cost-bandwidth").amountInclTax, 86400000);
const summary = summarizeQuote(h200);
assert.equal(summary.totalRevenue, 448920000);
assert.equal(summary.totalCost, 368352000);
assert.equal(summary.totalRevenueExclTax, 423509433.96);
assert.equal(summary.totalCostExclTax, 334037201.53);
assert.equal(summary.grossProfit, 89472232.43);
assert.equal(summary.grossMarginRate, 21.13);
assert.equal(summary.costPerDeviceMonth, 86988.85);

// 参数类别布局与旧蓝图迁移。
const legacyLayoutBlueprint = createH200Blueprint();
delete legacyLayoutBlueprint.parameterGroups;
legacyLayoutBlueprint.parameters = legacyLayoutBlueprint.parameters.map(parameter => {
  const { groupId, isKey, ...legacyParameter } = parameter;
  return legacyParameter;
}).filter(parameter => parameter.id !== "discount-rate");
legacyLayoutBlueprint.parameters.push({
  id: "legacy-custom",
  name: "旧自定义参数",
  key: "legacy_custom",
  value: 1,
  category: "custom",
});
const migratedLayout = normalizeAiComputeParameterLayout(legacyLayoutBlueprint);
assert.equal(migratedLayout.parameterGroups.length, DEFAULT_PARAMETER_GROUPS.length);
assert.equal(
  migratedLayout.parameters.find(parameter => parameter.id === "device-count").groupId,
  PARAMETER_GROUP_IDS.scale,
);
assert.equal(
  migratedLayout.parameters.find(parameter => parameter.id === "machine-price").groupId,
  PARAMETER_GROUP_IDS.investment,
);
assert.equal(
  migratedLayout.parameters.find(parameter => parameter.id === "legacy-custom").groupId,
  PARAMETER_GROUP_IDS.unclassified,
);
assert.equal(
  migratedLayout.parameters.find(parameter => parameter.id === "device-count").isKey,
  true,
);
assert.equal(
  migratedLayout.parameters.find(parameter => parameter.id === "bandwidth-cost-price").isKey,
  false,
);
const migratedProjectCycle = migratedLayout.parameters.find(parameter => parameter.id === "years");
assert.equal(migratedProjectCycle.key, "years");
assert.equal(migratedProjectCycle.locked, true);
assert.equal(migratedProjectCycle.groupId, PARAMETER_GROUP_IDS.scale);
const repairedProjectCycle = normalizeAiComputeParameterLayout({
  ...migratedLayout,
  parameters: migratedLayout.parameters.map(parameter =>
    parameter.id === "years"
      ? { ...parameter, key: "changed_year_key", value: 12.8, locked: false }
      : parameter
  ),
}).parameters.find(parameter => parameter.id === "years");
assert.equal(repairedProjectCycle.key, "years");
assert.equal(repairedProjectCycle.value, 10);
assert.equal(repairedProjectCycle.locked, true);
const migratedDiscountRate = migratedLayout.parameters.find(parameter => parameter.id === "discount-rate");
assert.equal(migratedDiscountRate.key, "discount_rate");
assert.equal(migratedDiscountRate.value, 5.5);
assert.equal(migratedDiscountRate.locked, true);
const projectInitializedDiscountRate = normalizeAiComputeParameterLayout(
  legacyLayoutBlueprint,
  { fallbackDiscountRatePercent: ictDiscountRateToAiComputePercent(0.05) },
).parameters.find(parameter => parameter.id === "discount-rate");
assert.equal(projectInitializedDiscountRate.value, 5);
assert.equal(
  initializeAiComputeDiscountRate(migratedLayout, 5).parameters
    .find(parameter => parameter.id === "discount-rate").value,
  5,
);
assert.equal(
  migrateAiComputeDiscountRateOwnership(migratedLayout, 3, 5).parameters
    .find(parameter => parameter.id === "discount-rate").value,
  5,
);
assert.equal(
  migrateAiComputeDiscountRateOwnership(
    initializeAiComputeDiscountRate(migratedLayout, 6),
    4,
    5,
  ).parameters.find(parameter => parameter.id === "discount-rate").value,
  6,
);
assert.equal(normalizeAiComputeDiscountRatePercent(-1), 0);
assert.equal(normalizeAiComputeDiscountRatePercent(5.123456), 5.1235);
assert.equal(normalizeAiComputeDiscountRatePercent(120), 100);
assert.equal(getAiComputeDiscountRatePercent(migratedLayout.parameters), 5.5);
assert.equal(getAiComputeDiscountRateDecimal(migratedLayout.parameters), 0.055);

const customGroup = {
  id: "parameter-group-custom",
  name: "自定义类别",
  description: "测试类别",
  builtin: false,
};
const reorderedGroups = reorderAiComputeParameterGroup(
  [...migratedLayout.parameterGroups, customGroup],
  customGroup.id,
  PARAMETER_GROUP_IDS.scale,
);
assert.equal(reorderedGroups[0].id, customGroup.id);
assert.equal(canDeleteAiComputeParameterGroup(customGroup, migratedLayout.parameters), true);
assert.equal(
  canDeleteAiComputeParameterGroup(
    customGroup,
    [{ ...migratedLayout.parameters[0], groupId: customGroup.id }],
  ),
  false,
);
assert.equal(
  canDeleteAiComputeParameterGroup(migratedLayout.parameterGroups[0], []),
  false,
);

const movedAcrossGroups = moveAiComputeParameter(
  migratedLayout.parameters,
  "device-count",
  PARAMETER_GROUP_IDS.operations,
);
assert.equal(
  movedAcrossGroups.find(parameter => parameter.id === "device-count").groupId,
  PARAMETER_GROUP_IDS.operations,
);
const operationIds = movedAcrossGroups
  .filter(parameter => parameter.groupId === PARAMETER_GROUP_IDS.operations)
  .map(parameter => parameter.id);
assert.equal(operationIds.at(-1), "device-count");
const movedWithinGroup = moveAiComputeParameterByOffset(
  movedAcrossGroups,
  "device-count",
  -1,
);
const reorderedOperationIds = movedWithinGroup
  .filter(parameter => parameter.groupId === PARAMETER_GROUP_IDS.operations)
  .map(parameter => parameter.id);
assert.equal(
  reorderedOperationIds.indexOf("device-count"),
  operationIds.indexOf("device-count") - 1,
);

const persistedLayout = JSON.parse(JSON.stringify({
  ...migratedLayout,
  parameterGroups: reorderedGroups.map(group =>
    group.id === customGroup.id ? { ...group, name: "交付类" } : group
  ),
  parameters: movedWithinGroup,
}));
const reloadedLayout = normalizeAiComputeParameterLayout(persistedLayout);
assert.equal(reloadedLayout.parameterGroups[0].id, customGroup.id);
assert.equal(reloadedLayout.parameterGroups[0].name, "交付类");
assert.deepEqual(
  Array.from(reloadedLayout.parameters
    .filter(parameter => parameter.groupId === PARAMETER_GROUP_IDS.operations)
    .map(parameter => parameter.id)),
  Array.from(reorderedOperationIds),
);

const layoutFingerprintBefore = getAiComputeSyncFingerprint(migratedLayout);
const layoutFingerprintAfter = getAiComputeSyncFingerprint({
  ...migratedLayout,
  parameterGroups: reorderedGroups,
  parameters: moveAiComputeParameter(
    migratedLayout.parameters,
    "device-count",
    PARAMETER_GROUP_IDS.operations,
  ).map(parameter => ({ ...parameter, isKey: !parameter.isKey })),
});
assert.equal(layoutFingerprintAfter, layoutFingerprintBefore);

const outputBlueprint = createH200Blueprint();
outputBlueprint.mappings = outputBlueprint.mappings.map(mapping =>
  mapping.lineItemId === "revenue-cabinet"
    ? { ...mapping, ictSubjectCode: "rev_it_cloud", ictSubjectName: "移动云-定制化收入" }
    : mapping
);
outputBlueprint.revenueItems = outputBlueprint.revenueItems.map(current =>
  current.id === "revenue-bandwidth" ? { ...current, outputEnabled: false } : current
);
outputBlueprint.mappings = outputBlueprint.mappings.filter(mapping => mapping.lineItemId !== "cost-capital");
const output = buildAiComputeQuoteOutput(outputBlueprint);
const merged = output.find(current => current.ictSubjectCode === "rev_it_cloud");
assert.equal(merged.sourceLineItemIds.length, 2);
assert.equal(output.some(current => current.ictSubjectCode === "rev_ct_line"), false);
assert.equal(output.some(current => current.sourceLineItemIds.includes("cost-capital")), false);

// 智算业务项资金计划。
// 1. 第一年度一次性计划。
const firstYearAmounts = buildAiComputeFirstYearAmounts(1000);
assert.equal(firstYearAmounts["1"], 1000);
assert.equal(firstYearAmounts["2"], 0);
assert.equal(Object.keys(firstYearAmounts).length, 10);

// 2. 平均分年计划按分处理尾差，周期外年份为 0。
const evenAmounts = buildAiComputeEvenAmounts(100, 3);
assert.equal(evenAmounts["1"], 33.33);
assert.equal(evenAmounts["2"], 33.33);
assert.equal(evenAmounts["3"], 33.34);
assert.equal(evenAmounts["4"], 0);

// 2.1 项目周期始终归一化为 1-10 年整数。
assert.equal(normalizeAiComputeProjectCycleValue(0), 1);
assert.equal(normalizeAiComputeProjectCycleValue(1), 1);
assert.equal(normalizeAiComputeProjectCycleValue(5.9), 5);
assert.equal(normalizeAiComputeProjectCycleValue(10), 10);
assert.equal(normalizeAiComputeProjectCycleValue(99), 10);
assert.equal(getAiComputeProjectCycleYears([{ id: "years", key: "changed", value: 7.8 }]), 7);

// 3. 手工修改后切换为 manual，合计正确。
let manualPlan = createDefaultAiComputeFundingPlan(100);
manualPlan = updateAiComputeFundingPlanYear(manualPlan, 1, 40);
manualPlan = updateAiComputeFundingPlanYear(manualPlan, 2, 50);
assert.equal(manualPlan.mode, "manual");
assert.equal(sumAiComputeFundingPlan(manualPlan), 90);
assert.deepEqual(
  normalizeAiComputeFundingPlan(manualPlan, 100, 5).yearlyAmounts,
  normalizeAiComputeFundingPlan(manualPlan, 100, 10).yearlyAmounts,
);
assert.deepEqual(
  normalizeAiComputeFundingPlan(createDefaultAiComputeFundingPlan(100), 100, 5).yearlyAmounts,
  normalizeAiComputeFundingPlan(createDefaultAiComputeFundingPlan(100), 100, 10).yearlyAmounts,
);
assert.notDeepEqual(
  normalizeAiComputeFundingPlan({ ...manualPlan, mode: "even" }, 100, 5).yearlyAmounts,
  normalizeAiComputeFundingPlan({ ...manualPlan, mode: "even" }, 100, 10).yearlyAmounts,
);

// 4. 计划与科目金额不一致时识别差额。
const manualValidation = validateAiComputeFundingPlan(manualPlan, 100);
assert.equal(manualValidation.consistent, false);
assert.equal(manualValidation.difference, 10);

// 5. 多个业务项映射同一 ICT 科目时逐年合并。
const fundingOutputBlueprint = blueprintWithItems([
  {
    ...item("funding-a", "资金项 A", expression([{ type: "constant", value: 150 }])),
    fundingPlan: {
      enabled: true,
      mode: "manual",
      yearlyAmounts: { "1": 100, "2": 50 },
    },
  },
  {
    ...item("funding-b", "资金项 B", expression([{ type: "constant", value: 300 }])),
    fundingPlan: {
      enabled: true,
      mode: "manual",
      yearlyAmounts: { "1": 300, "2": 0 },
    },
  },
]);
fundingOutputBlueprint.mappings = [
  {
    id: "mapping-funding-a",
    lineItemId: "funding-a",
    side: "cost",
    ictSubjectCode: "cost_it_device",
    ictSubjectName: "主要设备/甲供材料",
    enabled: true,
  },
  {
    id: "mapping-funding-b",
    lineItemId: "funding-b",
    side: "cost",
    ictSubjectCode: "cost_it_device",
    ictSubjectName: "主要设备/甲供材料",
    enabled: true,
  },
];
const mergedFundingOutput = buildAiComputeQuoteOutputFundingPlans(fundingOutputBlueprint);
assert.equal(mergedFundingOutput.length, 1);
assert.equal(mergedFundingOutput[0].yearlyAmounts["1"], 400);
assert.equal(mergedFundingOutput[0].yearlyAmounts["2"], 50);
assert.equal(mergedFundingOutput[0].totalAmount, 450);
assert.deepEqual(Array.from(mergedFundingOutput[0].sourceLineItemIds), ["funding-a", "funding-b"]);

const practicalFundingBlueprint = blueprintWithItems([
  {
    ...item("practical-a", "业务项 A", expression([{ type: "constant", value: 90 }])),
    fundingPlan: {
      enabled: true,
      mode: "manual",
      yearlyAmounts: { "1": 50, "2": 40 },
    },
  },
  {
    ...item("practical-b", "业务项 B", expression([{ type: "constant", value: 40 }])),
    fundingPlan: {
      enabled: true,
      mode: "manual",
      yearlyAmounts: { "1": 40 },
    },
  },
]);
practicalFundingBlueprint.mappings = ["practical-a", "practical-b"].map(lineItemId => ({
  id: `mapping-${lineItemId}`,
  lineItemId,
  side: "cost",
  ictSubjectCode: "cost_it_other",
  ictSubjectName: "其他投入",
  enabled: true,
}));
const practicalFundingOutput = buildAiComputeQuoteOutputFundingPlans(practicalFundingBlueprint);
assert.equal(practicalFundingOutput[0].yearlyAmounts["1"], 90);
assert.equal(practicalFundingOutput[0].yearlyAmounts["2"], 40);

// 6. 禁用项、关闭输出、未映射项和关闭计划不参与年度合并。
const excludedFundingBlueprint = blueprintWithItems([
  {
    ...item("disabled-item", "禁用项", expression([{ type: "constant", value: 10 }]), false),
    fundingPlan: createDefaultAiComputeFundingPlan(10),
  },
  {
    ...item("output-off", "未输出项", expression([{ type: "constant", value: 20 }])),
    outputEnabled: false,
    fundingPlan: createDefaultAiComputeFundingPlan(20),
  },
  {
    ...item("unmapped", "未映射项", expression([{ type: "constant", value: 30 }])),
    fundingPlan: createDefaultAiComputeFundingPlan(30),
  },
  {
    ...item("plan-off", "计划关闭项", expression([{ type: "constant", value: 40 }])),
    fundingPlan: { ...createDefaultAiComputeFundingPlan(40), enabled: false },
  },
]);
excludedFundingBlueprint.mappings = ["disabled-item", "output-off", "plan-off"].map(lineItemId => ({
  id: `mapping-${lineItemId}`,
  lineItemId,
  side: "cost",
  ictSubjectCode: "cost_it_device",
  ictSubjectName: "主要设备/甲供材料",
  enabled: true,
}));
assert.equal(buildAiComputeQuoteOutputFundingPlans(excludedFundingBlueprint).length, 0);

// 输出到 ICT：差异预览复用合并结果，正式 payload 同时写入金额、计划和来源追踪。
const existingIctState = {
  project: {
    id: "project-1",
    name: "测试项目",
    customer_name: "客户",
    project_years: 5,
    discount_rate: 6,
    cashflow_model: "model_a",
  },
  lifecycleState: {
    profileJson: { projectName: "测试项目", customerName: "客户" },
    parametersJson: {},
    backgroundJson: {},
    inputPayloadJson: {
      cost_it_device: { incl_tax: "80", tax_rate: "13" },
    },
  },
  cashflowState: {
    cashflowModel: "model_a",
    paymentModelJson: {},
    yearlyCashflowJson: { stale: true },
    sectorCashflowJson: {},
    assumptionsJson: {
      costIt: {
        device: { incl: 80, excl: 70.8, tax: 13 },
        construction: { incl: 0, excl: 0, tax: 9 },
      },
      subjectFundingPlans: {
        "cost:costIt:device": {
          id: "cost:costIt:device",
          subjectRef: { side: "cost", groupId: "costIt", key: "device" },
          mode: "upfront",
          annualInclValues: [80, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          enabled: true,
          source: "manual",
        },
      },
    },
    metricsJson: { stale: true },
  },
};
const exportPreview = buildAiComputeIctExportPreview(
  fundingOutputBlueprint,
  existingIctState,
  "project-1",
);
assert.equal(exportPreview.rows.length, 1);
assert.equal(exportPreview.rows[0].originalAmount, 80);
assert.equal(exportPreview.rows[0].writtenAmount, 450);
assert.equal(exportPreview.rows[0].yearlyAmounts[0], 400);
assert.equal(exportPreview.rows[0].yearlyAmounts[1], 50);
assert.deepEqual(Array.from(exportPreview.rows[0].sourceLineItemNames), ["资金项 A", "资金项 B"]);

const h200ExportPreview = buildAiComputeIctExportPreview(h200, existingIctState, "project-1");
assert.equal(h200ExportPreview.projectYears, 5);
const h200GpuRevenueRow = h200ExportPreview.rows.find(row => row.ictSubjectCode === "rev_it_cloud");
assert.equal(h200GpuRevenueRow.writtenAmount, 345600000);
assert.equal(h200GpuRevenueRow.amountExclTax, 326037735.85);
assert.equal(h200GpuRevenueRow.taxRate, 6);

const exportPayloads = buildAiComputeIctExportPayloads(exportPreview, existingIctState);
assert.equal(exportPreview.projectYears, 10);
assert.equal(exportPayloads.lifecycleState.inputPayloadJson.project_name, "测试项目");
assert.equal(exportPayloads.lifecycleState.inputPayloadJson.customer_name, "客户");
assert.equal(exportPayloads.lifecycleState.inputPayloadJson.property_rights, "客户");
assert.equal(exportPayloads.lifecycleState.inputPayloadJson.cashflow_model, "model_a");
assert.deepEqual(
  Array.from(exportPayloads.lifecycleState.inputPayloadJson.rev_distribution),
  [1, 0, 0, 0, 0, 0, 0, 0, 0, 0],
);
assert.deepEqual(
  Array.from(exportPayloads.lifecycleState.inputPayloadJson.cost_distribution),
  [1, 0, 0, 0, 0, 0, 0, 0, 0, 0],
);
assert.equal(exportPayloads.lifecycleState.inputPayloadJson.rev_it_integration.incl_tax, "0");
assert.equal(exportPayloads.lifecycleState.inputPayloadJson.rev_it_integration.tax_rate, "6");
assert.equal(exportPayloads.lifecycleState.inputPayloadJson.cost_it_construction.tax_rate, "9");
assert.equal(
  ICT_SUBJECT_DEFINITIONS.every(subject => (
    typeof exportPayloads.lifecycleState.inputPayloadJson[subject.subjectCode]?.incl_tax === "string"
    && typeof exportPayloads.lifecycleState.inputPayloadJson[subject.subjectCode]?.tax_rate === "string"
  )),
  true,
);
assert.equal(exportPayloads.lifecycleState.inputPayloadJson.cost_it_device.incl_tax, "450");
assert.equal(exportPayloads.lifecycleState.inputPayloadJson.project_years, 10);
assert.equal(exportPayloads.lifecycleState.parametersJson.projectYears, 10);
assert.equal(exportPayloads.cashflowState.assumptionsJson.projectYears, 10);
assert.equal(exportPayloads.cashflowState.assumptionsJson.costIt.device.incl, 450);
assert.equal(
  exportPayloads.cashflowState.assumptionsJson.subjectFundingPlans["cost:costIt:device"].annualInclValues[0],
  400,
);
assert.equal(
  exportPayloads.cashflowState.assumptionsJson.subjectFundingPlans["cost:costIt:device"].annualInclValues[1],
  50,
);
assert.equal(
  exportPayloads.cashflowState.assumptionsJson.subjectFundingPlans["cost:costIt:device"].source,
  "ai_compute_quote",
);
assert.equal(exportPayloads.cashflowState.assumptionsJson.aiComputeQuoteImport.projectId, "project-1");
assert.deepEqual(Object.keys(exportPayloads.cashflowState.yearlyCashflowJson), []);
assert.deepEqual(Object.keys(exportPayloads.cashflowState.metricsJson), []);

// 实时同步由智算折现率覆盖 ICT lifecycle、cashflow assumptions 和正式计算输入。
const fivePercentIctState = JSON.parse(JSON.stringify(existingIctState));
fivePercentIctState.project.discount_rate = 0.05;
fivePercentIctState.lifecycleState.inputPayloadJson.discount_rate = "0.05";
fivePercentIctState.cashflowState.assumptionsJson.discountRate = 0.05;
const fivePercentQuoteBlueprint = {
  ...fundingOutputBlueprint,
  parameters: [
    ...fundingOutputBlueprint.parameters,
    {
      id: "discount-rate",
      name: "项目折现率",
      key: "discount_rate",
      value: 5,
      unit: "%",
    },
  ],
};
const autoPreview = buildAiComputeAutoSyncPreview(
  fivePercentQuoteBlueprint,
  fivePercentIctState,
  "project-1",
);
const autoPayloads = buildAiComputeIctExportPayloads(autoPreview, fivePercentIctState);
assert.equal(autoPayloads.lifecycleState.inputPayloadJson.discount_rate, "0.05");
assert.equal(autoPayloads.cashflowState.assumptionsJson.discountRate, 0.05);
assert.equal(autoPayloads.lifecycleState.parametersJson.discountRate, 0.05);
assert.equal(autoPayloads.lifecycleState.inputPayloadJson.project_years, 10);
const finalizedIct = finalizeIctInputWithFundingPlans(
  autoPayloads.lifecycleState.inputPayloadJson,
  autoPayloads.cashflowState.assumptionsJson.subjectFundingPlans,
);
assert.equal(finalizedIct.coverage.valid, true);
assert.deepEqual(Array.from(finalizedIct.input.cost_cashflow_excl).slice(0, 2), ["400.00", "50.00"]);

// 已映射项目被禁用、关闭输出或资金计划无效时，受控 ICT 科目按 0 同步，避免残留旧值。
const zeroedPreview = buildAiComputeAutoSyncPreview(
  excludedFundingBlueprint,
  existingIctState,
  "project-1",
);
assert.equal(zeroedPreview.rows.length, 1);
assert.equal(zeroedPreview.rows[0].writtenAmount, 0);
assert.equal(zeroedPreview.rows[0].syncStatus, "zeroed_error");
assert.equal(zeroedPreview.rows[0].yearlyAmounts.every(value => value === 0), true);

// 兼容读取旧反写标记，但 ICT 人工金额不再覆盖智算公式。
const singleSourceBlueprint = {
  ...fundingOutputBlueprint,
  costItems: [fundingOutputBlueprint.costItems[0]],
  mappings: [fundingOutputBlueprint.mappings[0]],
};
const singlePreview = buildAiComputeAutoSyncPreview(singleSourceBlueprint, existingIctState, "project-1");
const syncedSingle = applySuccessfulAiComputeSync(
  singleSourceBlueprint,
  singlePreview,
  1,
  "2026-06-12T00:00:00.000Z",
);
const singleOverridePlan = {
  "cost:costIt:device": {
    id: "cost:costIt:device",
    subjectRef: { side: "cost", groupId: "costIt", key: "device" },
    mode: "custom",
    annualInclValues: [160, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    enabled: true,
    source: "manual",
  },
};
const singleReconciled = reconcileAiComputeBlueprintFromIct(
  syncedSingle,
  { cost_it_device: { incl_tax: "160", tax_rate: "13" } },
  { subjectFundingPlans: singleOverridePlan },
);
assert.equal(singleReconciled.changed, true);
assert.equal(singleReconciled.blueprint.costItems[0].formulaControlStatus, "ict_override");
assert.equal(singleReconciled.blueprint.costItems[0].ictOverride.amountInclTax, 160);
const overrideCalculated = calculateQuoteBlueprint(singleReconciled.blueprint);
assert.equal(overrideCalculated.costItems[0].amountInclTax, 150);
const mappingControlRestored = clearAiComputeControlForMappingChange(
  singleReconciled.blueprint,
  "funding-a",
);
assert.equal(mappingControlRestored.costItems[0].formulaControlStatus, "formula");
assert.equal(mappingControlRestored.costItems[0].ictOverride, undefined);
const restoredSingle = restoreAiComputeFormulaControl(singleReconciled.blueprint, "funding-a");
assert.equal(restoredSingle.costItems[0].formulaControlStatus, "formula");
assert.equal(restoredSingle.costItems[0].ictOverride, undefined);

// 改映射后，新科目写入当前金额，旧科目作为释放记录清零并从成功快照移除。
const remappedSingle = {
  ...syncedSingle,
  mappings: syncedSingle.mappings.map(mapping => ({
    ...mapping,
    ictSubjectCode: "cost_it_other",
    ictSubjectName: "其他投入",
  })),
};
const remappedPreview = buildAiComputeAutoSyncPreview(
  remappedSingle,
  existingIctState,
  "project-1",
);
const remappedNewRow = remappedPreview.rows.find(row => row.ictSubjectCode === "cost_it_other");
const releasedOldRow = remappedPreview.rows.find(row => row.ictSubjectCode === "cost_it_device");
assert.equal(remappedNewRow.writtenAmount, 150);
assert.equal(releasedOldRow.writtenAmount, 0);
assert.equal(releasedOldRow.syncStatus, "released_mapping");
assert.equal(releasedOldRow.yearlyAmounts.every(value => value === 0), true);
const remappedPayloads = buildAiComputeIctExportPayloads(remappedPreview, existingIctState);
assert.equal(remappedPayloads.lifecycleState.inputPayloadJson.cost_it_other.incl_tax, "150");
assert.equal(remappedPayloads.lifecycleState.inputPayloadJson.cost_it_device.incl_tax, "0");
assert.equal(
  remappedPayloads.cashflowState.assumptionsJson.subjectFundingPlans["cost:costIt:device"]
    .annualInclValues.every(value => value === 0),
  true,
);
assert.deepEqual(
  Array.from(remappedPayloads.cashflowState.assumptionsJson.aiComputeQuoteImport.releasedSubjectCodes),
  ["cost_it_device"],
);
const remappedFinalized = finalizeIctInputWithFundingPlans(
  remappedPayloads.lifecycleState.inputPayloadJson,
  remappedPayloads.cashflowState.assumptionsJson.subjectFundingPlans,
);
assert.equal(remappedFinalized.coverage.valid, true);
const remappedSynced = applySuccessfulAiComputeSync(
  remappedSingle,
  remappedPreview,
  2,
  "2026-06-15T00:00:00.000Z",
);
assert.equal(Boolean(remappedSynced.syncState.subjects["cost:cost_it_device"]), false);
assert.equal(Boolean(remappedSynced.syncState.subjects["cost:cost_it_other"]), true);

// 旧版本若未保存旧科目快照，仍应从未被人工修改的 ICT 导入痕迹恢复所有权并释放残留。
const legacyDeviceTrace = {
  source: "ai_compute_quote",
  sourceLabel: "来自智算报价测算",
  projectId: "project-1",
  scenarioId: "test",
  blueprintId: "test",
  sourceLineItemIds: ["funding-a"],
  importedAt: "2026-06-12T00:00:00.000Z",
};
const missingSnapshotIctState = JSON.parse(JSON.stringify(existingIctState));
missingSnapshotIctState.lifecycleState.inputPayloadJson.cost_it_device = {
  incl_tax: "150",
  excl_tax: "132.74",
  tax_rate: "13",
  import_trace: legacyDeviceTrace,
};
missingSnapshotIctState.cashflowState.assumptionsJson.costIt.device = {
  incl: 150,
  excl: 132.74,
  tax: 13,
  importTrace: legacyDeviceTrace,
};
missingSnapshotIctState.cashflowState.assumptionsJson.subjectFundingPlans["cost:costIt:device"] = {
  id: "cost:costIt:device",
  subjectRef: { side: "cost", groupId: "costIt", key: "device" },
  mode: "custom",
  annualInclValues: [150, 0, 0, 0, 0, 0, 0, 0, 0, 0],
  enabled: true,
  source: "ai_compute_quote",
  lastChangeReason: "ai_compute_quote_import",
  importTrace: legacyDeviceTrace,
};
const missingSnapshotPreview = buildAiComputeAutoSyncPreview(
  remappedSynced,
  missingSnapshotIctState,
  "project-1",
);
const recoveredLegacyRelease = missingSnapshotPreview.rows.find(row =>
  row.ictSubjectCode === "cost_it_device"
);
assert.equal(recoveredLegacyRelease.syncStatus, "released_mapping");
const missingSnapshotPayloads = buildAiComputeIctExportPayloads(
  missingSnapshotPreview,
  missingSnapshotIctState,
);
assert.equal(missingSnapshotPayloads.lifecycleState.inputPayloadJson.cost_it_device.incl_tax, "0");
assert.deepEqual(
  Array.from(missingSnapshotPayloads.cashflowState.assumptionsJson.aiComputeQuoteImport.releasedSubjectCodes),
  ["cost_it_device"],
);
const afterLegacyReleasePreview = buildAiComputeAutoSyncPreview(
  remappedSynced,
  {
    ...missingSnapshotIctState,
    lifecycleState: missingSnapshotPayloads.lifecycleState,
    cashflowState: missingSnapshotPayloads.cashflowState,
  },
  "project-1",
);
assert.equal(
  afterLegacyReleasePreview.rows.some(row =>
    row.ictSubjectCode === "cost_it_device" && row.syncStatus === "released_mapping"
  ),
  false,
);

// ICT 已人工调整过的遗留科目不通过导入痕迹兜底释放。
const manuallyChangedLegacyState = JSON.parse(JSON.stringify(missingSnapshotIctState));
manuallyChangedLegacyState.cashflowState.assumptionsJson.subjectFundingPlans[
  "cost:costIt:device"
].lastChangeReason = "manual_plan_edit";
const manuallyChangedLegacyPreview = buildAiComputeAutoSyncPreview(
  remappedSynced,
  manuallyChangedLegacyState,
  "project-1",
);
assert.equal(
  manuallyChangedLegacyPreview.rows.some(row =>
    row.ictSubjectCode === "cost_it_device" && row.syncStatus === "released_mapping"
  ),
  false,
);

// 释放规则覆盖完整 ICT 标准科目目录，不依赖甲供材料科目代码。
ICT_SUBJECT_DEFINITIONS.forEach(oldSubject => {
  const targetSubject = ICT_SUBJECT_DEFINITIONS.find(candidate =>
    candidate.side === oldSubject.side && candidate.subjectCode !== oldSubject.subjectCode
  );
  const genericItem = {
    ...item(
      "generic-line",
      "通用映射测试项",
      expression([{ type: "constant", value: 120 }]),
    ),
    side: oldSubject.side,
    fundingPlan: {
      enabled: true,
      mode: "manual",
      yearlyAmounts: { "1": 70, "2": 50 },
    },
  };
  const genericBlueprint = {
    ...blueprintWithItems([]),
    id: "catalog-release",
    scenarioId: "catalog-release-scenario",
    revenueItems: oldSubject.side === "revenue" ? [genericItem] : [],
    costItems: oldSubject.side === "cost" ? [genericItem] : [],
    mappings: [{
      id: "generic-mapping",
      lineItemId: genericItem.id,
      side: oldSubject.side,
      ictSubjectCode: targetSubject.subjectCode,
      ictSubjectName: targetSubject.standardSubjectName,
      enabled: true,
    }],
    syncState: undefined,
  };
  const genericTrace = {
    source: "ai_compute_quote",
    sourceLabel: "来自智算报价测算",
    projectId: "project-1",
    scenarioId: genericBlueprint.scenarioId,
    blueprintId: genericBlueprint.id,
    sourceLineItemIds: [genericItem.id],
    importedAt: "2026-06-12T00:00:00.000Z",
  };
  const genericState = JSON.parse(JSON.stringify(existingIctState));
  genericState.lifecycleState.inputPayloadJson[oldSubject.subjectCode] = {
    incl_tax: "120",
    excl_tax: "106.19",
    tax_rate: "13",
    import_trace: genericTrace,
  };
  const oldAssumptionItem = {
    incl: 120,
    excl: 106.19,
    tax: 13,
    importTrace: genericTrace,
  };
  if (oldSubject.groupId === "revNonItCt") {
    genericState.cashflowState.assumptionsJson.revNonItCt = oldAssumptionItem;
  } else {
    genericState.cashflowState.assumptionsJson[oldSubject.groupId] = {
      ...(genericState.cashflowState.assumptionsJson[oldSubject.groupId] || {}),
      [oldSubject.key]: oldAssumptionItem,
    };
  }
  const oldPlanId = `${oldSubject.side}:${oldSubject.groupId}:${oldSubject.key}`;
  genericState.cashflowState.assumptionsJson.subjectFundingPlans[oldPlanId] = {
    id: oldPlanId,
    subjectRef: {
      side: oldSubject.side,
      groupId: oldSubject.groupId,
      key: oldSubject.key,
    },
    mode: "custom",
    annualInclValues: [70, 50, 0, 0, 0, 0, 0, 0, 0, 0],
    enabled: true,
    source: "ai_compute_quote",
    lastChangeReason: "ai_compute_quote_import",
    importTrace: genericTrace,
  };

  const genericPreview = buildAiComputeAutoSyncPreview(
    genericBlueprint,
    genericState,
    "project-1",
  );
  const genericRelease = genericPreview.rows.find(row =>
    row.ictSubjectCode === oldSubject.subjectCode
  );
  assert.equal(
    genericRelease?.syncStatus,
    "released_mapping",
    `${oldSubject.subjectCode} 应生成释放记录`,
  );
  const genericPayloads = buildAiComputeIctExportPayloads(genericPreview, genericState);
  assert.equal(
    genericPayloads.lifecycleState.inputPayloadJson[oldSubject.subjectCode].incl_tax,
    "0",
    `${oldSubject.subjectCode} 应清零`,
  );
});

// 删除映射同样释放旧科目。
const unmappedSingle = { ...syncedSingle, mappings: [] };
const unmappedSyncPreview = buildAiComputeAutoSyncPreview(
  unmappedSingle,
  existingIctState,
  "project-1",
);
assert.equal(unmappedSyncPreview.rows.length, 1);
assert.equal(unmappedSyncPreview.rows[0].syncStatus, "released_mapping");

// 历史 ICT 覆盖标记不再暂停智算输出。
const pausedSingle = {
  ...syncedSingle,
  costItems: syncedSingle.costItems.map(current => ({
    ...current,
    formulaControlStatus: "ict_override",
    ictOverride: {
      ictSubjectCode: "cost_it_device",
      amountInclTax: 160,
      taxRate: 13,
      yearlyAmounts: [160, 0, 0, 0, 0, 0, 0, 0, 0, 0],
      modifiedAt: "2026-06-15T00:00:00.000Z",
    },
  })),
};
const pausedPreview = buildAiComputeAutoSyncPreview(pausedSingle, existingIctState, "project-1");
assert.equal(pausedPreview.rows[0].syncStatus, "ready");
assert.equal(pausedPreview.rows[0].writtenAmount, 150);
const pausedSynced = applySuccessfulAiComputeSync(
  pausedSingle,
  pausedPreview,
  2,
  "2026-06-15T00:00:00.000Z",
);
assert.equal(Boolean(pausedSynced.syncState.subjects["cost:cost_it_device"]), true);

// 旧版合并冲突标记可读取，但不会让智算公式失效。
const syncedMerged = applySuccessfulAiComputeSync(
  fundingOutputBlueprint,
  autoPreview,
  1,
  "2026-06-12T00:00:00.000Z",
);
// 原科目仍有其他来源时只写入剩余合计，不生成释放记录。
const partiallyRemapped = {
  ...syncedMerged,
  mappings: syncedMerged.mappings.map(mapping =>
    mapping.lineItemId === "funding-a"
      ? { ...mapping, ictSubjectCode: "cost_it_other", ictSubjectName: "其他投入" }
      : mapping
  ),
};
const partiallyRemappedPreview = buildAiComputeAutoSyncPreview(
  partiallyRemapped,
  existingIctState,
  "project-1",
);
const remainingDeviceRow = partiallyRemappedPreview.rows.find(row =>
  row.ictSubjectCode === "cost_it_device"
);
assert.equal(remainingDeviceRow.syncStatus, "ready");
assert.equal(remainingDeviceRow.writtenAmount, 300);
assert.equal(remainingDeviceRow.yearlyAmounts[0], 300);
assert.equal(
  partiallyRemappedPreview.rows.some(row => row.syncStatus === "released_mapping"),
  false,
);

const mergedReconciled = reconcileAiComputeBlueprintFromIct(
  syncedMerged,
  { cost_it_device: { incl_tax: "460", tax_rate: "13" } },
  {
    subjectFundingPlans: {
      "cost:costIt:device": {
        ...singleOverridePlan["cost:costIt:device"],
        annualInclValues: [410, 50, 0, 0, 0, 0, 0, 0, 0, 0],
      },
    },
  },
);
assert.equal(mergedReconciled.changed, true);
assert.equal(mergedReconciled.conflicts.length, 1);
assert.equal(mergedReconciled.blueprint.costItems.every(current =>
  current.formulaControlStatus === "merge_conflict"
), true);
assert.deepEqual(
  Array.from(calculateQuoteBlueprint(mergedReconciled.blueprint).costItems, current => current.amountInclTax),
  [150, 300],
);

// ICT 正式同步一次只选择一个金额来源；同一来源内仍按 side + ICT subjectCode 聚合。
const aggregateSourceA = {
  ...singleSourceBlueprint,
  id: "amount-source-a",
  scenarioId: "amount-source-a",
  name: "来源 A",
  costItems: [{
    ...singleSourceBlueprint.costItems[0],
    taxRate: 13,
  }],
};
const aggregateSourceB = {
  ...singleSourceBlueprint,
  id: "amount-source-b",
  scenarioId: "amount-source-b",
  name: "来源 B",
	  costItems: [{
	    ...singleSourceBlueprint.costItems[0],
	    id: "funding-b",
	    name: "资金项 B",
	    formula: expression([{ type: "constant", value: 300 }]),
	    taxRate: 6,
	    fundingPlan: {
      enabled: true,
      mode: "manual",
      yearlyAmounts: { "1": 300, "2": 0 },
    },
  }],
  mappings: [{
    ...singleSourceBlueprint.mappings[0],
    id: "mapping-funding-b",
    lineItemId: "funding-b",
  }],
};
const aggregateSources = [
  { sourceId: "amount-source-a", sourceName: "来源 A", blueprint: aggregateSourceA },
  { sourceId: "amount-source-b", sourceName: "来源 B", blueprint: aggregateSourceB },
];
assert.equal(validateIntelligentComputeSources(aggregateSources.slice(0, 1)).valid, true);
const aggregatePreview = buildIntelligentComputeAggregatePreview(
  aggregateSources,
  existingIctState,
  "project-1",
  {
    "cost:cost_it_other": {
      side: "cost",
      ictSubjectCode: "cost_it_other",
      amountInclTax: 25,
      taxRate: 6,
      yearlyAmounts: [25, 0, 0, 0, 0, 0, 0, 0, 0, 0],
      sourceLineItemIds: ["old-source:old-line"],
    },
  },
);
assert.equal(aggregatePreview.rows.length, ICT_SUBJECT_DEFINITIONS.length);
const aggregateDeviceRow = aggregatePreview.rows.find(row => row.ictSubjectCode === "cost_it_device");
assert.equal(aggregateDeviceRow.syncStatus, "ready");
assert.equal(aggregateDeviceRow.writtenAmount, 150);
assert.equal(aggregateDeviceRow.amountExclTax, 132.74);
assert.deepEqual(Array.from(aggregateDeviceRow.yearlyAmounts.slice(0, 2)), [100, 50]);
assert.deepEqual(
  Array.from(aggregateDeviceRow.sourceLineItemIds),
  ["amount-source-a:funding-a"],
);
const aggregateReleasedRow = aggregatePreview.rows.find(row => row.ictSubjectCode === "cost_it_other");
assert.equal(aggregateReleasedRow.syncStatus, "released_mapping");
assert.equal(aggregateReleasedRow.writtenAmount, 0);
const aggregateZeroedRow = aggregatePreview.rows.find(row => row.ictSubjectCode === "rev_it_integration");
assert.equal(aggregateZeroedRow.syncStatus, "zeroed_absent");
assert.equal(aggregateZeroedRow.writtenAmount, 0);
assert.equal(aggregateZeroedRow.taxRate, 6);

const intelligentTraceState = JSON.parse(JSON.stringify(existingIctState));
const intelligentTrace = {
  source: "intelligent_compute",
  sourceLabel: "来自智算金额来源",
  projectId: "project-1",
  scenarioId: "intelligent-compute-aggregate",
  blueprintId: "intelligent-compute-aggregate",
  sourceLineItemIds: ["old-source:old-line"],
  amountSourceIds: ["old-source"],
  sourceLineItems: [{ amountSourceId: "old-source", lineItemId: "old-line" }],
  importedAt: "2026-06-15T00:00:00.000Z",
};
intelligentTraceState.lifecycleState.inputPayloadJson.cost_it_other = {
  incl_tax: "25",
  excl_tax: "23.58",
  tax_rate: "6",
  import_trace: intelligentTrace,
};
intelligentTraceState.cashflowState.assumptionsJson.costIt.other = {
  incl: 25,
  excl: 23.58,
  tax: 6,
  importTrace: intelligentTrace,
};
intelligentTraceState.cashflowState.assumptionsJson.subjectFundingPlans["cost:costIt:other"] = {
  id: "cost:costIt:other",
  subjectRef: { side: "cost", groupId: "costIt", key: "other" },
  mode: "custom",
  annualInclValues: [25, 0, 0, 0, 0, 0, 0, 0, 0, 0],
  enabled: true,
  source: "intelligent_compute",
  lastChangeReason: "intelligent_compute_import",
  importTrace: intelligentTrace,
};
const aggregateTraceReleasePreview = buildIntelligentComputeAggregatePreview(
  aggregateSources,
  intelligentTraceState,
  "project-1",
);
const recoveredIntelligentRelease = aggregateTraceReleasePreview.rows.find(row =>
  row.ictSubjectCode === "cost_it_other"
);
assert.equal(recoveredIntelligentRelease.syncStatus, "released_mapping");
assert.equal(recoveredIntelligentRelease.writtenAmount, 0);

// 即使旧状态传入多个启用来源，正式同步也只使用第一个选中的来源。
const singleSelectedPreview = buildIntelligentComputeAggregatePreview(
  aggregateSources,
  existingIctState,
  "project-1",
);
assert.equal(
  singleSelectedPreview.rows.find(row => row.ictSubjectCode === "cost_it_device").writtenAmount,
  150,
);
const aggregatePayloads = buildAiComputeIctExportPayloads(aggregatePreview, existingIctState);
const aggregateTrace = aggregatePayloads.cashflowState.assumptionsJson
  .subjectFundingPlans["cost:costIt:device"].importTrace;
assert.deepEqual(Array.from(aggregateTrace.amountSourceIds), ["amount-source-a"]);
assert.deepEqual(
  Array.from(aggregateTrace.sourceLineItems, item => `${item.amountSourceId}:${item.lineItemId}`),
  ["amount-source-a:funding-a"],
);
assert.equal(aggregatePayloads.lifecycleState.inputPayloadJson.rev_it_integration.incl_tax, "0");
assert.deepEqual(
  Array.from(
    aggregatePayloads.cashflowState.assumptionsJson
      .subjectFundingPlans["revenue:revIt:integration"].annualInclValues,
  ),
  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
);
const aggregateFinalized = finalizeIctInputWithFundingPlans(
  aggregatePayloads.lifecycleState.inputPayloadJson,
  aggregatePayloads.cashflowState.assumptionsJson.subjectFundingPlans,
);
assert.equal(aggregateFinalized.coverage.valid, true);
assert.equal(
  aggregatePayloads.cashflowState.assumptionsJson.aiComputeQuoteImport.zeroedSubjectCodes.includes("rev_it_integration"),
  true,
);

const unmappedExportBlueprint = {
  ...fundingOutputBlueprint,
  costItems: [
    ...fundingOutputBlueprint.costItems,
    {
      ...item("unmapped-export", "未映射导出项", expression([{ type: "constant", value: 25 }])),
      fundingPlan: createDefaultAiComputeFundingPlan(25),
    },
  ],
};
const unmappedPreview = buildAiComputeIctExportPreview(
  unmappedExportBlueprint,
  existingIctState,
  "project-1",
);
assert.equal(unmappedPreview.skippedUnmappedItems.some(current => current.id === "unmapped-export"), true);

const originalDeviceCount = h200.parameters.find(parameter => parameter.id === "device-count").value;
const sensitivity = runAiComputeQuoteSensitivity(h200, {
  parameterId: "device-count",
  min: 32,
  max: 64,
  step: 16,
});
assert.deepEqual(Array.from(sensitivity, row => row.parameterValue), [32, 48, 64]);
assert.equal(h200.parameters.find(parameter => parameter.id === "device-count").value, originalDeviceCount);

console.log("AI compute quote tests passed.");
