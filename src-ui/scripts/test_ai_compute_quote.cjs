const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const ts = require("typescript");

const moduleCache = new Map();
function loadTs(relativePath) {
  const sourcePath = path.join(__dirname, "../src/features/ai-compute-quote", relativePath);
  if (moduleCache.has(sourcePath)) return moduleCache.get(sourcePath).exports;
  const source = fs.readFileSync(sourcePath, "utf8");
  const transpiled = ts.transpileModule(source, {
    compilerOptions: {
      esModuleInterop: true,
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
    },
  });
  const moduleRef = { exports: {} };
  moduleCache.set(sourcePath, moduleRef);
  const localRequire = request => {
    if (request.startsWith("./")) return loadTs(`${request.slice(2)}.ts`);
    return require(request);
  };
  vm.runInNewContext(transpiled.outputText, {
    module: moduleRef,
    exports: moduleRef.exports,
    require: localRequire,
  }, { filename: sourcePath });
  return moduleRef.exports;
}

const {
  buildAiComputeQuoteOutput,
  calculateQuoteBlueprint,
  evaluateQuoteFormula,
  runAiComputeQuoteSensitivity,
  summarizeQuote,
} = loadTs("calculations.ts");
const { createH200Blueprint } = loadTs("presets.ts");
const {
  clampFormulaCursor,
  insertFormulaTokensAt,
  removeFormulaTokenAt,
  removeFormulaTokenBeforeCursor,
} = loadTs("formulaTokenEditing.ts");

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
assert.equal(summary.grossProfit, 80568000);
assert.equal(summary.grossMarginRate, 17.95);

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
