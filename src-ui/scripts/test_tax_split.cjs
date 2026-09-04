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
  return loadTsFile(path.join(__dirname, "../src/lib", relativePath));
}

const {
  splitInclAmount,
  splitInclAmountForExclDelta,
  isInclRepresentable,
  restoreTaxSplitParts,
  resolveSerializedTaxSplitParts,
  serializeTaxSplitParts,
} = loadTs("taxAmount.ts");
const { validateFinancialData } = loadTs("financeValidator.ts");
const {
  buildPptTaxAmountRows,
  formatPptSplitNote,
  getPptTaxSplitSummary,
} = loadTs("pptTaxRows.ts");

// 独立的整数分口径对照实现（不经 decimal.js），作为拆分正确性的 oracle。
const rdiv = (n, d) => Math.floor((2 * n + d) / (2 * d)); // round(n/d) 半进位，n,d>0
const centsSelfConsistent = (cents, rate) => {
  const excl = rdiv(cents * 100, 100 + rate);
  return rdiv(excl * (100 + rate), 100) === cents;
};
const toCents = value => Math.round(value * 100);

const RATES = [1, 3, 5, 6, 9, 13];

// --- 1. 典型案例：240 元 @6% 拆成 120 + 120，各笔不含税 113.21 ---
{
  const parts = splitInclAmount(240, 6);
  assert.ok(parts, "240@6% 应可拆分");
  assert.equal(parts.length, 2);
  assert.equal(JSON.stringify(parts.map(p => p.incl)), JSON.stringify([120, 120]));
  assert.equal(JSON.stringify(parts.map(p => p.excl)), JSON.stringify([113.21, 113.21]));
}

// --- 2. 自洽金额与边界值不拆分 ---
assert.equal(splitInclAmount(100, 6), null, "100@6% 自洽，无需拆分");
assert.equal(splitInclAmount(0, 6), null);
assert.equal(splitInclAmount(0.01, 6), null);
assert.equal(isInclRepresentable(100, 6), true);
assert.equal(isInclRepresentable(240, 6), false);

// --- 3. 税率组尾差定向拆分：原科目已闭合，也可搜索恰好 ±0.01 的拆分建议 ---
{
  const parts = splitInclAmountForExclDelta(0.10, 6, 0.09, 0.01);
  assert.ok(parts, "0.10@6% 应存在不含税增加 0.01 的两笔拆分建议");
  assert.equal(parts.reduce((sum, part) => sum + toCents(part.incl), 0), 10);
  assert.equal(parts.reduce((sum, part) => sum + toCents(part.excl), 0), 10);
  for (const part of parts) {
    assert.ok(centsSelfConsistent(toCents(part.incl), 6), "建议中的每笔金额必须双向闭合");
  }
  assert.equal(splitInclAmountForExclDelta(0.10, 6, 0.09, -0.01), null);

  const negativeParts = splitInclAmountForExclDelta(0.20, 6, 0.19, -0.01);
  assert.ok(negativeParts, "0.20@6% 应存在不含税减少 0.01 的两笔拆分建议");
  assert.equal(negativeParts.reduce((sum, part) => sum + toCents(part.incl), 0), 20);
  assert.equal(negativeParts.reduce((sum, part) => sum + toCents(part.excl), 0), 18);
}

// --- 4. 穷举 + 抽样验证：拆分结果每笔自洽、和等于总额、两笔差不超过 1 分 ---
for (const rate of RATES) {
  const check = totalCents => {
    const total = totalCents / 100;
    const parts = splitInclAmount(total, rate);
    if (centsSelfConsistent(totalCents, rate)) {
      assert.equal(parts, null, `自洽金额 ${total}@${rate}% 不应拆分`);
      return;
    }
    assert.ok(parts, `不自洽金额 ${total}@${rate}% 必须拆分成功`);
    const centsParts = parts.map(p => toCents(p.incl));
    assert.equal(centsParts.reduce((a, b) => a + b, 0), totalCents, `${total}@${rate}% 拆分之和必须等于总额`);
    assert.ok(Math.abs(centsParts[0] - centsParts[1]) <= 1, `${total}@${rate}% 两笔应最接近对半`);
    for (const part of parts) {
      assert.ok(centsSelfConsistent(toCents(part.incl), rate), `${total}@${rate}% 子笔 ${part.incl} 必须自洽`);
      assert.equal(part.excl, rdiv(toCents(part.incl) * 100, 100 + rate) / 100, `${total}@${rate}% 子笔不含税口径`);
    }
  };
  // 穷举 0.02 ～ 2000 元
  for (let cents = 2; cents <= 200000; cents++) check(cents);
  // 抽样 2000 ～ 100 万元
  for (let i = 0; i < 20000; i++) {
    check(200000 + Math.floor(Math.random() * (100000000 - 200000)));
  }
}
console.log("splitInclAmount: 穷举与抽样验证通过");

// --- 5. restoreTaxSplitParts：合法明细还原，非法明细整体丢弃 ---
{
  const restored = restoreTaxSplitParts(
    [{ incl_tax: "120", excl_tax: "113.21" }, { incl_tax: "120", excl_tax: "113.21" }],
    240,
    6,
  );
  assert.ok(restored);
  assert.equal(JSON.stringify(restored.map(p => p.incl)), JSON.stringify([120, 120]));
  assert.equal(JSON.stringify(restored.map(p => p.excl)), JSON.stringify([113.21, 113.21]));
  assert.deepEqual(
    JSON.parse(JSON.stringify(serializeTaxSplitParts(restored))),
    [
      { incl_tax: "120.00", excl_tax: "113.21" },
      { incl_tax: "120.00", excl_tax: "113.21" },
    ],
  );

  // cashflow assumptions 没有拆分时，必须保留 lifecycle/snapshot 中的有效拆分。
  assert.deepEqual(
    JSON.parse(JSON.stringify(resolveSerializedTaxSplitParts(
      undefined,
      [{ incl_tax: "120.00" }, { incl_tax: "120.00" }],
      240,
      6,
    ))),
    [
      { incl_tax: "120.00", excl_tax: "113.21" },
      { incl_tax: "120.00", excl_tax: "113.21" },
    ],
  );

  // 合计不等于科目含税 → 丢弃
  assert.equal(restoreTaxSplitParts([{ incl_tax: "120" }, { incl_tax: "119.99" }], 240, 6), null);
  // 子笔不自洽（240 本身）→ 丢弃
  assert.equal(restoreTaxSplitParts([{ incl_tax: "240" }, { incl_tax: "240" }], 480, 6), null);
  // 非数组 / 单笔 → 丢弃
  assert.equal(restoreTaxSplitParts(undefined, 240, 6), null);
  assert.equal(restoreTaxSplitParts([{ incl_tax: "240" }], 240, 6), null);
}
console.log("restoreTaxSplitParts: 通过");

// --- 6. financeValidator：拆分科目免于合计反推核验，未拆分时保持原报错 ---
{
  const splitItem = {
    incl: 240,
    excl: 226.42,
    tax: 6,
    splitParts: [
      { incl: 120, excl: 113.21 },
      { incl: 120, excl: 113.21 },
    ],
  };
  const { errors } = validateFinancialData({ ct: { product: splitItem } }, {});
  assert.equal(errors.length, 0, `拆分科目不应报错，实际：${JSON.stringify(errors)}`);

  // 同数据不带拆分：B2（含税 ≠ 不含税反推）必须照常报错，防止回归
  const plainItem = { incl: 240, excl: 226.42, tax: 6 };
  const plain = validateFinancialData({ ct: { product: plainItem } }, {});
  assert.equal(plain.errors.length, 1);
  assert.equal(plain.errors[0].field, "incl");

  // 拆分组跳过 C1（round-of-sum 与逐行锚点差 1 分的场景）：0.09@6% → 0.04 + 0.05
  const tinySplit = {
    incl: 0.09,
    excl: 0.09,
    tax: 6,
    splitParts: [
      { incl: 0.04, excl: 0.04 },
      { incl: 0.05, excl: 0.05 },
    ],
  };
  const tiny = validateFinancialData({ ct: { product: tinySplit } }, {});
  assert.equal(tiny.errors.length, 0, `拆分组应跳过 C1，实际：${JSON.stringify(tiny.errors)}`);

  // 拆分明细与合计不符（状态被破坏）时仍能被 C2 拦截
  const brokenSplit = {
    incl: 240,
    excl: 226.43,
    tax: 6,
    splitParts: [
      { incl: 120, excl: 113.21 },
      { incl: 120, excl: 113.22 },
    ],
  };
  const broken = validateFinancialData({ ct: { product: brokenSplit } }, {});
  assert.ok(broken.errors.length > 0, "被破坏的拆分明细应报错");
}
console.log("financeValidator: 通过");

// --- 7. 税率组汇总尾差：只提供精确归零建议，不自动修改原科目 ---
{
  const revenue = {
    it: {
      integration: { incl: 0.10, excl: 0.09, tax: 6 },
      maintenance: { incl: 0.10, excl: 0.09, tax: 6 },
    },
  };
  const result = validateFinancialData(revenue, {});
  assert.equal(result.errors.length, 1);
  assert.equal(result.errors[0].key, "[汇总误差-公式C1]");
  assert.equal(result.errors[0].difference, "-0.01");
  assert.ok(result.errors[0].splitSuggestions?.length, "汇总尾差应附带拆分建议");

  const suggestion = result.errors[0].splitSuggestions[0];
  assert.equal(suggestion.differenceBefore, "-0.01");
  assert.equal(suggestion.differenceAfter, "0.00");
  assert.equal(suggestion.exclAdjustment, "0.01");
  assert.equal(suggestion.parts.reduce((sum, part) => sum + toCents(part.incl), 0), 10);
  assert.equal(suggestion.parts.reduce((sum, part) => sum + toCents(part.excl), 0), 10);

  // 校验器只返回建议，输入对象保持原值。
  assert.equal(revenue.it.integration.incl, 0.10);
  assert.equal(revenue.it.integration.excl, 0.09);
  assert.equal(revenue.it.maintenance.incl, 0.10);
  assert.equal(revenue.it.maintenance.excl, 0.09);

  // 用户点击“应用”后的正式状态必须通过同一个 0 容差校验，警示不会再次出现。
  const appliedRevenue = JSON.parse(JSON.stringify(revenue));
  const targetKey = suggestion.subjectKey.split(".")[1];
  appliedRevenue.it[targetKey] = {
    ...appliedRevenue.it[targetKey],
    excl: suggestion.parts.reduce((sum, part) => sum + part.excl, 0),
    splitParts: suggestion.parts,
  };
  const appliedResult = validateFinancialData(appliedRevenue, {});
  assert.equal(appliedResult.errors.length, 0, JSON.stringify(appliedResult.errors));
}
console.log("tax-group split suggestion: 通过");

// --- 8. PPT 明细展示：默认合并，用户选择后才展开拆分子笔 ---
{
  const splitItem = {
    incl: 1038,
    excl: 979.24,
    tax: 6,
    splitParts: [
      { incl: 519, excl: 489.62 },
      { incl: 519, excl: 489.62 },
    ],
  };

  const mergedRows = buildPptTaxAmountRows(splitItem, "merged");
  assert.equal(mergedRows.length, 1);
  assert.equal(mergedRows[0].incl, 1038);
  assert.equal(mergedRows[0].excl, 979.24);
  assert.equal(formatPptSplitNote(mergedRows[0]), "");

  const splitRows = buildPptTaxAmountRows(splitItem, "split");
  assert.equal(splitRows.length, 2);
  assert.equal(splitRows.reduce((sum, row) => sum + toCents(row.incl), 0), 103800);
  assert.equal(splitRows.reduce((sum, row) => sum + toCents(row.excl), 0), 97924);
  assert.equal(formatPptSplitNote(splitRows[0]), "拆分第1笔/共2笔");
  assert.equal(formatPptSplitNote(splitRows[1]), "拆分第2笔/共2笔");

  const summary = getPptTaxSplitSummary({
    revenue: { it: { integration: splitItem } },
    // costMix 当前不在 PPT 的四张明细表中，不应让选项误报数量。
    cost: { mix: { marketing: splitItem } },
  });
  assert.equal(summary.subjectCount, 1);
  assert.equal(summary.addedRows, 1);

  const damagedRows = buildPptTaxAmountRows({
    ...splitItem,
    splitParts: [{ incl: 519 }, { incl: 518.99 }],
  }, "split");
  assert.equal(damagedRows.length, 1, "损坏拆分必须安全回退为科目汇总行");
  assert.equal(damagedRows[0].excl, 979.24);
}
console.log("PPT tax split rows: 通过");

console.log("\n全部测试通过 ✅");
