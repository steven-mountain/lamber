import Decimal from "decimal.js";
import {
  describeQuoteFormula,
  evaluateExpressionFormula,
  getFormulaLineItemReferences,
} from "./formulaEngine";
import type {
  AiComputeQuoteBlueprint,
  AiComputeQuoteFormula,
  AiComputeQuoteLineItem,
  AiComputeQuoteOutputSubjectAmount,
  AiComputeQuoteParameter,
  AiComputeQuoteSensitivityConfig,
  AiComputeQuoteSensitivityRow,
  AiComputeQuoteSummary,
  FormulaEvaluationResult,
} from "./types";

const ZERO = new Decimal(0);

function finiteNumber(value: unknown) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function money(value: Decimal.Value) {
  return new Decimal(value).toDecimalPlaces(2, Decimal.ROUND_HALF_UP).toNumber();
}

export function evaluateQuoteFormula(
  formula: AiComputeQuoteFormula,
  parameters: AiComputeQuoteParameter[],
  lineItems: AiComputeQuoteLineItem[] = [],
  lineItemValues: Map<string, number> = new Map(),
): FormulaEvaluationResult {
  return evaluateExpressionFormula(formula, parameters, lineItems, lineItemValues);
}

function detectCircularReferences(lineItems: AiComputeQuoteLineItem[]) {
  const itemMap = new Map(lineItems.map(item => [item.id, item]));
  const visitState = new Map<string, "visiting" | "visited">();
  const stack: string[] = [];
  const cycleErrors = new Map<string, string>();

  const visit = (itemId: string) => {
    if (visitState.get(itemId) === "visited") return;
    if (visitState.get(itemId) === "visiting") {
      const start = stack.indexOf(itemId);
      const cycleIds = [...stack.slice(Math.max(0, start)), itemId];
      const cycleNames = cycleIds.map(id => itemMap.get(id)?.name || id);
      const message = `循环引用：${cycleNames.join(" → ")}`;
      cycleIds.forEach(id => cycleErrors.set(id, message));
      return;
    }

    const item = itemMap.get(itemId);
    if (!item) return;
    visitState.set(itemId, "visiting");
    stack.push(itemId);
    getFormulaLineItemReferences(item.formula)
      .filter(referenceId => itemMap.has(referenceId))
      .forEach(visit);
    stack.pop();
    visitState.set(itemId, "visited");
  };

  lineItems.forEach(item => visit(item.id));
  return cycleErrors;
}

function calculateExclTax(amountInclTax: number, taxRateValue: unknown) {
  const taxRate = Math.max(0, finiteNumber(taxRateValue));
  return money(new Decimal(amountInclTax).div(new Decimal(1).add(new Decimal(taxRate).div(100))));
}

export function calculateQuoteLineItem(
  item: AiComputeQuoteLineItem,
  parameters: AiComputeQuoteParameter[],
  lineItems: AiComputeQuoteLineItem[] = [],
  lineItemValues: Map<string, number> = new Map(),
): AiComputeQuoteLineItem {
  if (!item.enabled) {
    return {
      ...item,
      amountInclTax: 0,
      amountExclTax: 0,
      calculationStatus: "valid",
      calculationError: undefined,
      calculationWarnings: ["当前计算项已禁用，金额按 0 处理"],
    };
  }

  const evaluated = evaluateQuoteFormula(item.formula, parameters, lineItems, lineItemValues);
  const amountInclTax = evaluated.status === "valid" ? evaluated.value : 0;
  return {
    ...item,
    amountInclTax,
    amountExclTax: calculateExclTax(amountInclTax, item.taxRate),
    calculationStatus: evaluated.status,
    calculationError: evaluated.errors[0],
    calculationWarnings: evaluated.warnings,
  };
}

export function calculateQuoteBlueprint(blueprint: AiComputeQuoteBlueprint): AiComputeQuoteBlueprint {
  const sourceItems = [...blueprint.revenueItems, ...blueprint.costItems];
  const itemMap = new Map(sourceItems.map(item => [item.id, item]));
  const cycleErrors = detectCircularReferences(sourceItems);
  const calculatedItems = new Map<string, AiComputeQuoteLineItem>();
  const lineItemValues = new Map<string, number>();

  const calculateItem = (itemId: string): AiComputeQuoteLineItem | undefined => {
    const cached = calculatedItems.get(itemId);
    if (cached) return cached;
    const source = itemMap.get(itemId);
    if (!source) return undefined;

    const cycleError = cycleErrors.get(itemId);
    if (cycleError) {
      const failed = {
        ...source,
        amountInclTax: 0,
        amountExclTax: 0,
        calculationStatus: "error" as const,
        calculationError: cycleError,
        calculationWarnings: [],
      };
      calculatedItems.set(itemId, failed);
      lineItemValues.set(itemId, 0);
      return failed;
    }

    getFormulaLineItemReferences(source.formula).forEach(calculateItem);
    const contextItems = sourceItems.map(item => calculatedItems.get(item.id) || item);
    const calculated = calculateQuoteLineItem(source, blueprint.parameters, contextItems, lineItemValues);
    calculatedItems.set(itemId, calculated);
    lineItemValues.set(itemId, calculated.amountInclTax);
    return calculated;
  };

  sourceItems.forEach(item => calculateItem(item.id));
  return {
    ...blueprint,
    revenueItems: blueprint.revenueItems.map(item => calculatedItems.get(item.id) || item),
    costItems: blueprint.costItems.map(item => calculatedItems.get(item.id) || item),
  };
}

export function summarizeQuote(blueprint: AiComputeQuoteBlueprint): AiComputeQuoteSummary {
  const calculated = calculateQuoteBlueprint(blueprint);
  const totalRevenue = money(calculated.revenueItems.reduce((sum, item) => sum.add(item.amountInclTax), ZERO));
  const totalCost = money(calculated.costItems.reduce((sum, item) => sum.add(item.amountInclTax), ZERO));
  const grossProfit = money(new Decimal(totalRevenue).sub(totalCost));
  const grossMarginRate = totalRevenue === 0
    ? 0
    : new Decimal(grossProfit).div(totalRevenue).mul(100).toDecimalPlaces(2).toNumber();
  const deviceCount = finiteNumber(calculated.parameters.find(parameter => parameter.key === "device_count")?.value);
  const years = finiteNumber(calculated.parameters.find(parameter => parameter.key === "years")?.value);
  const denominator = new Decimal(deviceCount).mul(years).mul(12);
  const costPerDeviceMonth = denominator.lte(0)
    ? 0
    : money(new Decimal(totalCost).div(denominator));

  return { totalRevenue, totalCost, grossProfit, grossMarginRate, costPerDeviceMonth };
}

export function buildAiComputeQuoteOutput(
  blueprint: AiComputeQuoteBlueprint,
): AiComputeQuoteOutputSubjectAmount[] {
  const calculated = calculateQuoteBlueprint(blueprint);
  const lineItems = [...calculated.revenueItems, ...calculated.costItems];
  const lineItemMap = new Map(lineItems.map(item => [item.id, item]));
  const grouped = new Map<string, AiComputeQuoteOutputSubjectAmount>();

  calculated.mappings.forEach(mapping => {
    const item = lineItemMap.get(mapping.lineItemId);
    if (
      !mapping.enabled
      || !mapping.ictSubjectCode
      || !item
      || !item.enabled
      || !item.outputEnabled
      || item.side !== mapping.side
      || item.calculationStatus !== "valid"
    ) {
      return;
    }

    const key = `${mapping.side}:${mapping.ictSubjectCode}`;
    const existing = grouped.get(key);
    if (existing) {
      existing.amountInclTax = money(new Decimal(existing.amountInclTax).add(item.amountInclTax));
      existing.amountExclTax = money(new Decimal(existing.amountExclTax).add(item.amountExclTax));
      existing.sourceLineItemIds.push(item.id);
      return;
    }

    grouped.set(key, {
      side: mapping.side,
      ictSubjectCode: mapping.ictSubjectCode,
      ictSubjectName: mapping.ictSubjectName,
      amountInclTax: item.amountInclTax,
      amountExclTax: item.amountExclTax,
      sourceLineItemIds: [item.id],
    });
  });

  return Array.from(grouped.values());
}

export function runAiComputeQuoteSensitivity(
  blueprint: AiComputeQuoteBlueprint,
  config: AiComputeQuoteSensitivityConfig,
): AiComputeQuoteSensitivityRow[] {
  if (!Number.isFinite(config.min) || !Number.isFinite(config.max) || !Number.isFinite(config.step)) return [];
  if (config.step <= 0 || config.max < config.min) return [];
  if (!blueprint.parameters.some(parameter => parameter.id === config.parameterId)) return [];

  const rows: AiComputeQuoteSensitivityRow[] = [];
  const min = new Decimal(config.min);
  const max = new Decimal(config.max);
  const step = new Decimal(config.step);
  let current = min;
  let guard = 0;

  while (current.lte(max) && guard < 500) {
    const parameterValue = current.toNumber();
    const candidate: AiComputeQuoteBlueprint = {
      ...blueprint,
      parameters: blueprint.parameters.map(parameter =>
        parameter.id === config.parameterId ? { ...parameter, value: parameterValue } : parameter
      ),
    };
    rows.push({ parameterValue, ...summarizeQuote(candidate) });
    current = current.add(step);
    guard += 1;
  }

  return rows;
}

export { describeQuoteFormula };
