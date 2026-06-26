import { calculateQuoteBlueprint } from "./calculations";
import {
  getFormulaParameterReferences,
  normalizeQuoteFormula,
} from "./formulaEngine";
import {
  isAiComputeDiscountRateParameter,
  isAiComputeProjectCycleParameter,
} from "./fundingPlans";
import type { IctResult } from "../../utils/projectService";
import type {
  AiComputeQuoteBlueprint,
  AiComputeQuoteFormula,
  AiComputeQuoteLineItem,
  AiComputeQuoteParameter,
  AiComputeQuoteSummary,
} from "./types";

export type QuoteNarrativeMeta = {
  projectName: string;
  sourceName?: string;
  projectYears: number;
  discountRatePercent: number;
};

function formatPlain(value: number) {
  if (!Number.isFinite(value)) return "0";
  return value.toLocaleString("zh-CN", { maximumFractionDigits: 4 });
}

function formatWan(value: number) {
  return (value / 10000).toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function formatPercentFromDecimal(value: number) {
  if (!Number.isFinite(value)) return "--";
  return (value * 100).toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

const OPERATOR_SYMBOL: Record<string, string> = {
  "*": "×",
  "/": "÷",
  "+": "+",
  "-": "-",
};

/**
 * 把公式渲染成「参数名(实际数值) × 参数名(实际数值) × 常量」的形式，
 * 让看不懂公式的同事能直观看出每个数字的来源。
 */
function renderFormulaWithValues(
  formula: AiComputeQuoteFormula,
  parameterMap: Map<string, AiComputeQuoteParameter>,
  itemNameMap: Map<string, string>,
  itemValueMap: Map<string, number>,
) {
  const normalized = normalizeQuoteFormula(formula, [...parameterMap.values()]);
  return normalized.tokens
    .map(token => {
      if (token.type === "parameter") {
        const parameter = parameterMap.get(token.id);
        if (!parameter) return token.name || token.id;
        const unit = parameter.unit ? parameter.unit : "";
        return `${parameter.name}(${formatPlain(parameter.value)}${unit})`;
      }
      if (token.type === "line_item") {
        const name = itemNameMap.get(token.id) || token.name || token.id;
        const value = itemValueMap.get(token.id);
        return value === undefined ? name : `${name}(${formatWan(value)}万元)`;
      }
      if (token.type === "constant") return formatPlain(token.value);
      if (token.type === "operator") return OPERATOR_SYMBOL[token.operator] || token.operator;
      if (token.type === "left_parenthesis") return "(";
      if (token.type === "right_parenthesis") return ")";
      if (token.type === "function") return `${token.name}(`;
      return "，";
    })
    .join(" ");
}

function describeConclusion(ictResult: IctResult | null) {
  if (!ictResult) return "尚未同步 ICT，下方为智算测算的预估口径，正式效益指标需同步后确认。";
  const margin = Number(ictResult.margin_rate);
  const npv = Number(ictResult.npv);
  if (npv < 0 || margin < 0) {
    return "存在负收益指标（NPV 或毛利率为负），建议复核收入、成本与资金计划后再决策。";
  }
  if (margin < 0.08) {
    return "收益空间较窄（毛利率低于 8%），建议重点关注主要成本构成。";
  }
  return "各项正式 ICT 指标均处于正向区间，整体收益良好。";
}

export function buildQuoteNarrative(
  blueprint: AiComputeQuoteBlueprint,
  summary: AiComputeQuoteSummary,
  ictResult: IctResult | null,
  meta: QuoteNarrativeMeta,
): string {
  const calculated = calculateQuoteBlueprint(blueprint);
  const parameterMap = new Map(calculated.parameters.map(parameter => [parameter.id, parameter]));
  const allItems = [...calculated.revenueItems, ...calculated.costItems];
  const itemNameMap = new Map(allItems.map(item => [item.id, item.name]));
  const itemValueMap = new Map(allItems.map(item => [item.id, item.amountInclTax]));

  const enabledRevenue = calculated.revenueItems.filter(item => item.enabled);
  const enabledCost = calculated.costItems.filter(item => item.enabled);

  // 关键参数：只列出被启用计算项实际引用的参数，排除已在抬头展示的周期与折现率。
  const referencedParameterIds = new Set<string>();
  [...enabledRevenue, ...enabledCost].forEach(item => {
    getFormulaParameterReferences(item.formula).forEach(id => referencedParameterIds.add(id));
  });
  const keyParameters = calculated.parameters.filter(
    parameter =>
      referencedParameterIds.has(parameter.id)
      && !isAiComputeProjectCycleParameter(parameter)
      && !isAiComputeDiscountRateParameter(parameter),
  );

  const renderItemLine = (item: AiComputeQuoteLineItem) => {
    const result = `${formatWan(item.amountInclTax)} 万元`;
    if (item.calculationStatus && item.calculationStatus !== "valid") {
      return `• ${item.name}：${item.calculationError || "公式未完成，暂按 0 计"}`;
    }
    const formula = renderFormulaWithValues(item.formula, parameterMap, itemNameMap, itemValueMap);
    return formula
      ? `• ${item.name} = ${formula} = ${result}`
      : `• ${item.name} = ${result}`;
  };

  const lines: string[] = [];

  const titleSource = meta.sourceName ? `（${meta.sourceName}）` : "";
  lines.push(`【智算测算说明 · ${meta.projectName || "未命名项目"}${titleSource}】`);
  lines.push(
    `本次按 ${formatPlain(meta.projectYears)} 年项目周期、${formatPlain(meta.discountRatePercent)}% 折现率测算。`,
  );
  lines.push("");

  lines.push("一、关键参数");
  if (keyParameters.length === 0) {
    lines.push("• （无）");
  } else {
    keyParameters.forEach(parameter => {
      lines.push(`• ${parameter.name}：${formatPlain(parameter.value)}${parameter.unit || ""}`);
    });
  }
  lines.push("");

  lines.push(`二、收入怎么算（合计 ${formatWan(summary.totalRevenue)} 万元，含税）`);
  if (enabledRevenue.length === 0) {
    lines.push("• （无启用的收入项）");
  } else {
    enabledRevenue.forEach(item => lines.push(renderItemLine(item)));
  }
  lines.push("");

  lines.push(`三、成本怎么算（合计 ${formatWan(summary.totalCost)} 万元，含税）`);
  if (enabledCost.length === 0) {
    lines.push("• （无启用的成本项）");
  } else {
    enabledCost.forEach(item => lines.push(renderItemLine(item)));
  }
  lines.push("");

  lines.push("四、效益结论");
  lines.push(
    `• 总收入 ${formatWan(summary.totalRevenue)} 万元，总成本 ${formatWan(summary.totalCost)} 万元（均含税）。`,
  );
  if (ictResult) {
    lines.push(`• ICT 毛利率：${formatPercentFromDecimal(Number(ictResult.margin_rate))}%`);
    lines.push(`• 净现值率：${formatPercentFromDecimal(Number(ictResult.npv_rate))}%`);
    lines.push(`• ICT NPV：${formatWan(Number(ictResult.npv))} 万元`);
    lines.push(`• 动态回收期：${ictResult.dynamic_payback} 年`);
  } else {
    lines.push(`• 智算毛利率（预估）：${formatPercentFromDecimal(summary.grossMarginRate / 100)}%`);
    lines.push("• NPV、净现值率、回收期等正式指标需同步 ICT 后查看。");
  }
  lines.push(`• 结论：${describeConclusion(ictResult)}`);
  lines.push("");
  lines.push("（说明：以上金额按含税口径，公式中括号内为各参数的实际取值。）");

  return lines.join("\n");
}
