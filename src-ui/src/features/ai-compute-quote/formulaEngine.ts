import Decimal from "decimal.js";
import type {
  AiComputeQuoteExpressionFormula,
  AiComputeQuoteFormula,
  AiComputeQuoteFormulaToken,
  AiComputeQuoteLegacyFormula,
  AiComputeQuoteLineItem,
  AiComputeQuoteParameter,
  FormulaEvaluationResult,
} from "./types";

interface FormulaValueContext {
  parameters: Map<string, AiComputeQuoteParameter>;
  lineItems: Map<string, AiComputeQuoteLineItem>;
  lineItemValues: Map<string, number>;
}

type ParseResult =
  | { ok: true; value: Decimal }
  | { ok: false; status: "incomplete" | "error"; message: string };

const OPERATOR_BY_LEGACY_TYPE: Record<AiComputeQuoteLegacyFormula["type"], "+" | "-" | "*" | "/"> = {
  add: "+",
  subtract: "-",
  multiply: "*",
  divide: "/",
};

const PRECEDENCE: Record<"+" | "-" | "*" | "/", number> = {
  "+": 1,
  "-": 1,
  "*": 2,
  "/": 2,
};

function money(value: Decimal.Value) {
  return new Decimal(value).toDecimalPlaces(2, Decimal.ROUND_HALF_UP).toNumber();
}

export function isExpressionFormula(formula: AiComputeQuoteFormula): formula is AiComputeQuoteExpressionFormula {
  return "version" in formula && formula.version === 2 && Array.isArray(formula.tokens);
}

export function normalizeQuoteFormula(
  formula: AiComputeQuoteFormula,
  parameters: AiComputeQuoteParameter[] = [],
): AiComputeQuoteExpressionFormula {
  if (isExpressionFormula(formula)) {
    return {
      version: 2,
      tokens: formula.tokens.map(token => ({ ...token })),
    };
  }

  const parameterNames = new Map(parameters.map(parameter => [parameter.id, parameter.name]));
  const operator = OPERATOR_BY_LEGACY_TYPE[formula.type];
  const tokens: AiComputeQuoteFormulaToken[] = [];
  formula.operands.forEach((operand, index) => {
    if (index > 0) tokens.push({ type: "operator", operator });
    if (operand.type === "constant") {
      tokens.push({ type: "constant", value: operand.value });
    } else {
      tokens.push({
        type: "parameter",
        id: operand.parameterId,
        name: parameterNames.get(operand.parameterId) || operand.parameterId,
      });
    }
  });
  return { version: 2, tokens };
}

export function getFormulaLineItemReferences(formula: AiComputeQuoteFormula) {
  if (!isExpressionFormula(formula)) return [];
  return Array.from(new Set(
    formula.tokens
      .filter((token): token is Extract<AiComputeQuoteFormulaToken, { type: "line_item" }> => token.type === "line_item")
      .map(token => token.id),
  ));
}

export function getFormulaParameterReferences(formula: AiComputeQuoteFormula) {
  if (!isExpressionFormula(formula)) {
    return Array.from(new Set(
      formula.operands
        .filter(operand => operand.type === "parameter")
        .map(operand => operand.parameterId),
    ));
  }
  return Array.from(new Set(
    formula.tokens
      .filter((token): token is Extract<AiComputeQuoteFormulaToken, { type: "parameter" }> => token.type === "parameter")
      .map(token => token.id),
  ));
}

export function describeQuoteFormula(
  formula: AiComputeQuoteFormula,
  parameters: AiComputeQuoteParameter[],
  lineItems: AiComputeQuoteLineItem[] = [],
) {
  const normalized = normalizeQuoteFormula(formula, parameters);
  const parameterNames = new Map(parameters.map(parameter => [parameter.id, parameter.name]));
  const lineItemNames = new Map(lineItems.map(item => [item.id, item.name]));
  return normalized.tokens.map(token => {
    if (token.type === "parameter") return `{${parameterNames.get(token.id) || token.name || token.id}}`;
    if (token.type === "line_item") return `{${lineItemNames.get(token.id) || token.name || token.id}}`;
    if (token.type === "constant") return String(token.value);
    if (token.type === "operator") return token.operator === "*" ? "×" : token.operator === "/" ? "÷" : token.operator;
    if (token.type === "left_parenthesis") return "(";
    if (token.type === "right_parenthesis") return ")";
    if (token.type === "function") return `${token.name}(`;
    return ",";
  }).join(" ");
}

class FormulaParser {
  private index = 0;
  readonly warnings: string[] = [];

  constructor(
    private readonly tokens: AiComputeQuoteFormulaToken[],
    private readonly context: FormulaValueContext,
  ) {}

  parse(): ParseResult {
    if (this.tokens.length === 0) {
      return { ok: false, status: "incomplete", message: "公式不完整：请插入参数、计算结果或固定值" };
    }
    const result = this.parseExpression(0);
    if (!result.ok) return result;
    if (this.index < this.tokens.length) {
      const token = this.tokens[this.index];
      return {
        ok: false,
        status: token.type === "right_parenthesis" || token.type === "comma" ? "error" : "incomplete",
        message: token.type === "right_parenthesis" ? "公式错误：存在多余的右括号" : "公式不完整：缺少运算符",
      };
    }
    return result;
  }

  private parseExpression(minPrecedence: number): ParseResult {
    let left = this.parsePrimary();
    if (!left.ok) return left;

    while (this.index < this.tokens.length) {
      const token = this.tokens[this.index];
      if (token.type !== "operator" || PRECEDENCE[token.operator] < minPrecedence) break;
      this.index += 1;
      const right = this.parseExpression(PRECEDENCE[token.operator] + 1);
      if (!right.ok) return right;
      if (token.operator === "/" && right.value.isZero()) {
        return { ok: false, status: "error", message: "公式错误：除数不能为 0" };
      }
      if (token.operator === "+") left = { ok: true, value: left.value.add(right.value) };
      if (token.operator === "-") left = { ok: true, value: left.value.sub(right.value) };
      if (token.operator === "*") left = { ok: true, value: left.value.mul(right.value) };
      if (token.operator === "/") left = { ok: true, value: left.value.div(right.value) };
    }
    return left;
  }

  private parsePrimary(): ParseResult {
    const token = this.tokens[this.index];
    if (!token) return { ok: false, status: "incomplete", message: "公式不完整：缺少计算对象" };

    if (token.type === "constant") {
      this.index += 1;
      if (!Number.isFinite(token.value)) {
        return { ok: false, status: "error", message: "公式错误：固定值不是有效数字" };
      }
      return { ok: true, value: new Decimal(token.value) };
    }

    if (token.type === "parameter") {
      this.index += 1;
      const parameter = this.context.parameters.get(token.id);
      if (!parameter) return { ok: false, status: "error", message: `公式错误：引用的参数“${token.name || token.id}”不存在` };
      if (!Number.isFinite(parameter.value)) return { ok: false, status: "error", message: `公式错误：参数“${parameter.name}”不是有效数字` };
      return { ok: true, value: new Decimal(parameter.value) };
    }

    if (token.type === "line_item") {
      this.index += 1;
      const lineItem = this.context.lineItems.get(token.id);
      if (!lineItem) return { ok: false, status: "error", message: `公式错误：引用的计算项“${token.name || token.id}”不存在` };
      if (!lineItem.enabled) {
        this.warnings.push(`计算项“${lineItem.name}”已禁用，按 0 参与计算`);
        return { ok: true, value: new Decimal(0) };
      }
      if (lineItem.calculationStatus && lineItem.calculationStatus !== "valid") {
        return { ok: false, status: "error", message: `公式错误：引用项“${lineItem.name}”计算失败` };
      }
      const value = this.context.lineItemValues.get(token.id);
      if (value === undefined) return { ok: false, status: "error", message: `公式错误：引用项“${lineItem.name}”尚未计算` };
      return { ok: true, value: new Decimal(value) };
    }

    if (token.type === "left_parenthesis") {
      this.index += 1;
      const result = this.parseExpression(0);
      if (!result.ok) return result;
      if (this.tokens[this.index]?.type !== "right_parenthesis") {
        return { ok: false, status: "incomplete", message: "公式不完整：缺少右括号" };
      }
      this.index += 1;
      return result;
    }

    if (token.type === "function" && token.name === "SUM") {
      this.index += 1;
      const values: Decimal[] = [];
      if (this.tokens[this.index]?.type === "right_parenthesis") {
        return { ok: false, status: "error", message: "公式错误：SUM 至少需要一个参数" };
      }
      while (this.index < this.tokens.length) {
        const value = this.parseExpression(0);
        if (!value.ok) return value;
        values.push(value.value);
        const separator = this.tokens[this.index];
        if (separator?.type === "comma") {
          this.index += 1;
          continue;
        }
        if (separator?.type === "right_parenthesis") {
          this.index += 1;
          return { ok: true, value: values.reduce((sum, item) => sum.add(item), new Decimal(0)) };
        }
        return { ok: false, status: "incomplete", message: "公式不完整：SUM 缺少逗号或右括号" };
      }
      return { ok: false, status: "incomplete", message: "公式不完整：SUM 缺少右括号" };
    }

    if (token.type === "operator") return { ok: false, status: "incomplete", message: "公式不完整：运算符前缺少计算对象" };
    if (token.type === "right_parenthesis") return { ok: false, status: "error", message: "公式错误：存在多余的右括号" };
    if (token.type === "comma") return { ok: false, status: "error", message: "公式错误：逗号位置无效" };
    return { ok: false, status: "error", message: "公式错误：无法识别的公式内容" };
  }
}

export function evaluateExpressionFormula(
  formula: AiComputeQuoteFormula,
  parameters: AiComputeQuoteParameter[],
  lineItems: AiComputeQuoteLineItem[] = [],
  lineItemValues: Map<string, number> = new Map(),
): FormulaEvaluationResult {
  const normalized = normalizeQuoteFormula(formula, parameters);
  const parser = new FormulaParser(normalized.tokens, {
    parameters: new Map(parameters.map(parameter => [parameter.id, parameter])),
    lineItems: new Map(lineItems.map(item => [item.id, item])),
    lineItemValues,
  });
  const result = parser.parse();
  if (!result.ok) {
    return { value: 0, errors: [result.message], warnings: parser.warnings, status: result.status };
  }
  if (!result.value.isFinite()) {
    return { value: 0, errors: ["公式错误：计算结果无效"], warnings: parser.warnings, status: "error" };
  }
  return { value: money(result.value), errors: [], warnings: parser.warnings, status: "valid" };
}
