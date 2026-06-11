import { useEffect, useMemo, useState } from "react";
import { Button } from "../../components/ui/button";
import { Input } from "../../components/ui/input";
import {
  clampFormulaCursor,
  insertFormulaTokensAt,
  removeFormulaTokenAt,
  removeFormulaTokenBeforeCursor,
} from "./formulaTokenEditing";
import type {
  AiComputeQuoteBlueprint,
  AiComputeQuoteExpressionFormula,
  AiComputeQuoteFormulaToken,
  AiComputeQuoteLineItem,
} from "./types";

interface QuoteFormulaCalculatorProps {
  blueprint: AiComputeQuoteBlueprint;
  item: AiComputeQuoteLineItem;
  onChange: (formula: AiComputeQuoteExpressionFormula) => void;
}

function tokenLabel(
  token: AiComputeQuoteFormulaToken,
  blueprint: AiComputeQuoteBlueprint,
) {
  if (token.type === "parameter") {
    return blueprint.parameters.find(parameter => parameter.id === token.id)?.name || token.name || token.id;
  }
  if (token.type === "line_item") {
    return [...blueprint.revenueItems, ...blueprint.costItems].find(item => item.id === token.id)?.name || token.name || token.id;
  }
  if (token.type === "constant") return String(token.value);
  if (token.type === "operator") return token.operator === "*" ? "×" : token.operator === "/" ? "÷" : token.operator;
  if (token.type === "left_parenthesis") return "(";
  if (token.type === "right_parenthesis") return ")";
  if (token.type === "function") return "SUM(";
  return ",";
}

function tokenTone(token: AiComputeQuoteFormulaToken) {
  if (token.type === "parameter") return "bg-success-soft text-success-foreground";
  if (token.type === "line_item") return "bg-primary-soft text-primary";
  if (token.type === "constant") return "bg-warning-soft text-warning-foreground";
  return "bg-secondary text-foreground";
}

function FormulaPreview({
  blueprint,
  tokens,
}: {
  blueprint: AiComputeQuoteBlueprint;
  tokens: AiComputeQuoteFormulaToken[];
}) {
  if (tokens.length === 0) return <span className="text-warning-foreground">公式不完整</span>;

  return (
    <span className="inline-flex flex-wrap items-center gap-x-1.5 gap-y-1">
      {tokens.map((token, index) => {
        const label = tokenLabel(token, blueprint);
        if (token.type === "parameter" || token.type === "line_item") {
          return (
            <code
              key={`preview-${index}`}
              className={`rounded-md px-1.5 py-0.5 font-sans text-[0.94em] font-semibold ${
                token.type === "parameter"
                  ? "bg-success-soft text-success-foreground"
                  : "bg-primary-soft text-primary"
              }`}
            >
              {label}
            </code>
          );
        }
        if (token.type === "constant") {
          return (
            <code
              key={`preview-${index}`}
              className="rounded-md bg-warning-soft px-1.5 py-0.5 font-mono text-[0.94em] font-semibold text-warning-foreground"
            >
              {label}
            </code>
          );
        }
        return <span key={`preview-${index}`}>{label}</span>;
      })}
    </span>
  );
}

function FormulaCursor({
  active,
  position,
  onSelect,
}: {
  active: boolean;
  position: number;
  onSelect: () => void;
}) {
  return (
    <button
      type="button"
      aria-label={`将公式插入点移动到位置 ${position + 1}`}
      className="group flex h-8 w-3 shrink-0 items-center justify-center rounded-sm"
      title="点击移动插入点"
      onClick={onSelect}
    >
      <span className={`h-6 rounded-full transition-all ${
        active
          ? "w-0.5 bg-primary shadow-[0_0_0_2px_hsl(var(--primary-soft))]"
          : "w-px bg-transparent group-hover:bg-primary/45"
      }`} />
    </button>
  );
}

export default function QuoteFormulaCalculator({
  blueprint,
  item,
  onChange,
}: QuoteFormulaCalculatorProps) {
  const formula = item.formula as AiComputeQuoteExpressionFormula;
  const [parameterId, setParameterId] = useState("");
  const [lineItemId, setLineItemId] = useState("");
  const [constantValue, setConstantValue] = useState("1");
  const [cursor, setCursor] = useState(formula.tokens.length);
  const allLineItems = useMemo(
    () => [...blueprint.revenueItems, ...blueprint.costItems],
    [blueprint.costItems, blueprint.revenueItems],
  );

  useEffect(() => {
    setCursor(current => clampFormulaCursor(current, formula.tokens.length));
  }, [formula.tokens.length]);

  const setTokens = (tokens: AiComputeQuoteFormulaToken[], nextCursor = cursor) => {
    setCursor(clampFormulaCursor(nextCursor, tokens.length));
    onChange({ version: 2, tokens });
  };
  const insertAtCursor = (...tokens: AiComputeQuoteFormulaToken[]) => {
    const next = insertFormulaTokensAt(formula.tokens, cursor, tokens);
    setTokens(next.tokens, next.cursor);
  };
  const deleteToken = (tokenIndex: number) => {
    const next = removeFormulaTokenAt(formula.tokens, tokenIndex, cursor);
    setTokens(next.tokens, next.cursor);
  };
  const backspace = () => {
    const next = removeFormulaTokenBeforeCursor(formula.tokens, cursor);
    setTokens(next.tokens, next.cursor);
  };
  const insertParameter = () => {
    const parameter = blueprint.parameters.find(candidate => candidate.id === parameterId);
    if (!parameter) return;
    insertAtCursor({ type: "parameter", id: parameter.id, name: parameter.name });
    setParameterId("");
  };
  const insertLineItem = () => {
    const lineItem = allLineItems.find(candidate => candidate.id === lineItemId);
    if (!lineItem) return;
    insertAtCursor({ type: "line_item", id: lineItem.id, name: lineItem.name });
    setLineItemId("");
  };
  const insertConstant = () => {
    const value = Number(constantValue);
    if (!Number.isFinite(value)) return;
    insertAtCursor({ type: "constant", value });
  };

  const statusLabel = item.calculationStatus === "valid"
    ? `计算结果：${item.amountInclTax.toLocaleString("zh-CN", { maximumFractionDigits: 2 })} 元`
    : item.calculationError || "公式不完整";
  const statusClass = item.calculationStatus === "valid"
    ? "text-success-foreground"
    : item.calculationStatus === "error"
      ? "text-destructive"
      : "text-warning-foreground";

  return (
    <div className="rounded-lg bg-card/85 p-3 shadow-sm">
      <div className="mb-3 rounded-lg bg-muted/55 px-3 py-3">
        <div className="text-caption text-secondary-foreground">公式预览</div>
        <div className="mt-1 flex flex-wrap items-center gap-1.5 break-words text-body-strong">
          <span>{item.name}</span>
          <span>=</span>
          <FormulaPreview blueprint={blueprint} tokens={formula.tokens} />
        </div>
        <div className={`mt-2 numeric-value text-body-strong ${statusClass}`}>{statusLabel}</div>
        {item.calculationWarnings?.map(warning => (
          <div key={warning} className="mt-1 text-caption text-warning-foreground">{warning}</div>
        ))}
      </div>

      <div className="min-h-11 rounded-lg bg-muted/55 p-2">
        <div className="flex flex-wrap items-center gap-y-2">
          <FormulaCursor active={cursor === 0} position={0} onSelect={() => setCursor(0)} />
          {formula.tokens.map((token, index) => (
            <div key={`${item.id}-token-${index}`} className="contents">
              <div className={`group/token flex items-center rounded-md text-caption font-semibold ${tokenTone(token)}`}>
                <button
                  type="button"
                  className="px-2.5 py-1"
                  title="点击将插入点移动到该片段之后"
                  onClick={() => setCursor(index + 1)}
                >
                  {tokenLabel(token, blueprint)}
                </button>
                <button
                  type="button"
                  aria-label={`删除公式片段 ${tokenLabel(token, blueprint)}`}
                  className="mr-1 rounded px-1 text-current opacity-0 transition-opacity hover:bg-card/60 group-hover/token:opacity-100 focus-visible:opacity-100 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/30"
                  title="删除该公式片段"
                  onClick={() => deleteToken(index)}
                >
                  ×
                </button>
              </div>
              <FormulaCursor
                active={cursor === index + 1}
                position={index + 1}
                onSelect={() => setCursor(index + 1)}
              />
            </div>
          ))}
          {formula.tokens.length === 0 && (
            <span className="px-2 text-caption text-secondary-foreground">公式为空，当前插入点位于开头</span>
          )}
        </div>
        <div className="mt-1 px-1 text-[10px] text-secondary-foreground">
          插入点：第 {cursor + 1} 个位置，共 {formula.tokens.length + 1} 个位置
        </div>
      </div>

      <div className="mt-3 grid gap-3 xl:grid-cols-2">
        <div className="rounded-lg bg-muted/35 p-3">
          <div className="mb-2 text-caption font-bold text-secondary-foreground">插入参数</div>
          <div className="flex gap-2">
            <select
              className="h-9 min-w-0 flex-1 rounded-md border border-input bg-card px-2 text-sm"
              value={parameterId}
              onChange={event => setParameterId(event.target.value)}
            >
              <option value="">选择参数区参数</option>
              {blueprint.parameters.map(parameter => <option key={parameter.id} value={parameter.id}>{parameter.name}</option>)}
            </select>
            <Button size="sm" variant="secondary" disabled={!parameterId} onClick={insertParameter}>插入</Button>
          </div>
        </div>

        <div className="rounded-lg bg-muted/35 p-3">
          <div className="mb-2 text-caption font-bold text-secondary-foreground">插入已计算结果</div>
          <div className="flex gap-2">
            <select
              className="h-9 min-w-0 flex-1 rounded-md border border-input bg-card px-2 text-sm"
              value={lineItemId}
              onChange={event => setLineItemId(event.target.value)}
            >
              <option value="">选择其他收入/成本项</option>
              {allLineItems.filter(candidate => candidate.id !== item.id).map(candidate => (
                <option key={candidate.id} value={candidate.id}>
                  {candidate.side === "revenue" ? "收入" : "成本"} · {candidate.name}
                </option>
              ))}
            </select>
            <Button size="sm" variant="secondary" disabled={!lineItemId} onClick={insertLineItem}>插入</Button>
          </div>
        </div>
      </div>

      <div className="mt-3 flex flex-wrap items-end gap-3 rounded-lg bg-muted/35 p-3">
        <div>
          <div className="mb-2 text-caption font-bold text-secondary-foreground">运算符</div>
          <div className="flex gap-1">
            {(["+", "-", "*", "/"] as const).map(operator => (
              <Button key={operator} size="sm" variant="outline" onClick={() => insertAtCursor({ type: "operator", operator })}>
                {operator === "*" ? "×" : operator === "/" ? "÷" : operator}
              </Button>
            ))}
          </div>
        </div>
        <div>
          <div className="mb-2 text-caption font-bold text-secondary-foreground">括号与函数</div>
          <div className="flex gap-1">
            <Button size="sm" variant="outline" onClick={() => insertAtCursor({ type: "left_parenthesis" })}>(</Button>
            <Button size="sm" variant="outline" onClick={() => insertAtCursor({ type: "right_parenthesis" })}>)</Button>
            <Button size="sm" variant="outline" onClick={() => insertAtCursor({ type: "function", name: "SUM" })}>SUM(</Button>
            <Button size="sm" variant="outline" onClick={() => insertAtCursor({ type: "comma" })}>,</Button>
          </div>
        </div>
        <div className="min-w-[190px] flex-1">
          <div className="mb-2 text-caption font-bold text-secondary-foreground">固定值</div>
          <div className="flex gap-2">
            <Input className="numeric-value" type="number" value={constantValue} onChange={event => setConstantValue(event.target.value)} />
            <Button size="sm" variant="secondary" onClick={insertConstant}>插入</Button>
          </div>
        </div>
        <div className="ml-auto flex gap-1">
          <Button size="sm" variant="ghost" disabled={cursor === 0} onClick={backspace}>回退</Button>
          <Button size="sm" variant="ghost" disabled={formula.tokens.length === 0} className="text-destructive" onClick={() => setTokens([], 0)}>清空</Button>
        </div>
      </div>

    </div>
  );
}
