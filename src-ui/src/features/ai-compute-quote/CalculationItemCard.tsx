import { useState, type ReactNode } from "react";
import AppIcon from "../../components/icons/AppIcon";
import { Button } from "../../components/ui/button";
import { ICT_SUBJECT_DEFINITIONS } from "../../lib/ictSubjectCatalog";
import {
  normalizeAiComputeFundingPlan,
  validateAiComputeFundingPlan,
} from "./fundingPlans";
import { normalizeQuoteFormula } from "./formulaEngine";
import CalculationItemDetailTabs from "./CalculationItemDetailTabs";
import { useAiComputeQuoteStore } from "./store";
import type {
  AiComputeQuoteFormulaToken,
  AiComputeQuoteLineItem,
} from "./types";

type SyncStatus = "idle" | "syncing" | "synced" | "error" | "conflict";

type Props = {
  item: AiComputeQuoteLineItem;
  totalAmount: number;
  projectCycleYears: number;
  syncStatus: SyncStatus;
};

function formatWan(value: number) {
  return (value / 10000).toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function formatNumber(value: number, digits = 2) {
  return value.toLocaleString("zh-CN", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

function formulaTokenLabel(
  token: AiComputeQuoteFormulaToken,
  parameters: { id: string; name: string }[],
  lineItems: { id: string; name: string }[],
) {
  if (token.type === "parameter") {
    return parameters.find(parameter => parameter.id === token.id)?.name || token.name || token.id;
  }
  if (token.type === "line_item") {
    return lineItems.find(lineItem => lineItem.id === token.id)?.name || token.name || token.id;
  }
  if (token.type === "constant") return String(token.value);
  if (token.type === "operator") return token.operator === "*" ? "×" : token.operator === "/" ? "÷" : token.operator;
  if (token.type === "left_parenthesis") return "(";
  if (token.type === "right_parenthesis") return ")";
  if (token.type === "function") return "SUM(";
  return "，";
}

function getFormulaSummary(
  tokens: AiComputeQuoteFormulaToken[],
  parameters: { id: string; name: string }[],
  lineItems: { id: string; name: string }[],
) {
  if (tokens.length === 0) return "公式不完整";
  return tokens.map(token => formulaTokenLabel(token, parameters, lineItems)).join(" ");
}

function getIctStatus(
  item: AiComputeQuoteLineItem,
  mappingName: string | undefined,
  syncStatus: SyncStatus,
) {
  if (item.outputEnabled === false) {
    return { label: "不参与同步输出", tone: "muted" as const };
  }
  if (!mappingName) {
    return { label: "未映射 ICT 科目", tone: "warning" as const };
  }
  if (syncStatus === "conflict") {
    return { label: `存在冲突 · ${mappingName}`, tone: "warning" as const };
  }
  if (syncStatus === "syncing") {
    return { label: `同步中 · ${mappingName}`, tone: "primary" as const };
  }
  if (syncStatus === "error") {
    return { label: `同步失败 · ${mappingName}`, tone: "warning" as const };
  }
  if (syncStatus === "synced") {
    return { label: `已同步 · ${mappingName}`, tone: "success" as const };
  }
  return { label: `待同步 · ${mappingName}`, tone: "primary" as const };
}

function StatusPill({
  children,
  tone,
}: {
  children: ReactNode;
  tone: "success" | "warning" | "primary" | "muted";
}) {
  const toneClass = tone === "success"
    ? "bg-success-soft text-success-foreground"
    : tone === "warning"
      ? "bg-warning-soft text-warning-foreground"
      : tone === "primary"
        ? "bg-primary-soft text-primary"
        : "bg-secondary text-secondary-foreground";

  return (
    <span className={`inline-flex min-w-0 items-center gap-1.5 rounded-full px-2.5 py-1 text-[11px] font-bold ${toneClass}`}>
      <span className="h-1.5 w-1.5 shrink-0 rounded-full bg-current opacity-70" />
      <span className="truncate">{children}</span>
    </span>
  );
}

export default function CalculationItemCard({
  item,
  totalAmount,
  projectCycleYears,
  syncStatus,
}: Props) {
  const store = useAiComputeQuoteStore();
  const [expanded, setExpanded] = useState(false);
  const mapping = store.blueprint.mappings.find(candidate => candidate.lineItemId === item.id);
  const mappingName = mapping?.ictSubjectName
    || ICT_SUBJECT_DEFINITIONS.find(subject => subject.subjectCode === mapping?.ictSubjectCode)?.standardSubjectName;
  const normalizedFormula = normalizeQuoteFormula(item.formula, store.blueprint.parameters);
  const amountShare = totalAmount > 0
    ? Math.max(0, Math.min(100, item.amountInclTax / totalAmount * 100))
    : 0;
  const fundingPlan = normalizeAiComputeFundingPlan(item.fundingPlan, item.amountInclTax, projectCycleYears);
  const fundingValidation = validateAiComputeFundingPlan(fundingPlan, item.amountInclTax);
  const ictStatus = getIctStatus(item, mappingName, syncStatus);
  const formulaSummary = getFormulaSummary(
    normalizedFormula.tokens,
    store.blueprint.parameters,
    [...store.blueprint.revenueItems, ...store.blueprint.costItems],
  );
  const accentClass = item.side === "revenue"
    ? "bg-success-soft/55 text-success-foreground"
    : "bg-warning-soft/65 text-warning-foreground";
  const progressClass = item.side === "revenue" ? "bg-success/70" : "bg-destructive/65";

  const selectMapping = (subjectCode: string) => {
    const subject = ICT_SUBJECT_DEFINITIONS.find(candidate =>
      candidate.side === item.side && candidate.subjectCode === subjectCode
    );
    if (!subject) {
      store.updateMapping(item.id, null);
      return;
    }
    store.updateMapping(item.id, {
      id: mapping?.id || `mapping-${item.id}`,
      lineItemId: item.id,
      side: item.side,
      ictSubjectCode: subject.subjectCode,
      ictSubjectName: subject.standardSubjectName,
      enabled: true,
    });
  };

  return (
    <article className={`overflow-hidden rounded-xl bg-card shadow-sm ${item.enabled ? "" : "opacity-70"}`}>
      <div className="p-4">
        <div className="flex flex-col gap-3 xl:flex-row xl:items-start">
          <div className="flex min-w-0 flex-1 items-start gap-3">
            <label className={`mt-0.5 flex h-8 w-8 shrink-0 cursor-pointer items-center justify-center rounded-lg ${accentClass}`} title="启用 / 停用">
              <input
                className="h-4 w-4"
                type="checkbox"
                checked={item.enabled}
                onChange={event => store.updateLineItem(item.side, item.id, { enabled: event.target.checked })}
              />
            </label>
            <div className="min-w-0">
              <div className="flex flex-wrap items-center gap-2">
                <h3 className="truncate text-base font-extrabold text-foreground">{item.name}</h3>
                {!item.enabled && (
                  <span className="rounded-full bg-secondary px-2 py-0.5 text-[10px] font-bold text-secondary-foreground">
                    已停用
                  </span>
                )}
                {item.calculationStatus !== "valid" && (
                  <span className="rounded-full bg-destructive-soft px-2 py-0.5 text-[10px] font-bold text-destructive">
                    公式异常
                  </span>
                )}
              </div>
              <div className="mt-1 flex min-w-0 items-center gap-1.5 text-caption text-secondary-foreground">
                <span className="shrink-0 font-semibold">公式</span>
                <span className="truncate font-medium text-foreground" title={formulaSummary}>{formulaSummary}</span>
              </div>
            </div>
          </div>

          <div className="grid shrink-0 grid-cols-1 gap-2 sm:min-w-[390px] sm:grid-cols-3">
            <div className="rounded-lg bg-muted/55 px-3 py-2">
              <div className="text-[10px] font-semibold text-secondary-foreground">税率</div>
              <div className="numeric-value mt-0.5 text-sm font-extrabold">{formatNumber(item.taxRate || 0)}%</div>
            </div>
            <div className={`rounded-lg px-3 py-2 ${item.side === "revenue" ? "bg-success-soft/55" : "bg-warning-soft/55"}`}>
              <div className="text-[10px] font-semibold text-secondary-foreground">含税金额（万元）</div>
              <div className="numeric-value mt-0.5 text-sm font-extrabold">{formatWan(item.amountInclTax)}</div>
            </div>
            <div className="rounded-lg bg-muted/55 px-3 py-2">
              <div className="text-[10px] font-semibold text-secondary-foreground">不含税（万元）</div>
              <div className="numeric-value mt-0.5 text-sm font-extrabold">{formatWan(item.amountExclTax)}</div>
            </div>
          </div>
        </div>

        <div className="mt-3 grid gap-3 lg:grid-cols-[minmax(180px,0.8fr)_minmax(0,1fr)_auto] lg:items-center">
          <div className="min-w-0">
            <div className="mb-1.5 flex items-center justify-between gap-2 text-[11px] font-semibold text-secondary-foreground">
              <span>占总{item.side === "revenue" ? "收入" : "成本"}</span>
              <span className="numeric-value text-foreground">{formatNumber(amountShare, 1)}%</span>
            </div>
            <div className="h-1.5 overflow-hidden rounded-full bg-muted">
              <div className={`h-full rounded-full ${progressClass}`} style={{ width: `${amountShare}%` }} />
            </div>
          </div>

          <div className="flex min-w-0 flex-wrap items-center gap-2">
            <StatusPill tone={ictStatus.tone}>{ictStatus.label}</StatusPill>
            <StatusPill tone={!fundingPlan.enabled ? "muted" : fundingValidation.consistent ? "success" : "warning"}>
              资金计划 · {!fundingPlan.enabled ? "已关闭" : fundingValidation.consistent ? "一致" : "有差异"}
            </StatusPill>
          </div>

          <Button
            variant={expanded ? "secondary" : "ghost"}
            size="sm"
            className="justify-self-start lg:justify-self-end"
            aria-expanded={expanded}
            onClick={() => setExpanded(value => !value)}
          >
            {expanded ? "收起详情" : "展开详情"}
            <AppIcon name={expanded ? "chevronUp" : "chevronDown"} />
          </Button>
        </div>

        {item.calculationStatus !== "valid" && !expanded && (
          <div className={`mt-3 rounded-lg px-3 py-2 text-caption font-semibold ${
            item.calculationStatus === "error"
              ? "bg-destructive-soft text-destructive"
              : "bg-warning-soft text-warning-foreground"
          }`}>
            {item.calculationError || "公式不完整"}
          </div>
        )}
      </div>

      {expanded && (
        <div className="bg-muted/45 p-3 sm:p-4">
          <CalculationItemDetailTabs
            blueprint={store.blueprint}
            item={item}
            mapping={mapping}
            normalizedFormula={normalizedFormula}
            projectCycleYears={projectCycleYears}
            onUpdateItem={patch => store.updateLineItem(item.side, item.id, patch)}
            onUpdateFormula={formula => store.updateFormula(item.side, item.id, formula)}
            onUpdateMapping={selectMapping}
            onRemove={() => store.removeLineItem(item.side, item.id)}
          />
        </div>
      )}
    </article>
  );
}
