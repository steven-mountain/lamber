import { useState } from "react";
import AppIcon from "../../components/icons/AppIcon";
import {
  AI_COMPUTE_FUNDING_PLAN_YEARS,
  normalizeAiComputeFundingPlan,
  updateAiComputeFundingPlanMode,
  updateAiComputeFundingPlanYear,
  validateAiComputeFundingPlan,
} from "./fundingPlans";
import type {
  AiComputeLineItemFundingPlan,
  AiComputeLineItemFundingPlanMode,
  AiComputeQuoteLineItem,
} from "./types";

type Props = {
  item: AiComputeQuoteLineItem;
  projectCycleYears: number;
  onChange: (plan: AiComputeLineItemFundingPlan) => void;
};

const MODE_LABELS: Record<AiComputeLineItemFundingPlanMode, string> = {
  first_year: "第一年度一次性",
  even: "平均分年",
  manual: "自定义年度金额",
};

const formatMoney = (value: number) =>
  value.toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export default function AiComputeFundingPlanEditor({ item, projectCycleYears, onChange }: Props) {
  const [open, setOpen] = useState(false);
  const plan = normalizeAiComputeFundingPlan(item.fundingPlan, item.amountInclTax, projectCycleYears);
  const validation = validateAiComputeFundingPlan(plan, item.amountInclTax);
  const actionLabel = item.side === "revenue" ? "收款" : "付款";
  const statusLabel = validation.consistent ? "一致" : "不一致";
  const statusClass = validation.consistent
    ? "bg-success-soft text-success-foreground"
    : "bg-warning-soft text-warning-foreground";

  return (
    <div className="mt-3 rounded-lg bg-card/75 p-3">
      <div className="flex flex-wrap items-center gap-2">
        <button
          type="button"
          onClick={() => setOpen(value => !value)}
          className="inline-flex items-center gap-1 rounded-md bg-secondary px-2.5 py-1.5 text-[11px] font-bold text-secondary-foreground transition-colors hover:bg-primary-soft hover:text-primary"
        >
          {open ? "收起资金计划" : "展开资金计划"}
          <AppIcon name={open ? "chevronUp" : "chevronDown"} />
        </button>
        <span className={`rounded-full px-2 py-0.5 text-[10px] font-bold ${statusClass}`}>
          {statusLabel}
        </span>
        {!plan.enabled && (
          <span className="rounded-full bg-secondary px-2 py-0.5 text-[10px] font-bold text-secondary-foreground">
            已关闭
          </span>
        )}
        {!validation.consistent && (
          <span className="text-[10px] font-semibold text-warning-foreground">
            差额 {formatMoney(validation.difference)} 元
          </span>
        )}
      </div>

      {open && (
        <div className="mt-3 rounded-lg bg-muted/60 p-3">
          <div className="flex flex-col gap-2 sm:flex-row sm:items-start sm:justify-between">
            <div>
              <div className="text-xs font-extrabold text-foreground">
                {item.name} · {actionLabel}计划
              </div>
              <div className="mt-0.5 text-[10px] font-semibold text-secondary-foreground">
                绑定 ID: {item.id}
              </div>
            </div>
            <label className="inline-flex items-center gap-1.5 text-[11px] font-bold text-secondary-foreground">
              <input
                type="checkbox"
                checked={plan.enabled}
                onChange={event => onChange({ ...plan, enabled: event.target.checked })}
              />
              启用计划
            </label>
          </div>

          <div className="mt-3 grid grid-cols-1 gap-2 sm:grid-cols-3">
            {(Object.keys(MODE_LABELS) as AiComputeLineItemFundingPlanMode[]).map(mode => (
              <label
                key={mode}
                className={`flex cursor-pointer items-center gap-2 rounded-md px-2.5 py-2 text-xs font-semibold transition-colors ${
                  plan.mode === mode
                    ? "bg-primary-soft text-primary"
                    : "bg-card text-secondary-foreground hover:bg-secondary"
                } ${plan.enabled ? "" : "pointer-events-none opacity-55"}`}
              >
                <input
                  type="radio"
                  name={`ai-compute-funding-mode-${item.id}`}
                  checked={plan.mode === mode}
                  disabled={!plan.enabled}
                  onChange={() => onChange(updateAiComputeFundingPlanMode(
                    plan,
                    mode,
                    item.amountInclTax,
                    projectCycleYears,
                  ))}
                />
                {MODE_LABELS[mode]}{actionLabel}
              </label>
            ))}
          </div>

          <div className="mt-3 grid grid-cols-2 gap-2 sm:grid-cols-5">
            {Array.from({ length: AI_COMPUTE_FUNDING_PLAN_YEARS }, (_, index) => {
              const year = index + 1;
              const outsideCycle = year > projectCycleYears;
              return (
                <label
                  key={year}
                  className={`flex flex-col gap-1 text-[10px] font-bold ${
                    outsideCycle ? "text-secondary-foreground/65" : "text-secondary-foreground"
                  }`}
                >
                  第 {year} 年{outsideCycle ? "（周期外）" : ""}
                  <input
                    type="number"
                    min={0}
                    disabled={!plan.enabled}
                    value={plan.yearlyAmounts[String(year)] || ""}
                    placeholder="0"
                    onChange={event => onChange(updateAiComputeFundingPlanYear(
                      plan,
                      year,
                      Number(event.target.value),
                    ))}
                    className={`numeric-value rounded-md px-2.5 py-2 text-xs font-semibold outline-none ring-1 ring-ring/20 focus:ring-ring ${
                      plan.enabled
                        ? "bg-card text-foreground"
                        : "bg-secondary text-secondary-foreground"
                    }`}
                  />
                </label>
              );
            })}
          </div>

          <div className={`mt-3 grid grid-cols-1 gap-2 rounded-md p-3 text-xs font-semibold sm:grid-cols-4 ${
            plan.enabled && !validation.consistent ? "bg-warning-soft" : "bg-card"
          }`}>
            <div>
              <span className="text-secondary-foreground">计划合计：</span>
              <span className="numeric-value text-foreground">{formatMoney(validation.plannedAmount)} 元</span>
            </div>
            <div>
              <span className="text-secondary-foreground">科目金额：</span>
              <span className="numeric-value text-foreground">{formatMoney(validation.subjectAmount)} 元</span>
            </div>
            <div className={validation.consistent ? "text-success-foreground" : "text-warning-foreground"}>
              校验结果：{validation.consistent ? "一致" : "不一致"}
            </div>
            <div className={validation.consistent ? "text-secondary-foreground" : "text-warning-foreground"}>
              差额：<span className="numeric-value">{formatMoney(validation.difference)} 元</span>
            </div>
          </div>

          <div className="mt-2 text-[10px] leading-relaxed text-secondary-foreground">
            当前项目周期为 {projectCycleYears} 年。周期外年份保留输入位置，自动计划默认填 0；不一致仅提示，不阻止保存。
          </div>
        </div>
      )}
    </div>
  );
}
