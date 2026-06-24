import { useEffect, useMemo, useState } from "react";
import type { IctSubjectDefinition, IctTaxItemLike } from "../lib/ictSubjectCatalog";
import {
  buildEqualAnnualInclValues,
  createDefaultSubjectFundingPlan,
  createSubjectFundingPlanId,
  normalizeAnnualPercentages,
  setSubjectFundingPlanEnabled,
  sumAnnualPercentages,
  updateSubjectFundingPlanAnnualValue,
  updateSubjectFundingPlanMode,
  updateSubjectFundingPlanPercentage,
  validateSubjectFundingPlan,
  type SubjectFundingPlan,
  type SubjectFundingPlanLastChangeReason,
  type SubjectFundingPlanMode,
  type SubjectFundingSubjectRef,
} from "../lib/ictSubjectFundingPlan";

type IctSubjectFundingPlanEditorProps = {
  subject: IctSubjectDefinition;
  item: IctTaxItemLike | null | undefined;
  plan?: SubjectFundingPlan;
  displayName: string;
  forceOpenToken?: number;
  onPlanChange: (plan: SubjectFundingPlan) => void;
};

const formatMoney = (value: number) =>
  new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(Number.isFinite(value) ? value : 0);

const modeLabels: Record<SubjectFundingPlanMode, string> = {
  upfront: "第一年一次性",
  equal: "平均分年",
  proportional: "按比例",
  custom: "自定义年度金额",
};

const formatPercent = (value: number) =>
  `${(Number.isFinite(value) ? Math.round(value * 10000) / 10000 : 0)
    .toLocaleString("zh-CN", { maximumFractionDigits: 4 })}%`;

const reasonTexts: Record<SubjectFundingPlanLastChangeReason, string> = {
  manual_plan_edit: "用户已手工维护计划",
  manual_amount_sync: "科目金额变更，计划已自动按原年度比例缩放调整",
  reverse_calculation_sync: "已随智能反算结果，按原年度比例自动调整",
  balance_allocation_sync: "已随差额承接结果，按原年度比例自动调整",
  ct_linkage_sync: "因 CT 业务金额联动而自动调整",
  auto_created_upfront: "系统自动创建默认计划（第一年一次性）",
  restored_after_zero: "金额从零恢复，已按最近一次有效年度结构自动恢复",
  legacy_migration: "旧项目已迁移为第一年一次性计划",
  ai_compute_quote_import: "已从智算报价测算写入金额和年度计划",
  intelligent_compute_import: "已从智算金额来源同步金额和年度计划",
};

export default function IctSubjectFundingPlanEditor({
  subject,
  item,
  plan,
  displayName,
  forceOpenToken = 0,
  onPlanChange,
}: IctSubjectFundingPlanEditorProps) {
  const [open, setOpen] = useState(false);
  const subjectAmountIncl = Number(item?.incl ?? 0);
  const actionLabel = subject.side === "revenue" ? "收款" : "付款";
  const subjectRef = useMemo<SubjectFundingSubjectRef>(() => ({
    side: subject.side,
    groupId: subject.groupId,
    key: subject.key,
  }), [subject.groupId, subject.key, subject.side]);
  const planId = createSubjectFundingPlanId(subjectRef);
  const activePlan = plan || createDefaultSubjectFundingPlan(subjectRef, subjectAmountIncl);
  const validation = validateSubjectFundingPlan(plan, subjectAmountIncl);
  const annualValues = activePlan.annualInclValues;
  const equalYears = activePlan.equalYears || 10;
  const annualPercentages = normalizeAnnualPercentages(activePlan.annualPercentages);
  const percentageSum = sumAnnualPercentages(activePlan.annualPercentages);
  const percentageComplete = Math.abs(percentageSum - 100) < 1e-4;
  const isDisabled = plan && !plan.enabled;

  useEffect(() => {
    if (forceOpenToken > 0) {
      setOpen(true);
    }
  }, [forceOpenToken]);

  const handleToggle = () => {
    if (!open && !plan && subjectAmountIncl > 0) {
      onPlanChange(createDefaultSubjectFundingPlan(subjectRef, subjectAmountIncl));
    }
    setOpen(value => !value);
  };

  const handleModeChange = (mode: SubjectFundingPlanMode) => {
    onPlanChange(updateSubjectFundingPlanMode(activePlan, subjectAmountIncl, mode, equalYears));
  };

  const handleEqualYearsChange = (years: number) => {
    onPlanChange({
      ...activePlan,
      mode: "equal",
      equalYears: years,
      annualInclValues: buildEqualAnnualInclValues(subjectAmountIncl, years),
      enabled: true,
      updatedAt: new Date().toISOString(),
    });
  };

  const handlePercentageChange = (yearIndex: number, percentage: number) => {
    onPlanChange(updateSubjectFundingPlanPercentage(activePlan, subjectAmountIncl, yearIndex, percentage));
  };

  const statusLabel = !plan
    ? "未维护"
    : isDisabled
      ? "已停用"
      : validation.valid
        ? "一致"
        : "待调整";
  const statusClass = !plan || isDisabled
    ? "bg-secondary text-secondary-foreground"
    : validation.valid
      ? "bg-success-soft text-success-foreground"
      : "bg-warning-soft text-warning-foreground";

  return (
    <div className="mt-1 flex flex-col gap-2">
      <div className="flex flex-wrap items-center gap-2">
        <button
          type="button"
          onClick={handleToggle}
          className="rounded-md bg-secondary px-2.5 py-1 text-[11px] font-bold text-secondary-foreground transition-colors hover:bg-primary-soft hover:text-primary"
        >
          {open ? `收起${actionLabel}计划` : `${actionLabel}计划`}
        </button>
        <span className={`rounded-full px-2 py-0.5 text-[10px] font-bold ${statusClass}`}>
          {statusLabel}
        </span>
        {plan && plan.enabled && !validation.valid && (
          <span className="text-[10px] font-semibold text-warning-foreground">
            {validation.difference > 0
              ? `尚有 ${formatMoney(validation.difference)} 元未安排${actionLabel}`
              : `${actionLabel}计划超出 ${formatMoney(Math.abs(validation.difference))} 元`}
          </span>
        )}
      </div>

      {open && (
        <div className="rounded-lg bg-muted/60 p-3">
          <div className="flex flex-col gap-1.5 sm:flex-row sm:items-start sm:justify-between">
            <div className="min-w-0">
              <div className="text-xs font-extrabold text-foreground">
                {displayName} · {actionLabel}计划
              </div>
              <div className="mt-0.5 text-[10px] font-semibold text-secondary-foreground">
                绑定 ID：{planId}
              </div>
            </div>
            <label className="inline-flex items-center gap-1.5 text-[11px] font-bold text-secondary-foreground">
              <input
                type="checkbox"
                checked={Boolean(plan?.enabled)}
                onChange={event => onPlanChange(setSubjectFundingPlanEnabled(activePlan, event.target.checked))}
              />
              启用计划
            </label>
          </div>

          <div className="mt-3 grid grid-cols-2 gap-2 sm:grid-cols-4">
            {(Object.keys(modeLabels) as SubjectFundingPlanMode[]).map(mode => (
              <label
                key={mode}
                className={`flex cursor-pointer items-center gap-2 rounded-md px-2.5 py-2 text-xs font-semibold transition-colors ${
                  activePlan.mode === mode ? "bg-primary-soft text-primary" : "bg-card text-secondary-foreground hover:bg-secondary"
                }`}
              >
                <input
                  type="radio"
                  name={`funding-mode-${planId}`}
                  checked={activePlan.mode === mode}
                  onChange={() => handleModeChange(mode)}
                />
                {modeLabels[mode]}{actionLabel}
              </label>
            ))}
          </div>

          {activePlan.mode === "equal" && (
            <div className="mt-3 flex flex-wrap items-center gap-2 text-xs font-semibold text-secondary-foreground">
              <span>分摊年限</span>
              <select
                value={equalYears}
                onChange={event => handleEqualYearsChange(Number(event.target.value))}
                className="rounded-md bg-card px-2.5 py-1.5 text-xs font-bold text-foreground outline-none ring-1 ring-ring/20 focus:ring-ring"
              >
                {Array.from({ length: 10 }, (_, index) => index + 1).map(year => (
                  <option key={year} value={year}>{year} 年</option>
                ))}
              </select>
              <span>从第 1 年开始平均分摊，尾差并入最后一年。</span>
            </div>
          )}

          {activePlan.mode === "proportional" && (
            <div className="mt-3 rounded-md bg-card/70 p-3">
              <div className="flex flex-wrap items-center justify-between gap-2 text-xs font-semibold text-secondary-foreground">
                <span>按比例分配（各年填写百分比，合计需为 100%，系统自动换算为年度金额）</span>
                <span className={percentageComplete ? "text-success-foreground" : "text-warning-foreground"}>
                  比例合计：{formatPercent(percentageSum)}
                  {percentageComplete ? "（已满 100%）" : `（${percentageSum > 100 ? "超出" : "尚差"} ${formatPercent(Math.abs(100 - percentageSum))}）`}
                </span>
              </div>
              <div className="mt-2 grid grid-cols-2 gap-2 sm:grid-cols-5">
                {Array.from({ length: 10 }, (_, index) => (
                  <label key={index} className="flex flex-col gap-1 text-[10px] font-bold text-secondary-foreground">
                    第 {index + 1} 年(%)
                    <input
                      type="number"
                      min={0}
                      max={100}
                      step="any"
                      value={Number(annualPercentages[index] || 0) === 0 ? "" : annualPercentages[index]}
                      placeholder="0"
                      onChange={event => handlePercentageChange(index, Number(event.target.value))}
                      className="rounded-md bg-card px-2.5 py-2 text-xs font-semibold text-foreground outline-none ring-1 ring-ring/20 focus:ring-ring"
                    />
                  </label>
                ))}
              </div>
              <div className="mt-2 text-[10px] leading-relaxed text-secondary-foreground">
                例如第 1 年填 95、第 6 年填 5，即第 1 年{actionLabel} 95%、第 6 年{actionLabel} 5%；下方为换算后的年度{actionLabel}金额（含税），尾差并入最后一个有比例的年度。
              </div>
            </div>
          )}

          <div className="mt-3 grid grid-cols-2 gap-2 sm:grid-cols-5">
            {Array.from({ length: 10 }, (_, index) => (
              <label key={index} className="flex flex-col gap-1 text-[10px] font-bold text-secondary-foreground">
                第 {index + 1} 年
                <input
                  type="number"
                  min={0}
                  readOnly={activePlan.mode !== "custom"}
                  value={Number(annualValues[index] || 0) === 0 ? "" : annualValues[index]}
                  placeholder="0"
                  onChange={event => {
                    const nextValue = Number(event.target.value);
                    onPlanChange(updateSubjectFundingPlanAnnualValue(activePlan, index, nextValue));
                  }}
                  className={`rounded-md px-2.5 py-2 text-xs font-semibold outline-none ring-1 ring-ring/20 focus:ring-ring ${
                    activePlan.mode === "custom" ? "bg-card text-foreground" : "bg-secondary text-secondary-foreground"
                  }`}
                />
              </label>
            ))}
          </div>

          <div className="mt-3 grid grid-cols-1 gap-2 rounded-md bg-card p-3 text-xs font-semibold sm:grid-cols-3">
            <div>
              <span className="text-secondary-foreground">计划合计：</span>
              <span className="numeric-value text-foreground">{formatMoney(validateSubjectFundingPlan(activePlan, subjectAmountIncl).plannedAmountIncl)} 元</span>
            </div>
            <div>
              <span className="text-secondary-foreground">科目金额：</span>
              <span className="numeric-value text-foreground">{formatMoney(subjectAmountIncl)} 元</span>
            </div>
            <div className={validateSubjectFundingPlan(activePlan, subjectAmountIncl).valid ? "text-success-foreground" : "text-warning-foreground"}>
              {validateSubjectFundingPlan(activePlan, subjectAmountIncl).valid
                ? "校验结果：一致"
                : validateSubjectFundingPlan(activePlan, subjectAmountIncl).difference > 0
                  ? `校验结果：尚有 ${formatMoney(validateSubjectFundingPlan(activePlan, subjectAmountIncl).difference)} 元未安排${actionLabel}`
                  : `校验结果：计划超出 ${formatMoney(Math.abs(validateSubjectFundingPlan(activePlan, subjectAmountIncl).difference))} 元`}
            </div>
          </div>

          <div className="mt-2 text-[10px] leading-relaxed text-secondary-foreground">
            覆盖校验通过后，本计划会参与正式年度现金流、NPV、IRR和回收期测算。
          </div>

          {activePlan.lastChangeReason && activePlan.lastChangeReason !== "manual_plan_edit" && (
            <div className="mt-2 rounded-md bg-primary-soft/50 px-2.5 py-1.5 text-[10px] font-semibold text-primary">
              <span className="mr-1">ℹ️</span>
              {reasonTexts[activePlan.lastChangeReason] || "计划已自动调整"}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
