import { useMemo, useState } from "react";
import type { IctSubjectDefinition, IctTaxItemLike } from "../lib/ictSubjectCatalog";
import {
  buildEqualAnnualInclValues,
  createDefaultSubjectFundingPlan,
  createSubjectFundingPlanId,
  setSubjectFundingPlanEnabled,
  updateSubjectFundingPlanAnnualValue,
  updateSubjectFundingPlanMode,
  validateSubjectFundingPlan,
  type CashflowCalculationSource,
  type SubjectFundingPlan,
  type SubjectFundingPlanMode,
  type SubjectFundingSubjectRef,
} from "../lib/ictSubjectFundingPlan";

type IctSubjectFundingPlanEditorProps = {
  subject: IctSubjectDefinition;
  item: IctTaxItemLike | null | undefined;
  plan?: SubjectFundingPlan;
  displayName: string;
  calculationSource?: CashflowCalculationSource;
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
  custom: "自定义年度金额",
};

export default function IctSubjectFundingPlanEditor({
  subject,
  item,
  plan,
  displayName,
  calculationSource = "legacy_model",
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
  const isDisabled = plan && !plan.enabled;

  const handleToggle = () => {
    if (!open && !plan) {
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
                checked={activePlan.enabled}
                onChange={event => onPlanChange(setSubjectFundingPlanEnabled(activePlan, event.target.checked))}
              />
              启用计划
            </label>
          </div>

          <div className="mt-3 grid grid-cols-1 gap-2 sm:grid-cols-3">
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
            {calculationSource === "subject_funding_plans"
              ? "当前已选择按科目收付款计划计算现金流。覆盖校验通过后，本计划会参与正式年度现金流、NPV和回收期测算。"
              : "当前沿用原资金模型计算现金流。该计划会保存并校验，但不会影响正式年度现金流、NPV或回收期。"}
          </div>
        </div>
      )}
    </div>
  );
}
