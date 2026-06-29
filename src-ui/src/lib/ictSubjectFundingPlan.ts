import type { IctSubjectGroupId, IctSubjectSide } from "./ictSubjectCatalog";

export type SubjectFundingPlanMode = "upfront" | "equal" | "proportional" | "custom";
export type SubjectFundingPlanSource = "manual" | "template" | "migration" | "ai_compute_quote" | "intelligent_compute";
export type CashflowCalculationSource = "legacy_model" | "subject_funding_plans";

export const SUBJECT_FUNDING_PLAN_MIGRATION_VERSION = 1;

export type SubjectFundingSubjectRef = {
  side: IctSubjectSide;
  groupId: IctSubjectGroupId;
  key: string;
};

export type SubjectFundingPlanLastChangeReason =
  | "manual_plan_edit"
  | "manual_amount_sync"
  | "reverse_calculation_sync"
  | "balance_allocation_sync"
  | "ct_linkage_sync"
  | "auto_created_upfront"
  | "restored_after_zero"
  | "legacy_migration"
  | "ai_compute_quote_import"
  | "intelligent_compute_import";

export type SubjectFundingPlanImportTrace = {
  source: "ai_compute_quote" | "intelligent_compute";
  sourceLabel: string;
  projectId: string;
  scenarioId: string;
  blueprintId: string;
  sourceLineItemIds: string[];
  amountSourceIds?: string[];
  sourceLineItems?: Array<{ amountSourceId: string; lineItemId: string }>;
  importedAt: string;
};

export interface SubjectFundingPlan {
  id: string;
  subjectRef: SubjectFundingSubjectRef;
  mode: SubjectFundingPlanMode;
  annualInclValues: number[];
  enabled: boolean;
  source: SubjectFundingPlanSource;
  equalYears?: number;
  annualPercentages?: number[];
  lastValidAnnualInclValues?: number[];
  lastChangeReason?: SubjectFundingPlanLastChangeReason;
  lastChangedAt?: string;
  updatedAt?: string;
  importTrace?: SubjectFundingPlanImportTrace;
}

export type SubjectFundingPlans = Record<string, SubjectFundingPlan>;

export type FundingPlanValidationResult = {
  valid: boolean;
  subjectAmountIncl: number;
  plannedAmountIncl: number;
  difference: number;
};

export type FundingPlanCoverageIssueType =
  | "missing_plan"
  | "disabled_plan"
  | "invalid_length"
  | "negative_annual_value"
  | "amount_mismatch"
  | "zero_subject_with_nonzero_plan";

export type SubjectFundingPlanCoverageSubject = {
  subjectRef: SubjectFundingSubjectRef;
  displayName: string;
  subjectAmountIncl: number;
  taxRate: number;
  isItScope?: boolean;
};

export type FundingPlanCoverageIssue = {
  type: FundingPlanCoverageIssueType;
  subjectRef: SubjectFundingSubjectRef;
  planId: string;
  displayName: string;
  side: IctSubjectSide;
  message: string;
  subjectAmountIncl: number;
  plannedAmountIncl?: number;
  difference?: number;
};

export type FundingPlanCoverageCounts = {
  revenueSubjectCount: number;
  costSubjectCount: number;
  revenuePlannedCount: number;
  costPlannedCount: number;
  zeroSubjectWithNonzeroPlanCount: number;
  issueCount: number;
};

export type FundingPlanCoverageResult = {
  valid: boolean;
  issues: FundingPlanCoverageIssue[];
  counts: FundingPlanCoverageCounts;
  revenueSubjects: SubjectFundingPlanCoverageSubject[];
  costSubjects: SubjectFundingPlanCoverageSubject[];
};

export type SubjectFundingPlanMigrationResult = {
  plans: SubjectFundingPlans;
  changed: boolean;
  completed: boolean;
  migrationVersion?: number;
  coverage: FundingPlanCoverageResult;
};

export type SubjectFundingAnnualCashflow = {
  source: "subject_funding_plans";
  annualRevenueIncl: number[];
  annualCostIncl: number[];
  annualRevenueExcl: number[];
  annualCostExcl: number[];
  annualItRevenueExcl: number[];
  annualItCostExcl: number[];
  annualNetIncl: number[];
  annualNetExcl: number[];
};

const PLAN_YEARS = 10;

const toMoneyCents = (value: number) => {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric <= 0) return 0;
  return Math.round(numeric * 100);
};

const toSignedMoneyCents = (value: unknown) => {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 0;
  return Math.round(numeric * 100);
};

const roundSignedFundingMoney = (value: number) => toSignedMoneyCents(value) / 100;

const formatFundingAmount = (value: number) =>
  roundSignedFundingMoney(value).toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

const sideLabel = (side: IctSubjectSide) => side === "revenue" ? "收入" : "投入";
const actionLabel = (side: IctSubjectSide) => side === "revenue" ? "收款" : "付款";

export const roundFundingMoney = (value: number) => toMoneyCents(value) / 100;

export const clampFundingPlanYears = (value: number) => {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return PLAN_YEARS;
  return Math.max(1, Math.min(Math.trunc(numeric), PLAN_YEARS));
};

export const createSubjectFundingPlanId = (ref: SubjectFundingSubjectRef) =>
  `${ref.side}:${ref.groupId}:${ref.key}`;

export const normalizeCashflowCalculationSource = (_value: unknown): CashflowCalculationSource =>
  "subject_funding_plans";

export const normalizeAnnualInclValues = (values: unknown): number[] => {
  const source = Array.isArray(values) ? values : [];
  return Array.from({ length: PLAN_YEARS }, (_, index) => {
    const numeric = Number(source[index] ?? 0);
    return roundFundingMoney(numeric);
  });
};

export const buildUpfrontAnnualInclValues = (subjectAmountIncl: number): number[] => {
  const values = Array(PLAN_YEARS).fill(0);
  values[0] = roundFundingMoney(subjectAmountIncl);
  return values;
};

export const buildEqualAnnualInclValues = (subjectAmountIncl: number, equalYears = PLAN_YEARS): number[] => {
  const values = Array(PLAN_YEARS).fill(0);
  const years = clampFundingPlanYears(equalYears);
  const totalCents = toMoneyCents(subjectAmountIncl);
  if (totalCents <= 0) return values;

  const baseCents = Math.floor(totalCents / years);
  const tailCents = totalCents - baseCents * years;
  for (let index = 0; index < years; index += 1) {
    values[index] = baseCents / 100;
  }
  values[years - 1] = (baseCents + tailCents) / 100;
  return values;
};

const PERCENT_EPSILON = 1e-6;

export const normalizeAnnualPercentages = (values: unknown): number[] => {
  const source = Array.isArray(values) ? values : [];
  return Array.from({ length: PLAN_YEARS }, (_, index) => {
    const numeric = Number(source[index] ?? 0);
    if (!Number.isFinite(numeric) || numeric <= 0) return 0;
    // store up to 4 decimal places of a percentage (e.g. 95, 5, 33.3333)
    return Math.min(100, Math.round(numeric * 10000) / 10000);
  });
};

export const sumAnnualPercentages = (values: unknown): number =>
  normalizeAnnualPercentages(values).reduce((sum, value) => sum + value, 0);

export const isAnnualPercentagesComplete = (values: unknown): boolean =>
  Math.abs(sumAnnualPercentages(values) - 100) < 1e-4;

export const buildDefaultAnnualPercentages = (): number[] => {
  const values = Array(PLAN_YEARS).fill(0);
  values[0] = 100;
  return values;
};

/**
 * Distribute a subject's inclusive amount across years by percentage.
 * Percentages are expected to sum to 100; any rounding tail is folded into
 * the last year that carries a non-zero percentage so the yearly total always
 * equals the subject amount (in integer cents).
 */
export const buildProportionalAnnualInclValues = (
  subjectAmountIncl: number,
  percentages: unknown,
): number[] => {
  const values = Array(PLAN_YEARS).fill(0);
  const pcts = normalizeAnnualPercentages(percentages);
  const totalCents = toMoneyCents(subjectAmountIncl);
  if (totalCents <= 0) return values;

  const pctSum = pcts.reduce((sum, value) => sum + value, 0);
  if (pctSum <= PERCENT_EPSILON) {
    values[0] = totalCents / 100;
    return values;
  }

  let allocatedCents = 0;
  let lastNonZeroIndex = -1;
  for (let index = 0; index < PLAN_YEARS; index += 1) {
    if (pcts[index] <= 0) continue;
    const cents = Math.round((totalCents * pcts[index]) / pctSum);
    values[index] = cents / 100;
    allocatedCents += cents;
    lastNonZeroIndex = index;
  }

  // Fold the rounding remainder into the last non-zero year.
  const diffCents = totalCents - allocatedCents;
  if (diffCents !== 0 && lastNonZeroIndex >= 0) {
    values[lastNonZeroIndex] = roundFundingMoney(
      values[lastNonZeroIndex] + diffCents / 100,
    );
  }
  return values;
};

export const createDefaultSubjectFundingPlan = (
  subjectRef: SubjectFundingSubjectRef,
  subjectAmountIncl: number,
): SubjectFundingPlan => ({
  id: createSubjectFundingPlanId(subjectRef),
  subjectRef,
  mode: "upfront",
  annualInclValues: buildUpfrontAnnualInclValues(subjectAmountIncl),
  enabled: true,
  source: "manual",
  updatedAt: new Date().toISOString(),
});

export const createLegacyMigrationSubjectFundingPlan = (
  subjectRef: SubjectFundingSubjectRef,
  subjectAmountIncl: number,
): SubjectFundingPlan => ({
  id: createSubjectFundingPlanId(subjectRef),
  subjectRef,
  mode: "upfront",
  annualInclValues: buildUpfrontAnnualInclValues(subjectAmountIncl),
  enabled: true,
  source: "migration",
  lastChangeReason: "legacy_migration",
  lastChangedAt: new Date().toISOString(),
  updatedAt: new Date().toISOString(),
});

export const normalizeSubjectFundingPlan = (value: unknown): SubjectFundingPlan | null => {
  if (!value || typeof value !== "object") return null;
  const raw = value as Partial<SubjectFundingPlan> & {
    subject_ref?: Partial<SubjectFundingSubjectRef>;
    annual_incl_values?: unknown;
    equal_years?: number;
    annual_percentages?: unknown;
  };
  const rawRef = raw.subjectRef || raw.subject_ref;
  if (!rawRef || rawRef.side !== "revenue" && rawRef.side !== "cost" || !rawRef.groupId || !rawRef.key) {
    return null;
  }

  const subjectRef: SubjectFundingSubjectRef = {
    side: rawRef.side,
    groupId: rawRef.groupId as IctSubjectGroupId,
    key: String(rawRef.key),
  };
  const mode: SubjectFundingPlanMode =
    raw.mode === "equal" || raw.mode === "custom" || raw.mode === "proportional" || raw.mode === "upfront"
      ? raw.mode
      : "upfront";
  const source: SubjectFundingPlanSource =
    raw.source === "template"
    || raw.source === "migration"
    || raw.source === "manual"
    || raw.source === "ai_compute_quote"
    || raw.source === "intelligent_compute"
      ? raw.source
      : "manual";

  return {
    id: createSubjectFundingPlanId(subjectRef),
    subjectRef,
    mode,
    annualInclValues: normalizeAnnualInclValues(raw.annualInclValues ?? raw.annual_incl_values),
    lastValidAnnualInclValues: normalizeAnnualInclValues(raw.lastValidAnnualInclValues),
    enabled: raw.enabled !== false,
    source,
    equalYears: clampFundingPlanYears(Number(raw.equalYears ?? raw.equal_years ?? PLAN_YEARS)),
    annualPercentages: normalizeAnnualPercentages(raw.annualPercentages ?? raw.annual_percentages),
    lastChangeReason: raw.lastChangeReason,
    lastChangedAt: typeof raw.lastChangedAt === "string" ? raw.lastChangedAt : undefined,
    updatedAt: typeof raw.updatedAt === "string" ? raw.updatedAt : undefined,
    importTrace: raw.importTrace,
  };
};

export const normalizeSubjectFundingPlans = (value: unknown): SubjectFundingPlans => {
  if (!value || typeof value !== "object") return {};
  return Object.values(value as Record<string, unknown>).reduce<SubjectFundingPlans>((plans, rawPlan) => {
    const plan = normalizeSubjectFundingPlan(rawPlan);
    if (!plan) return plans;
    plans[plan.id] = plan;
    return plans;
  }, {});
};

export function updateSubjectFundingPlanMode(
  plan: SubjectFundingPlan,
  subjectAmountIncl: number,
  mode: SubjectFundingPlanMode,
  equalYears: number = 10,
  reason: SubjectFundingPlanLastChangeReason = "manual_plan_edit"
): SubjectFundingPlan {
  // When switching to proportional, reuse stored percentages if present,
  // otherwise default to "100% in year 1" so the amount is fully scheduled.
  const annualPercentages = mode === "proportional"
    ? (sumAnnualPercentages(plan.annualPercentages) > 0
        ? normalizeAnnualPercentages(plan.annualPercentages)
        : buildDefaultAnnualPercentages())
    : plan.annualPercentages;

  const newValues = mode === "upfront"
    ? buildUpfrontAnnualInclValues(subjectAmountIncl)
    : mode === "equal"
      ? buildEqualAnnualInclValues(subjectAmountIncl, equalYears)
      : mode === "proportional"
        ? buildProportionalAnnualInclValues(subjectAmountIncl, annualPercentages)
        : normalizeAnnualInclValues(plan.annualInclValues);
  return {
    ...plan,
    mode,
    equalYears: mode === "equal" ? clampFundingPlanYears(equalYears) : plan.equalYears,
    annualPercentages,
    annualInclValues: newValues,
    enabled: true,
    lastValidAnnualInclValues: plan.annualInclValues,
    lastChangeReason: reason,
    lastChangedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}

export function updateSubjectFundingPlanPercentage(
  plan: SubjectFundingPlan,
  subjectAmountIncl: number,
  yearIndex: number,
  percentage: number,
  reason: SubjectFundingPlanLastChangeReason = "manual_plan_edit"
): SubjectFundingPlan {
  const annualPercentages = normalizeAnnualPercentages(plan.annualPercentages);
  if (yearIndex >= 0 && yearIndex < PLAN_YEARS) {
    const numeric = Number(percentage);
    annualPercentages[yearIndex] =
      Number.isFinite(numeric) && numeric > 0
        ? Math.min(100, Math.round(numeric * 10000) / 10000)
        : 0;
  }
  return {
    ...plan,
    mode: "proportional",
    annualPercentages,
    annualInclValues: buildProportionalAnnualInclValues(subjectAmountIncl, annualPercentages),
    enabled: true,
    lastValidAnnualInclValues: plan.annualInclValues,
    lastChangeReason: reason,
    lastChangedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}

export function updateSubjectFundingPlanAnnualValue(
  plan: SubjectFundingPlan,
  yearIndex: number,
  value: number,
  reason: SubjectFundingPlanLastChangeReason = "manual_plan_edit"
): SubjectFundingPlan {
  const annualInclValues = normalizeAnnualInclValues(plan.annualInclValues);
  if (yearIndex >= 0 && yearIndex < PLAN_YEARS) {
    annualInclValues[yearIndex] = roundFundingMoney(value);
  }
  return {
    ...plan,
    mode: "custom",
    annualInclValues,
    enabled: true,
    lastValidAnnualInclValues: plan.annualInclValues,
    lastChangeReason: reason,
    lastChangedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}

export const setSubjectFundingPlanEnabled = (
  plan: SubjectFundingPlan,
  enabled: boolean,
): SubjectFundingPlan => ({
  ...plan,
  enabled,
  updatedAt: new Date().toISOString(),
});

export const sumSubjectFundingPlanAnnualIncl = (plan: SubjectFundingPlan | null | undefined) =>
  normalizeAnnualInclValues(plan?.annualInclValues).reduce((sum, value) => sum + value, 0);

export const validateSubjectFundingPlan = (
  plan: SubjectFundingPlan | null | undefined,
  subjectAmountIncl: number,
): FundingPlanValidationResult => {
  const subjectCents = toMoneyCents(subjectAmountIncl);
  const plannedCents = toMoneyCents(sumSubjectFundingPlanAnnualIncl(plan));
  const differenceCents = subjectCents - plannedCents;
  return {
    valid: differenceCents === 0,
    subjectAmountIncl: subjectCents / 100,
    plannedAmountIncl: plannedCents / 100,
    difference: differenceCents / 100,
  };
};

const getRawAnnualValues = (plan: SubjectFundingPlan | null | undefined) =>
  Array.isArray(plan?.annualInclValues) ? plan.annualInclValues : [];

const sumRawAnnualCents = (plan: SubjectFundingPlan | null | undefined) =>
  getRawAnnualValues(plan).reduce((sum, value) => sum + toSignedMoneyCents(value), 0);

const hasRawNonzeroAnnualValue = (plan: SubjectFundingPlan | null | undefined) =>
  getRawAnnualValues(plan).some(value => Math.abs(toSignedMoneyCents(value)) > 0);

const pushCoverageIssue = (
  issues: FundingPlanCoverageIssue[],
  subject: SubjectFundingPlanCoverageSubject,
  type: FundingPlanCoverageIssueType,
  message: string,
  plannedAmountIncl?: number,
  difference?: number,
) => {
  issues.push({
    type,
    subjectRef: subject.subjectRef,
    planId: createSubjectFundingPlanId(subject.subjectRef),
    displayName: subject.displayName,
    side: subject.subjectRef.side,
    message,
    subjectAmountIncl: roundFundingMoney(subject.subjectAmountIncl),
    plannedAmountIncl,
    difference,
  });
};

export const validateSubjectFundingPlanCoverage = (
  subjects: SubjectFundingPlanCoverageSubject[],
  plans: SubjectFundingPlans,
): FundingPlanCoverageResult => {
  const issues: FundingPlanCoverageIssue[] = [];
  const counts: FundingPlanCoverageCounts = {
    revenueSubjectCount: 0,
    costSubjectCount: 0,
    revenuePlannedCount: 0,
    costPlannedCount: 0,
    zeroSubjectWithNonzeroPlanCount: 0,
    issueCount: 0,
  };

  subjects.forEach(subject => {
    const planId = createSubjectFundingPlanId(subject.subjectRef);
    const plan = plans[planId];
    const subjectCents = toMoneyCents(subject.subjectAmountIncl);
    const subjectAmount = subjectCents / 100;
    const label = sideLabel(subject.subjectRef.side);
    const action = actionLabel(subject.subjectRef.side);

    if (subjectCents <= 0) {
      if (plan && hasRawNonzeroAnnualValue(plan)) {
        counts.zeroSubjectWithNonzeroPlanCount += 1;
        pushCoverageIssue(
          issues,
          subject,
          "zero_subject_with_nonzero_plan",
          `${label}科目“${subject.displayName}”当前金额为 0，但${action}计划仍存在非零年度金额。`,
          sumRawAnnualCents(plan) / 100,
          -sumRawAnnualCents(plan) / 100,
        );
      }
      return;
    }

    if (subject.subjectRef.side === "revenue") counts.revenueSubjectCount += 1;
    else counts.costSubjectCount += 1;

    if (!plan) {
      pushCoverageIssue(
        issues,
        subject,
        "missing_plan",
        `${label}科目“${subject.displayName}”含税金额 ${formatFundingAmount(subjectAmount)} 元，尚未维护${action}计划。`,
      );
      return;
    }

    if (!plan.enabled) {
      pushCoverageIssue(
        issues,
        subject,
        "disabled_plan",
        `${label}科目“${subject.displayName}”已维护${action}计划，但当前计划已停用。`,
      );
      return;
    }

    const annualValues = getRawAnnualValues(plan);
    const rowIssueStart = issues.length;

    if (annualValues.length !== PLAN_YEARS) {
      pushCoverageIssue(
        issues,
        subject,
        "invalid_length",
        `${label}科目“${subject.displayName}”的${action}计划不是完整 10 年年度金额。`,
      );
    }

    if (annualValues.some(value => toSignedMoneyCents(value) < 0)) {
      pushCoverageIssue(
        issues,
        subject,
        "negative_annual_value",
        `${label}科目“${subject.displayName}”的${action}计划存在负数年度金额。`,
      );
    }

    const plannedCents = sumRawAnnualCents(plan);
    if (plannedCents !== subjectCents) {
      const plannedAmount = plannedCents / 100;
      const difference = (subjectCents - plannedCents) / 100;
      pushCoverageIssue(
        issues,
        subject,
        "amount_mismatch",
        `${label}科目“${subject.displayName}”的${action}计划合计 ${formatFundingAmount(plannedAmount)} 元，与科目含税金额 ${formatFundingAmount(subjectAmount)} 元不一致，差额 ${formatFundingAmount(difference)} 元。`,
        plannedAmount,
        difference,
      );
    }

    if (issues.length === rowIssueStart) {
      if (subject.subjectRef.side === "revenue") counts.revenuePlannedCount += 1;
      else counts.costPlannedCount += 1;
    }
  });

  counts.issueCount = issues.length;
  return {
    valid: issues.length === 0,
    issues,
    counts,
    revenueSubjects: subjects.filter(s => s.subjectRef.side === "revenue"),
    costSubjects: subjects.filter(s => s.subjectRef.side === "cost"),
  };
};

const isSubjectItScope = (subject: SubjectFundingPlanCoverageSubject) => {
  if (subject.isItScope !== undefined) return subject.isItScope;
  return subject.subjectRef.groupId === "revIt" || subject.subjectRef.groupId === "costIt";
};

const addAnnualValue = (target: number[], index: number, value: number) => {
  target[index] = roundSignedFundingMoney(target[index] + value);
};

/**
 * Restore a subject from tax-inclusive to tax-exclusive using the
 * "先还原后分摊" convention: de-tax the subject's *whole* inclusive total once
 * (rounded to cents), then split that exclusive total across years in
 * proportion to each year's inclusive cents. Any rounding tail is folded into
 * the last funded year so the yearly figures always sum exactly to the
 * subject's rounded tax-exclusive total — matching the 立项材料 subject table
 * and avoiding the 1-cent drift caused by de-taxing each year independently.
 *
 * @param inclCentsByYear per-year inclusive amounts, in integer cents
 * @param divisor 1 + taxRate (e.g. 1.13 for a 13% line)
 * @returns per-year exclusive amounts, in integer cents
 */
const distributeExclCentsByInclWeights = (
  inclCentsByYear: number[],
  divisor: number,
): number[] => {
  const result = Array(PLAN_YEARS).fill(0);
  const totalInclCents = inclCentsByYear.reduce((sum, cents) => sum + cents, 0);
  if (totalInclCents <= 0 || !(divisor > 0)) return result;

  const totalExclCents = Math.round(totalInclCents / divisor);
  let allocatedCents = 0;
  let lastFundedIndex = -1;
  for (let index = 0; index < PLAN_YEARS; index += 1) {
    const inclCents = inclCentsByYear[index];
    if (inclCents <= 0) continue;
    const cents = Math.round((totalExclCents * inclCents) / totalInclCents);
    result[index] = cents;
    allocatedCents += cents;
    lastFundedIndex = index;
  }

  const diffCents = totalExclCents - allocatedCents;
  if (diffCents !== 0 && lastFundedIndex >= 0) {
    result[lastFundedIndex] += diffCents;
  }
  return result;
};

export type BuildAnnualCashflowOptions = {
  /**
   * When true, subjects that carry an amount but have no enabled funding plan
   * (or whose plan is disabled) are still included in the cashflow as a
   * first-year ("upfront") payment, instead of being silently dropped. This
   * keeps the totals (and the IT breakdown) intact while still honoring any
   * subjects that DO have a maintained multi-year / proportional plan.
   */
  fallbackUnmaintainedToUpfront?: boolean;
};

export const buildAnnualCashflowFromSubjectFundingPlans = (
  subjects: SubjectFundingPlanCoverageSubject[],
  plans: SubjectFundingPlans,
  options: BuildAnnualCashflowOptions = {},
): SubjectFundingAnnualCashflow => {
  const annualRevenueIncl = Array(PLAN_YEARS).fill(0);
  const annualCostIncl = Array(PLAN_YEARS).fill(0);
  const annualRevenueExcl = Array(PLAN_YEARS).fill(0);
  const annualCostExcl = Array(PLAN_YEARS).fill(0);
  const annualItRevenueExcl = Array(PLAN_YEARS).fill(0);
  const annualItCostExcl = Array(PLAN_YEARS).fill(0);

  subjects.forEach(subject => {
    const subjectCents = toMoneyCents(subject.subjectAmountIncl);
    if (subjectCents <= 0) return;

    const plan = plans[createSubjectFundingPlanId(subject.subjectRef)];
    const taxRate = Number(subject.taxRate);
    const divisor = 1 + (Number.isFinite(taxRate) ? taxRate : 0) / 100;

    let annualValues: number[];
    if (plan?.enabled) {
      annualValues = normalizeAnnualInclValues(plan.annualInclValues);
    } else if (options.fallbackUnmaintainedToUpfront) {
      // No maintained plan → schedule the full amount in the first year.
      annualValues = buildUpfrontAnnualInclValues(subjectCents / 100);
    } else {
      return;
    }

    // 先把整笔含税额还原成不含税（取整一次），再按各年含税占比分摊，
    // 与立项材料科目表口径一致，避免逐年还原产生的尾差。
    const inclCentsByYear = annualValues.map(annualIncl => toMoneyCents(annualIncl));
    const exclCentsByYear = distributeExclCentsByInclWeights(inclCentsByYear, divisor);

    inclCentsByYear.forEach((inclCents, index) => {
      const inclValue = inclCents / 100;
      const exclValue = exclCentsByYear[index] / 100;
      if (subject.subjectRef.side === "revenue") {
        addAnnualValue(annualRevenueIncl, index, inclValue);
        addAnnualValue(annualRevenueExcl, index, exclValue);
        if (isSubjectItScope(subject)) addAnnualValue(annualItRevenueExcl, index, exclValue);
      } else {
        addAnnualValue(annualCostIncl, index, inclValue);
        addAnnualValue(annualCostExcl, index, exclValue);
        if (isSubjectItScope(subject)) addAnnualValue(annualItCostExcl, index, exclValue);
      }
    });
  });

  return {
    source: "subject_funding_plans",
    annualRevenueIncl,
    annualCostIncl,
    annualRevenueExcl,
    annualCostExcl,
    annualItRevenueExcl,
    annualItCostExcl,
    annualNetIncl: annualRevenueIncl.map((value, index) => roundSignedFundingMoney(value - annualCostIncl[index])),
    annualNetExcl: annualRevenueExcl.map((value, index) => roundSignedFundingMoney(value - annualCostExcl[index])),
  };
};

export const upsertSubjectFundingPlan = (
  plans: SubjectFundingPlans,
  plan: SubjectFundingPlan,
): SubjectFundingPlans => ({
  ...plans,
  [plan.id]: normalizeSubjectFundingPlan(plan) || plan,
});

export const removeSubjectFundingPlan = (
  plans: SubjectFundingPlans,
  subjectRef: SubjectFundingSubjectRef,
): SubjectFundingPlans => {
  const id = createSubjectFundingPlanId(subjectRef);
  if (!plans[id]) return plans;
  const nextPlans = { ...plans };
  delete nextPlans[id];
  return nextPlans;
};

/**
 * Synchronize a single subject's funding plan to a new inclusive amount.
 *
 * Rules:
 * - amount > 0, plan exists:  proportionally scale annualInclValues, preserve mode
 * - amount > 0, plan missing: auto-create an upfront plan
 * - amount <= 0, plan exists: remove the plan so the subject returns to "unmaintained"
 * - amount <= 0, plan missing: no-op
 *
 * All arithmetic uses integer-cents internally to avoid floating-point drift.
 */
export const syncSubjectFundingPlanToAmount = (
  plans: SubjectFundingPlans,
  subjectRef: SubjectFundingSubjectRef,
  newAmountIncl: number,
  reason: SubjectFundingPlanLastChangeReason = "manual_amount_sync"
): SubjectFundingPlans => {
  const id = createSubjectFundingPlanId(subjectRef);
  const existing = plans[id] ?? null;
  const newCents = toMoneyCents(newAmountIncl);

  // amount <= 0
  if (newCents <= 0) {
    if (!existing) return plans;
    const nextPlans = { ...plans };
    delete nextPlans[id];
    return nextPlans;
  }

  // amount > 0, plan missing → auto-create upfront
  if (!existing) {
    return {
      ...plans,
      [id]: createDefaultSubjectFundingPlan(subjectRef, newCents / 100),
    };
  }

  // amount > 0, proportional plan → re-derive yearly amounts from stored percentages
  if (existing.mode === "proportional" && sumAnnualPercentages(existing.annualPercentages) > 0) {
    return {
      ...plans,
      [id]: {
        ...existing,
        annualInclValues: buildProportionalAnnualInclValues(newCents / 100, existing.annualPercentages),
        lastValidAnnualInclValues: existing.annualInclValues,
        enabled: true,
        lastChangeReason: reason,
        lastChangedAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      },
    };
  }

  // amount > 0, plan exists → proportional scale
  const existingTotalCents = existing.annualInclValues.reduce((sum, val) => sum + Math.round(val * 100), 0);
  const isRecoveringFromZero = existingTotalCents === 0;
  const baseValues = (isRecoveringFromZero && existing.lastValidAnnualInclValues)
    ? normalizeAnnualInclValues(existing.lastValidAnnualInclValues)
    : existing.annualInclValues;
  const baseTotalCents = baseValues.reduce((sum, val) => sum + Math.round(val * 100), 0);

  if (baseTotalCents === 0) {
    // Fallback to upfront if even the base is zero
    return {
      ...plans,
      [id]: {
        ...existing,
        mode: "upfront",
        annualInclValues: buildUpfrontAnnualInclValues(newCents / 100),
        lastValidAnnualInclValues: existing.annualInclValues,
        enabled: true,
        lastChangeReason: isRecoveringFromZero ? "restored_after_zero" : reason,
        lastChangedAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      },
    };
  }

  // Proportional scale based on baseValues
  let currentTotal = 0;
  const scaledCents = baseValues.map(val => {
    const centValue = Math.round(val * 100);
    const scaled = Math.round((centValue / baseTotalCents) * newCents);
    currentTotal += scaled;
    return scaled;
  });

  const diff = newCents - currentTotal;
  for (let i = PLAN_YEARS - 1; i >= 0; i--) {
    if (scaledCents[i] !== 0) {
      scaledCents[i] += diff;
      break;
    }
  }
  // If no non-zero found, add to first year
  if (diff !== 0 && scaledCents.every(c => c === 0)) {
    scaledCents[0] += diff;
  }

  return {
    ...plans,
    [id]: {
      ...existing,
      annualInclValues: scaledCents.map(c => c / 100),
      lastValidAnnualInclValues: existing.annualInclValues,
      enabled: true,
      lastChangeReason: isRecoveringFromZero ? "restored_after_zero" : reason,
      lastChangedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    },
  };
};

/**
 * Batch-synchronize multiple subjects' funding plans in a single pass.
 * Each update is applied sequentially so later updates see earlier results.
 */
export const syncSubjectFundingPlansToAmounts = (
  plans: SubjectFundingPlans,
  updates: Array<{ subjectRef: SubjectFundingSubjectRef; newAmountIncl: number; reason?: SubjectFundingPlanLastChangeReason }>
): SubjectFundingPlans => {
  return updates.reduce((acc, update) => {
    return syncSubjectFundingPlanToAmount(acc, update.subjectRef, update.newAmountIncl, update.reason ?? "manual_amount_sync");
  }, plans);
};

export const initializeMissingSubjectFundingPlans = (
  plans: SubjectFundingPlans,
  subjects: Array<{ subjectRef: SubjectFundingSubjectRef; amountIncl: number }>
): SubjectFundingPlans => {
  const result = { ...plans };
  let changed = false;

  for (const { subjectRef, amountIncl } of subjects) {
    if (amountIncl <= 0) continue;
    const id = createSubjectFundingPlanId(subjectRef);
    if (!result[id]) {
      result[id] = {
        id,
        subjectRef,
        mode: "upfront",
        annualInclValues: buildUpfrontAnnualInclValues(amountIncl),
        enabled: true,
        source: "manual",
        lastChangeReason: "auto_created_upfront",
        lastChangedAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      };
      changed = true;
    }
  }

  return changed ? result : plans;
};

export const migrateLegacySubjectFundingPlans = (
  subjects: SubjectFundingPlanCoverageSubject[],
  plans: SubjectFundingPlans,
  currentMigrationVersion?: number | null,
): SubjectFundingPlanMigrationResult => {
  const normalizedPlans = normalizeSubjectFundingPlans(plans);

  if (Number(currentMigrationVersion) >= SUBJECT_FUNDING_PLAN_MIGRATION_VERSION) {
    const coverage = validateSubjectFundingPlanCoverage(subjects, normalizedPlans);
    return {
      plans: normalizedPlans,
      changed: false,
      completed: true,
      migrationVersion: SUBJECT_FUNDING_PLAN_MIGRATION_VERSION,
      coverage,
    };
  }

  let changed = false;
  const migratedPlans: SubjectFundingPlans = { ...normalizedPlans };

  subjects.forEach(subject => {
    if (toMoneyCents(subject.subjectAmountIncl) <= 0) return;
    const id = createSubjectFundingPlanId(subject.subjectRef);
    if (migratedPlans[id]) return;
    migratedPlans[id] = createLegacyMigrationSubjectFundingPlan(
      subject.subjectRef,
      subject.subjectAmountIncl,
    );
    changed = true;
  });

  const coverage = validateSubjectFundingPlanCoverage(subjects, migratedPlans);
  const completed = coverage.valid;

  return {
    plans: changed ? migratedPlans : normalizedPlans,
    changed,
    completed,
    migrationVersion: completed ? SUBJECT_FUNDING_PLAN_MIGRATION_VERSION : undefined,
    coverage,
  };
};

export interface SubjectFundingAnnualContribution {
  yearIndex: number;
  side: "revenue" | "cost";
  subjectRef: SubjectFundingSubjectRef;
  subjectDisplayName: string;
  annualInclAmount: number;
  annualExclAmount: number;
  taxRate: number;
}

export const buildAnnualCashflowSubjectContributions = (
  plans: SubjectFundingPlans,
  activeSubjects: SubjectFundingPlanCoverageSubject[]
): SubjectFundingAnnualContribution[][] => {
  const result: SubjectFundingAnnualContribution[][] = Array.from({ length: PLAN_YEARS }, () => []);

  for (const subject of activeSubjects) {
    if (subject.subjectAmountIncl <= 0) continue;
    const id = createSubjectFundingPlanId(subject.subjectRef);
    const plan = plans[id];
    if (!plan || !plan.enabled) continue; // Should be caught by validation earlier

    const annualValues = normalizeAnnualInclValues(plan.annualInclValues);
    // 与 buildAnnualCashflowFromSubjectFundingPlans 同源：先还原整笔、再按年分摊，
    // 保证下钻每年不含税之和等于科目还原后的不含税总额。
    const divisor = 1 + subject.taxRate / 100;
    const inclCentsByYear = annualValues.map(annualIncl => toMoneyCents(annualIncl));
    const exclCentsByYear = distributeExclCentsByInclWeights(inclCentsByYear, divisor);
    for (let yearIndex = 0; yearIndex < PLAN_YEARS; yearIndex++) {
      const incl = annualValues[yearIndex];
      if (incl === 0) continue;

      const excl = exclCentsByYear[yearIndex] / 100;
      result[yearIndex].push({
        yearIndex,
        side: subject.subjectRef.side,
        subjectRef: subject.subjectRef,
        subjectDisplayName: subject.displayName,
        annualInclAmount: incl,
        annualExclAmount: excl,
        taxRate: subject.taxRate,
      });
    }
  }

  return result;
};
