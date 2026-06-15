import Decimal from "decimal.js";
import type {
  AiComputeLineItemFundingPlan,
  AiComputeLineItemFundingPlanMode,
  AiComputeQuoteParameter,
} from "./types";

export const AI_COMPUTE_FUNDING_PLAN_YEARS = 10;
export const AI_COMPUTE_PROJECT_CYCLE_PARAMETER_ID = "years";
export const AI_COMPUTE_PROJECT_CYCLE_PARAMETER_KEY = "years";
export const AI_COMPUTE_DISCOUNT_RATE_PARAMETER_ID = "discount-rate";
export const AI_COMPUTE_DISCOUNT_RATE_PARAMETER_KEY = "discount_rate";
export const AI_COMPUTE_DEFAULT_DISCOUNT_RATE_PERCENT = 5.5;

const money = (value: Decimal.Value) =>
  new Decimal(value).toDecimalPlaces(2, Decimal.ROUND_HALF_UP).toNumber();

const normalizeAmount = (value: unknown) => {
  const amount = Number(value);
  return Number.isFinite(amount) ? money(Math.max(0, amount)) : 0;
};

export function normalizeAiComputeProjectCycleValue(value: unknown) {
  const rawYears = Number(value);
  if (!Number.isFinite(rawYears)) return AI_COMPUTE_FUNDING_PLAN_YEARS;
  return Math.max(1, Math.min(AI_COMPUTE_FUNDING_PLAN_YEARS, Math.floor(rawYears)));
}

export function isAiComputeProjectCycleParameter(parameter: Pick<AiComputeQuoteParameter, "id" | "key">) {
  return parameter.id === AI_COMPUTE_PROJECT_CYCLE_PARAMETER_ID
    || parameter.key === AI_COMPUTE_PROJECT_CYCLE_PARAMETER_KEY;
}

export function getAiComputeProjectCycleYears(parameters: AiComputeQuoteParameter[]) {
  const parameter = parameters.find(isAiComputeProjectCycleParameter);
  return normalizeAiComputeProjectCycleValue(parameter?.value);
}

export function normalizeAiComputeDiscountRatePercent(value: unknown) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return AI_COMPUTE_DEFAULT_DISCOUNT_RATE_PERCENT;
  return new Decimal(Math.max(0, Math.min(100, numeric))).toDecimalPlaces(4).toNumber();
}

export function isAiComputeDiscountRateParameter(parameter: Pick<AiComputeQuoteParameter, "id" | "key">) {
  return parameter.id === AI_COMPUTE_DISCOUNT_RATE_PARAMETER_ID
    || parameter.key === AI_COMPUTE_DISCOUNT_RATE_PARAMETER_KEY;
}

export function getAiComputeDiscountRatePercent(parameters: AiComputeQuoteParameter[]) {
  const parameter = parameters.find(isAiComputeDiscountRateParameter);
  return normalizeAiComputeDiscountRatePercent(parameter?.value);
}

export function getAiComputeDiscountRateDecimal(parameters: AiComputeQuoteParameter[]) {
  return new Decimal(getAiComputeDiscountRatePercent(parameters)).div(100).toNumber();
}

export function ictDiscountRateToAiComputePercent(value: unknown) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return AI_COMPUTE_DEFAULT_DISCOUNT_RATE_PERCENT;
  return normalizeAiComputeDiscountRatePercent(Math.abs(numeric) <= 1 ? numeric * 100 : numeric);
}

export function isAiComputeStableIctParameter(parameter: Pick<AiComputeQuoteParameter, "id" | "key">) {
  return isAiComputeProjectCycleParameter(parameter) || isAiComputeDiscountRateParameter(parameter);
}

export function createEmptyAiComputeYearlyAmounts(): Record<string, number> {
  return Object.fromEntries(
    Array.from({ length: AI_COMPUTE_FUNDING_PLAN_YEARS }, (_, index) => [String(index + 1), 0]),
  );
}

export function buildAiComputeFirstYearAmounts(totalAmount: number): Record<string, number> {
  return {
    ...createEmptyAiComputeYearlyAmounts(),
    "1": normalizeAmount(totalAmount),
  };
}

export function buildAiComputeEvenAmounts(
  totalAmount: number,
  projectCycleYears: number,
): Record<string, number> {
  const years = normalizeAiComputeProjectCycleValue(projectCycleYears);
  const totalCents = new Decimal(normalizeAmount(totalAmount)).mul(100).toDecimalPlaces(0).toNumber();
  const baseCents = Math.floor(totalCents / years);
  const remainder = totalCents - baseCents * years;
  const yearlyAmounts = createEmptyAiComputeYearlyAmounts();

  for (let index = 1; index <= years; index += 1) {
    yearlyAmounts[String(index)] = money(new Decimal(baseCents + (index === years ? remainder : 0)).div(100));
  }
  return yearlyAmounts;
}

export function createDefaultAiComputeFundingPlan(totalAmount: number): AiComputeLineItemFundingPlan {
  return {
    enabled: true,
    mode: "first_year",
    yearlyAmounts: buildAiComputeFirstYearAmounts(totalAmount),
  };
}

export function normalizeAiComputeFundingPlan(
  plan: AiComputeLineItemFundingPlan | null | undefined,
  totalAmount: number,
  projectCycleYears: number,
): AiComputeLineItemFundingPlan {
  if (!plan) return createDefaultAiComputeFundingPlan(totalAmount);
  const yearlyAmounts = createEmptyAiComputeYearlyAmounts();
  Object.keys(yearlyAmounts).forEach(year => {
    yearlyAmounts[year] = normalizeAmount(plan.yearlyAmounts?.[year]);
  });
  const mode: AiComputeLineItemFundingPlanMode = ["first_year", "even", "manual"].includes(plan.mode)
    ? plan.mode
    : "first_year";
  const normalized = {
    enabled: plan.enabled !== false,
    mode,
    yearlyAmounts,
  };
  return syncAiComputeFundingPlan(normalized, totalAmount, projectCycleYears);
}

export function syncAiComputeFundingPlan(
  plan: AiComputeLineItemFundingPlan,
  totalAmount: number,
  projectCycleYears: number,
): AiComputeLineItemFundingPlan {
  if (plan.mode === "manual") return plan;
  return {
    ...plan,
    yearlyAmounts: plan.mode === "even"
      ? buildAiComputeEvenAmounts(totalAmount, projectCycleYears)
      : buildAiComputeFirstYearAmounts(totalAmount),
  };
}

export function updateAiComputeFundingPlanMode(
  plan: AiComputeLineItemFundingPlan,
  mode: AiComputeLineItemFundingPlanMode,
  totalAmount: number,
  projectCycleYears: number,
): AiComputeLineItemFundingPlan {
  return syncAiComputeFundingPlan({ ...plan, mode }, totalAmount, projectCycleYears);
}

export function updateAiComputeFundingPlanYear(
  plan: AiComputeLineItemFundingPlan,
  year: number,
  amount: number,
): AiComputeLineItemFundingPlan {
  if (year < 1 || year > AI_COMPUTE_FUNDING_PLAN_YEARS) return plan;
  return {
    ...plan,
    mode: "manual",
    yearlyAmounts: {
      ...plan.yearlyAmounts,
      [String(year)]: normalizeAmount(amount),
    },
  };
}

export function sumAiComputeFundingPlan(plan: AiComputeLineItemFundingPlan) {
  return money(
    Object.values(plan.yearlyAmounts).reduce(
      (sum, amount) => sum.add(normalizeAmount(amount)),
      new Decimal(0),
    ),
  );
}

export function validateAiComputeFundingPlan(
  plan: AiComputeLineItemFundingPlan,
  totalAmount: number,
) {
  const plannedAmount = sumAiComputeFundingPlan(plan);
  const subjectAmount = normalizeAmount(totalAmount);
  const difference = money(new Decimal(subjectAmount).sub(plannedAmount));
  return {
    plannedAmount,
    subjectAmount,
    difference,
    consistent: Math.abs(difference) < 0.005,
  };
}
