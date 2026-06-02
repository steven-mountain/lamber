import Decimal from "decimal.js";
import type {
  IctSubjectDefinition,
  IctSubjectGroupId,
  IctTaxItemLike,
} from "./ictSubjectCatalog";

export type BalanceAllocationSide = "revenue" | "investment";

export interface BalanceSubjectRef {
  subjectCode: string;
  groupId: IctSubjectGroupId;
  key: string;
}

export interface BalanceAllocationRule {
  enabled: boolean;
  totalInclAmount: number | null;
  balancingSubject: BalanceSubjectRef | null;
}

export interface BalanceAllocationState {
  revenue: BalanceAllocationRule;
  investment: BalanceAllocationRule;
}

export interface BalanceSubjectItem {
  subject: IctSubjectDefinition;
  item: IctTaxItemLike | null;
}

export type BalanceRuleStatus =
  | "disabled"
  | "missing_total"
  | "missing_subject"
  | "subject_unavailable"
  | "negative"
  | "valid";

export interface BalanceRuleEvaluation {
  side: BalanceAllocationSide;
  status: BalanceRuleStatus;
  canApply: boolean;
  message: string | null;
  otherAllocated: number;
  autoAmount: number | null;
  balancingSubject: IctSubjectDefinition | null;
  balancingItem: IctTaxItemLike | null;
}

export const createDefaultBalanceAllocationRule = (): BalanceAllocationRule => ({
  enabled: false,
  totalInclAmount: null,
  balancingSubject: null,
});

export const createDefaultBalanceAllocationState = (): BalanceAllocationState => ({
  revenue: createDefaultBalanceAllocationRule(),
  investment: createDefaultBalanceAllocationRule(),
});

const parseNullableAmount = (value: unknown): number | null => {
  if (value === null || value === undefined || value === "") return null;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : null;
};

const normalizeSubjectRef = (value: unknown): BalanceSubjectRef | null => {
  if (!value || typeof value !== "object") return null;
  const raw = value as Record<string, unknown>;
  const subjectCode = String(raw.subjectCode ?? raw.subject_code ?? "").trim();
  const groupId = String(raw.groupId ?? raw.group_id ?? "").trim() as IctSubjectGroupId;
  const key = String(raw.key ?? "").trim();
  if (!subjectCode || !groupId || !key) return null;
  return { subjectCode, groupId, key };
};

export const normalizeBalanceAllocationRule = (value: unknown): BalanceAllocationRule => {
  if (!value || typeof value !== "object") return createDefaultBalanceAllocationRule();
  const raw = value as Record<string, unknown>;
  const totalInclAmount = parseNullableAmount(raw.totalInclAmount ?? raw.total_incl_amount);
  const balancingSubject = normalizeSubjectRef(raw.balancingSubject ?? raw.balancing_subject);
  return {
    enabled: Boolean(raw.enabled || totalInclAmount !== null || balancingSubject),
    totalInclAmount,
    balancingSubject,
  };
};

export const normalizeBalanceAllocationState = (value: unknown): BalanceAllocationState => {
  if (!value || typeof value !== "object") return createDefaultBalanceAllocationState();
  const raw = value as Record<string, unknown>;
  return {
    revenue: normalizeBalanceAllocationRule(
      raw.revenue ?? raw.revenueBalanceRule ?? raw.revenue_balance_rule,
    ),
    investment: normalizeBalanceAllocationRule(
      raw.investment ?? raw.investmentBalanceRule ?? raw.investment_balance_rule,
    ),
  };
};

export const getBalanceSubjectRef = (subject: IctSubjectDefinition): BalanceSubjectRef => ({
  subjectCode: subject.subjectCode,
  groupId: subject.groupId,
  key: subject.key,
});

export const getBalanceSubjectRefKey = (ref: BalanceSubjectRef | null | undefined) =>
  ref ? `${ref.groupId}.${ref.key}` : "";

export const isBalanceSubjectMatch = (
  ref: BalanceSubjectRef | null | undefined,
  subject: IctSubjectDefinition,
) => {
  if (!ref) return false;
  if (ref.subjectCode) return ref.subjectCode === subject.subjectCode;
  return ref.groupId === subject.groupId && ref.key === subject.key;
};

export const serializeBalanceSubjectRef = (ref: BalanceSubjectRef | null) => (
  ref
    ? {
        subject_code: ref.subjectCode,
        group_id: ref.groupId,
        key: ref.key,
      }
    : null
);

export const serializeBalanceAllocationRule = (rule: BalanceAllocationRule) => ({
  enabled: rule.enabled,
  total_incl_amount: rule.totalInclAmount,
  balancing_subject: serializeBalanceSubjectRef(rule.balancingSubject),
});

export const serializeBalanceAllocationState = (state: BalanceAllocationState) => ({
  revenue: serializeBalanceAllocationRule(state.revenue),
  investment: serializeBalanceAllocationRule(state.investment),
});

const toMoneyDecimal = (value: unknown) => {
  try {
    const decimal = new Decimal(String(value ?? 0));
    return decimal.isFinite() ? decimal : new Decimal(0);
  } catch {
    return new Decimal(0);
  }
};

const toRoundedMoneyNumber = (value: Decimal) =>
  Number(value.toDecimalPlaces(2, Decimal.ROUND_HALF_UP).toFixed(2));

const sideName = (side: BalanceAllocationSide) => (side === "revenue" ? "收入" : "投入");

export const evaluateBalanceRule = (
  side: BalanceAllocationSide,
  rule: BalanceAllocationRule,
  subjects: BalanceSubjectItem[],
): BalanceRuleEvaluation => {
  const base = {
    side,
    canApply: false,
    otherAllocated: 0,
    autoAmount: null,
    balancingSubject: null,
    balancingItem: null,
  };

  if (!rule.enabled) {
    return { ...base, status: "disabled", message: null };
  }

  if (rule.totalInclAmount === null) {
    return {
      ...base,
      status: "missing_total",
      message: `请输入${sideName(side)}含税总金额后自动承接。`,
    };
  }

  if (!rule.balancingSubject) {
    return {
      ...base,
      status: "missing_subject",
      message: `请选择${sideName(side)}侧差额承接科目。`,
    };
  }

  const balancingRow = subjects.find(({ subject }) => isBalanceSubjectMatch(rule.balancingSubject, subject));
  if (!balancingRow) {
    return {
      ...base,
      status: "subject_unavailable",
      message: "当前差额承接科目已不可用，请重新选择。",
    };
  }

  const otherAllocatedDecimal = subjects.reduce((sum, row) => {
    if (isBalanceSubjectMatch(rule.balancingSubject, row.subject)) return sum;
    return sum.plus(toMoneyDecimal(row.item?.incl));
  }, new Decimal(0));
  const totalDecimal = toMoneyDecimal(rule.totalInclAmount);
  const autoAmountDecimal = totalDecimal
    .minus(otherAllocatedDecimal)
    .toDecimalPlaces(2, Decimal.ROUND_HALF_UP);
  const otherAllocated = toRoundedMoneyNumber(otherAllocatedDecimal);

  if (autoAmountDecimal.isNegative()) {
    return {
      ...base,
      status: "negative",
      message: side === "revenue"
        ? "其他收入科目合计已超过收入含税总金额，请调整金额或总额。"
        : "其他投入科目合计已超过投入含税总金额，请调整金额或总额。",
      otherAllocated,
      autoAmount: null,
      balancingSubject: balancingRow.subject,
      balancingItem: balancingRow.item,
    };
  }

  return {
    side,
    status: "valid",
    canApply: true,
    message: null,
    otherAllocated,
    autoAmount: toRoundedMoneyNumber(autoAmountDecimal),
    balancingSubject: balancingRow.subject,
    balancingItem: balancingRow.item,
  };
};
