import Decimal from "decimal.js";
import {
  ICT_SUBJECT_DEFINITIONS,
  type IctSubjectDefinition,
} from "./ictSubjectCatalog";

export const DEFAULT_SELECTION_FEE_TARGET_SUBJECT_CODE = "cost_it_integration";
export const SELECTION_FEE_SERVICE_SUBJECT_CODE = "cost_it_bidding";

export const SELECTION_FEE_TARGET_SUBJECTS = ICT_SUBJECT_DEFINITIONS.filter(
  subject => subject.side === "cost" && subject.subjectCode !== SELECTION_FEE_SERVICE_SUBJECT_CODE,
);

export const normalizeSelectionFeeTargetSubjectCode = (value: unknown) => {
  const candidate = String(value ?? "").trim();
  return SELECTION_FEE_TARGET_SUBJECTS.some(subject => subject.subjectCode === candidate)
    ? candidate
    : DEFAULT_SELECTION_FEE_TARGET_SUBJECT_CODE;
};

export const resolveSelectionFeeTargetSubject = (value: unknown): IctSubjectDefinition => {
  const subjectCode = normalizeSelectionFeeTargetSubjectCode(value);
  return SELECTION_FEE_TARGET_SUBJECTS.find(subject => subject.subjectCode === subjectCode)
    ?? SELECTION_FEE_TARGET_SUBJECTS.find(subject => subject.subjectCode === DEFAULT_SELECTION_FEE_TARGET_SUBJECT_CODE)!;
};

const parseMoney = (value: unknown): Decimal | null => {
  try {
    const parsed = new Decimal(String(value ?? "").trim() || "0");
    return parsed.isFinite() ? parsed : null;
  } catch {
    return null;
  }
};

export interface SelectionFeeWriteAmounts {
  valid: boolean;
  targetIncl: number;
  serviceFeeIncl: number;
  limitIncl: number;
  message: string | null;
}

/**
 * 甄选服务费由供应商承担且已包含在最高限价中：
 * 目标投入科目 = 最高限价 - 甄选服务费（等价于供应商报价 + 上浮）；
 * 中标服务费 = 甄选服务费。
 */
export const calculateSelectionFeeWriteAmounts = (
  limitValue: unknown,
  serviceFeeValue: unknown,
): SelectionFeeWriteAmounts => {
  const limit = parseMoney(limitValue);
  const serviceFee = parseMoney(serviceFeeValue);
  if (!limit || !serviceFee || limit.isNegative() || serviceFee.isNegative()) {
    return {
      valid: false,
      targetIncl: 0,
      serviceFeeIncl: 0,
      limitIncl: 0,
      message: "甄选最高限价或甄选服务费不是有效的非负金额。",
    };
  }

  const roundedLimit = limit.toDecimalPlaces(2, Decimal.ROUND_HALF_UP);
  const roundedServiceFee = serviceFee.toDecimalPlaces(2, Decimal.ROUND_HALF_UP);
  const target = roundedLimit.minus(roundedServiceFee);
  if (target.isNegative()) {
    return {
      valid: false,
      targetIncl: 0,
      serviceFeeIncl: Number(roundedServiceFee.toFixed(2)),
      limitIncl: Number(roundedLimit.toFixed(2)),
      message: "甄选服务费不能大于甄选最高限价。",
    };
  }

  return {
    valid: true,
    targetIncl: Number(target.toFixed(2)),
    serviceFeeIncl: Number(roundedServiceFee.toFixed(2)),
    limitIncl: Number(roundedLimit.toFixed(2)),
    message: null,
  };
};
