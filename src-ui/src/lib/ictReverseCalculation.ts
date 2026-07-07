import {
  getSubjectExcelDisplayName,
  type IctSubjectDefinition,
  type IctSubjectGroupId,
  type IctSubjectSide,
  type IctTaxItemLike,
} from "./ictSubjectCatalog";
import {
  isBalanceSubjectMatch,
  type BalanceRuleEvaluation,
} from "./ictBalanceAllocation";
import { normalizeTaxPairFromIncl } from "./taxAmount";

export type ReverseSubjectRef = {
  side: IctSubjectSide;
  groupId: IctSubjectGroupId;
  key: string;
  subjectCode: string;
};

export type ReverseSubjectOption = {
  ref: ReverseSubjectRef;
  subject: IctSubjectDefinition;
  item: IctTaxItemLike | null;
  displayName: string;
  disabledReason?: string;
};

export type ReverseSubjectState = {
  revIt: Record<string, IctTaxItemLike>;
  revCt: Record<string, IctTaxItemLike>;
  revNonItCt: IctTaxItemLike;
  costIt: Record<string, IctTaxItemLike>;
  costCt: Record<string, IctTaxItemLike>;
  costMix: Record<string, IctTaxItemLike>;
};

export type ReverseCalculationMode = "normal" | "locked_total_structure" | "blocked";

export type LockedTotalStructureContext = {
  side: IctSubjectSide;
  sideLabel: string;
  totalInclAmount: number;
  targetSubject: IctSubjectDefinition;
  targetItem: IctTaxItemLike | null;
  targetDisplayName: string;
  balancingSubject: IctSubjectDefinition;
  balancingItem: IctTaxItemLike | null;
  balancingDisplayName: string;
  fixedOtherInclAmount: number;
  reallocatablePoolInclAmount: number;
  beforeTargetInclAmount: number;
  beforeBalancingInclAmount: number;
};

export type ResolvedReverseCalculationContext =
  | { mode: "normal"; option: ReverseSubjectOption }
  | { mode: "locked_total_structure"; option: ReverseSubjectOption; structure: LockedTotalStructureContext }
  | { mode: "blocked"; option: ReverseSubjectOption | null; message: string };

export const getReverseSubjectRef = (subject: IctSubjectDefinition): ReverseSubjectRef => ({
  side: subject.side,
  groupId: subject.groupId,
  key: subject.key,
  subjectCode: subject.subjectCode,
});

export const getReverseSubjectRefKey = (ref: ReverseSubjectRef | null | undefined) =>
  ref ? `${ref.side}:${ref.subjectCode || `${ref.groupId}.${ref.key}`}` : "";

export const isReverseSubjectRefMatch = (
  ref: ReverseSubjectRef | null | undefined,
  subject: IctSubjectDefinition,
) => {
  if (!ref) return false;
  if (ref.subjectCode) return ref.subjectCode === subject.subjectCode;
  return ref.side === subject.side && ref.groupId === subject.groupId && ref.key === subject.key;
};

const isSameSubject = (left: IctSubjectDefinition, right: IctSubjectDefinition) =>
  left.subjectCode === right.subjectCode
  || (left.groupId === right.groupId && left.key === right.key);

export const getSubjectItemFromState = (
  groups: ReverseSubjectState,
  subject: IctSubjectDefinition,
): IctTaxItemLike | null => {
  if (subject.groupId === "revIt") return groups.revIt[subject.key] || null;
  if (subject.groupId === "revCt") return groups.revCt[subject.key] || null;
  if (subject.groupId === "revNonItCt") return groups.revNonItCt || null;
  if (subject.groupId === "costIt") return groups.costIt[subject.key] || null;
  if (subject.groupId === "costCt") return groups.costCt[subject.key] || null;
  if (subject.groupId === "costMix") return groups.costMix[subject.key] || null;
  return null;
};

const sideLabel = (side: IctSubjectSide) => (side === "revenue" ? "收入" : "投入");

const isEffectiveBalanceRule = (evaluation: BalanceRuleEvaluation | null | undefined) =>
  Boolean(evaluation?.status === "valid" && evaluation.canApply && evaluation.balancingSubject);

const readIncl = (item: IctTaxItemLike | null | undefined) => {
  const value = Number(item?.incl ?? 0);
  return Number.isFinite(value) ? value : 0;
};

const roundMoney = (value: number) => Number((Number.isFinite(value) ? value : 0).toFixed(2));

export const getReverseEligibleSubjects = (
  side: IctSubjectSide,
  subjects: Array<{ subject: IctSubjectDefinition; item: IctTaxItemLike | null }>,
  sameSideBalanceEvaluation?: BalanceRuleEvaluation | null,
): ReverseSubjectOption[] => {
  const balanceActive = isEffectiveBalanceRule(sameSideBalanceEvaluation);
  return subjects
    .filter(({ subject }) => subject.side === side)
    .map(({ subject, item }) => {
      const isBalancingSubject = balanceActive
        && Boolean(sameSideBalanceEvaluation?.balancingSubject)
        && isBalanceSubjectMatch(sameSideBalanceEvaluation?.balancingSubject, subject);
      return {
        ref: getReverseSubjectRef(subject),
        subject,
        item,
        displayName: getSubjectExcelDisplayName(subject, item),
        disabledReason: isBalancingSubject
          ? `${sideLabel(side)}差额承接科目由总金额自动计算，不能作为反算目标。`
          : undefined,
      };
    });
};

export const findReverseSubjectOption = (
  options: ReverseSubjectOption[],
  refKey: string,
) => options.find(option => getReverseSubjectRefKey(option.ref) === refKey) || null;

export const readSubjectInclAmount = (groups: ReverseSubjectState, ref: ReverseSubjectRef) => {
  return Number(findSubjectItemByRef(groups, ref)?.incl ?? 0);
};

const findSubjectItemByRef = (groups: ReverseSubjectState, ref: ReverseSubjectRef) => {
  if (ref.groupId === "revIt") return groups.revIt[ref.key] || null;
  if (ref.groupId === "revCt") return groups.revCt[ref.key] || null;
  if (ref.groupId === "revNonItCt") return groups.revNonItCt || null;
  if (ref.groupId === "costIt") return groups.costIt[ref.key] || null;
  if (ref.groupId === "costCt") return groups.costCt[ref.key] || null;
  if (ref.groupId === "costMix") return groups.costMix[ref.key] || null;
  return null;
};

const taxItemFromIncl = (current: IctTaxItemLike | null | undefined, incl: number): IctTaxItemLike => {
  const tax = Number(current?.tax ?? 0);
  // 反算候选仅用于内部评估，不落库：含税保留候选值，不含税按财务口径精确还原。
  // 最终写入时由 useIctCalculations 检查可表示性（拒绝或按设置自动修正）。
  const pair = normalizeTaxPairFromIncl(Number.isFinite(incl) ? incl : 0, tax);
  return {
    ...(current || {}),
    incl: pair.enteredIncl,
    tax,
    excl: pair.excl,
  };
};

const setSubjectItem = (
  groups: ReverseSubjectState,
  subject: IctSubjectDefinition,
  item: IctTaxItemLike,
) => {
  if (subject.groupId === "revIt") groups.revIt = { ...groups.revIt, [subject.key]: item };
  else if (subject.groupId === "revCt") groups.revCt = { ...groups.revCt, [subject.key]: item };
  else if (subject.groupId === "revNonItCt") groups.revNonItCt = item;
  else if (subject.groupId === "costIt") groups.costIt = { ...groups.costIt, [subject.key]: item };
  else if (subject.groupId === "costCt") groups.costCt = { ...groups.costCt, [subject.key]: item };
  else if (subject.groupId === "costMix") groups.costMix = { ...groups.costMix, [subject.key]: item };
};

const applyRevCtPairedAmount = (
  groups: ReverseSubjectState,
  key: string,
  amount: number,
) => {
  if (key !== "product" && key !== "line") return;
  const pairedKey = key === "product" ? "other" : "bandwidth";
  groups.costCt = {
    ...groups.costCt,
    [pairedKey]: taxItemFromIncl(groups.costCt[pairedKey], amount),
  };
};

export const applySubjectInclAmountToState = (
  groups: ReverseSubjectState,
  subject: IctSubjectDefinition,
  amount: number,
): ReverseSubjectState => {
  const next: ReverseSubjectState = {
    revIt: groups.revIt,
    revCt: groups.revCt,
    revNonItCt: groups.revNonItCt,
    costIt: groups.costIt,
    costCt: groups.costCt,
    costMix: groups.costMix,
  };
  const current = getSubjectItemFromState(next, subject);
  setSubjectItem(next, subject, taxItemFromIncl(current, amount));

  if (subject.groupId === "revCt") {
    applyRevCtPairedAmount(next, subject.key, amount);
  }

  return next;
};

export const applyLockedTotalStructureAmountsToState = (
  groups: ReverseSubjectState,
  targetSubject: IctSubjectDefinition,
  targetAmount: number,
  balancingSubject: IctSubjectDefinition,
  balancingAmount: number,
): ReverseSubjectState => applySubjectInclAmountToState(
  applySubjectInclAmountToState(groups, targetSubject, targetAmount),
  balancingSubject,
  balancingAmount,
);

export const buildLockedTotalStructureContext = (params: {
  option: ReverseSubjectOption;
  subjects: Array<{ subject: IctSubjectDefinition; item: IctTaxItemLike | null }>;
  sameSideBalanceEvaluation?: BalanceRuleEvaluation | null;
}): LockedTotalStructureContext | null => {
  const { option, subjects, sameSideBalanceEvaluation } = params;
  if (!isEffectiveBalanceRule(sameSideBalanceEvaluation)) return null;
  const balancingSubject = sameSideBalanceEvaluation?.balancingSubject;
  if (!balancingSubject || isSameSubject(option.subject, balancingSubject)) return null;

  const sideRows = subjects.filter(row => row.subject.side === option.ref.side);
  const fixedOtherInclAmount = sideRows.reduce((sum, row) => {
    if (isSameSubject(row.subject, option.subject) || isSameSubject(row.subject, balancingSubject)) {
      return sum;
    }
    return sum + readIncl(row.item);
  }, 0);
  const totalInclAmount = roundMoney(
    Number(sameSideBalanceEvaluation?.autoAmount ?? 0)
    + Number(sameSideBalanceEvaluation?.otherAllocated ?? 0),
  );
  const reallocatablePoolInclAmount = roundMoney(totalInclAmount - fixedOtherInclAmount);
  const balancingItem = sideRows.find(row => isSameSubject(row.subject, balancingSubject))?.item
    ?? sameSideBalanceEvaluation?.balancingItem
    ?? null;

  if (reallocatablePoolInclAmount < -0.004) return null;

  return {
    side: option.ref.side,
    sideLabel: sideLabel(option.ref.side),
    totalInclAmount,
    targetSubject: option.subject,
    targetItem: option.item,
    targetDisplayName: option.displayName,
    balancingSubject,
    balancingItem,
    balancingDisplayName: getSubjectExcelDisplayName(balancingSubject, balancingItem),
    fixedOtherInclAmount: roundMoney(fixedOtherInclAmount),
    reallocatablePoolInclAmount: Math.max(0, reallocatablePoolInclAmount),
    beforeTargetInclAmount: readIncl(option.item),
    beforeBalancingInclAmount: readIncl(balancingItem),
  };
};

export const resolveReverseCalculationContext = (params: {
  option: ReverseSubjectOption | null;
  subjects: Array<{ subject: IctSubjectDefinition; item: IctTaxItemLike | null }>;
  sameSideBalanceEvaluation?: BalanceRuleEvaluation | null;
}): ResolvedReverseCalculationContext => {
  const { option, subjects, sameSideBalanceEvaluation } = params;
  if (!option) {
    return { mode: "blocked", option: null, message: "请选择需要反算的计费科目。" };
  }
  if (option.disabledReason) {
    return { mode: "blocked", option, message: option.disabledReason };
  }
  if (sameSideBalanceEvaluation?.status === "negative" && sameSideBalanceEvaluation.message) {
    return { mode: "blocked", option, message: sameSideBalanceEvaluation.message };
  }
  if (!isEffectiveBalanceRule(sameSideBalanceEvaluation)) {
    return { mode: "normal", option };
  }
  if (sameSideBalanceEvaluation?.balancingSubject && isSameSubject(option.subject, sameSideBalanceEvaluation.balancingSubject)) {
    return {
      mode: "blocked",
      option,
      message: `${sideLabel(option.ref.side)}差额承接科目由总金额自动计算，不能作为反算目标。`,
    };
  }

  const structure = buildLockedTotalStructureContext({ option, subjects, sameSideBalanceEvaluation });
  if (!structure) {
    return {
      mode: "blocked",
      option,
      message: `当前${sideLabel(option.ref.side)}总金额锁定规则不完整或不可用，无法执行结构反算。`,
    };
  }

  return { mode: "locked_total_structure", option, structure };
};

export const buildLockedTotalStructureSamplePoints = (
  reallocatablePoolInclAmount: number,
  currentTargetInclAmount: number,
) => {
  const pool = Math.max(0, roundMoney(reallocatablePoolInclAmount));
  const points = new Set<number>([0, pool, Math.min(pool, Math.max(0, roundMoney(currentTargetInclAmount)))]);
  for (let i = 1; i < 10; i++) {
    points.add(roundMoney(pool * i / 10));
  }
  return Array.from(points).sort((a, b) => a - b);
};

export const validateReverseCalculationContext = (params: {
  option: ReverseSubjectOption | null;
  sameSideBalanceEvaluation?: BalanceRuleEvaluation | null;
}) => {
  const { option, sameSideBalanceEvaluation } = params;
  if (!option) return "请选择需要反算的计费科目。";
  if (option.disabledReason) return option.disabledReason;

  if (sameSideBalanceEvaluation?.status === "negative" && sameSideBalanceEvaluation.message) {
    return sameSideBalanceEvaluation.message;
  }

  if (isEffectiveBalanceRule(sameSideBalanceEvaluation)) {
    return `当前${sideLabel(option.ref.side)}总金额已锁定，并启用了差额承接。${sideLabel(option.ref.side)}科目反算会涉及结构调整，本阶段暂不支持。请先清空${sideLabel(option.ref.side)}总金额或差额承接科目后再执行，或改为反算另一侧科目。`;
  }

  return null;
};
