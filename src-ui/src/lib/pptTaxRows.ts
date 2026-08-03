import {
  getProjectDataSubjectItem,
  ICT_SUBJECT_DEFINITIONS,
  type IctSubjectGroupId,
} from "./ictSubjectCatalog";
import {
  restoreTaxSplitParts,
  roundMoneyHalfUp,
  type TaxSplitPart,
} from "./taxAmount";

export type PptTaxRowMode = "merged" | "split";

export type PptTaxAmountRow = TaxSplitPart & {
  splitIndex: number | null;
  splitCount: number;
};

export type PptTaxSplitSummary = {
  subjectCount: number;
  addedRows: number;
};

type PptTaxItemLike = {
  incl?: unknown;
  excl?: unknown;
  tax?: unknown;
  tax_rate?: unknown;
  splitParts?: unknown;
  split_parts?: unknown;
};

const PPT_DETAIL_GROUPS = new Set<IctSubjectGroupId>([
  "revIt",
  "revCt",
  "costIt",
  "costCt",
]);

export const getValidPptTaxSplitParts = (
  item: PptTaxItemLike | null | undefined,
): TaxSplitPart[] | null => {
  if (!item) return null;
  return restoreTaxSplitParts(
    item.splitParts ?? item.split_parts,
    item.incl,
    item.tax ?? item.tax_rate,
  );
};

/**
 * PPT 默认按科目汇总展示；仅在用户明确选择拆分时展开为子笔行。
 * 合并模式仍优先使用有效子笔之和作为不含税金额，保证展示口径与测算一致。
 */
export const buildPptTaxAmountRows = (
  item: PptTaxItemLike | null | undefined,
  mode: PptTaxRowMode,
): PptTaxAmountRow[] => {
  if (!item) return [];
  const splitParts = getValidPptTaxSplitParts(item);
  if (mode === "split" && splitParts) {
    return splitParts.map((part, index) => ({
      ...part,
      splitIndex: index + 1,
      splitCount: splitParts.length,
    }));
  }

  const excl = splitParts
    ? splitParts.reduce((sum, part) => sum + part.excl, 0)
    : item.excl;
  return [{
    incl: roundMoneyHalfUp(item.incl),
    excl: roundMoneyHalfUp(excl),
    splitIndex: null,
    splitCount: 1,
  }];
};

/** 只统计当前 PPT 投资收益页实际展示的 IT/CT 收入与投入科目。 */
export const getPptTaxSplitSummary = (projectData: any): PptTaxSplitSummary =>
  ICT_SUBJECT_DEFINITIONS.reduce<PptTaxSplitSummary>((summary, subject) => {
    if (!PPT_DETAIL_GROUPS.has(subject.groupId)) return summary;
    const parts = getValidPptTaxSplitParts(
      getProjectDataSubjectItem(projectData, subject) as PptTaxItemLike | null,
    );
    if (!parts) return summary;
    summary.subjectCount += 1;
    summary.addedRows += parts.length - 1;
    return summary;
  }, { subjectCount: 0, addedRows: 0 });

export const formatPptSplitNote = (row: PptTaxAmountRow): string =>
  row.splitIndex === null ? "" : `拆分第${row.splitIndex}笔/共${row.splitCount}笔`;
