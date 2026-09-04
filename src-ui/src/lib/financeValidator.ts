import Decimal from 'decimal.js';
import { splitInclAmountForExclDelta, type TaxSplitPart } from './taxAmount';

export interface TaxItem {
  incl: string | number;
  excl: string | number;
  tax: string | number; // Note: using "tax" key to match IctLifecycle.tsx state
  /** 拆分明细：存在时科目按两笔子金额逐笔校验（各自闭合），不再用合计跑反推核验。 */
  splitParts?: { incl: string | number; excl: string | number }[];
}

export interface ValidationReport {
  side: 'income' | 'expense';
  taxRate: number;
  key: string;
  expectedExcl: string;
  actualExcl: string;
  difference: string;
  /**
   * 出错的字段方向：
   * - 'excl'（默认）：录入不含税 ≠ round(含税 ÷ (1+税率))
   * - 'incl'：录入含税 ≠ round(不含税 × (1+税率))，即与财务系统展示口径不一致
   */
  field?: 'excl' | 'incl';
  /** 税率组汇总尾差可通过真实业务明细拆分归零时，提供只读候选，不自动改数。 */
  splitSuggestions?: TaxGroupSplitSuggestion[];
}

export interface TaxGroupSplitSuggestion {
  subjectKey: string;
  originalIncl: string;
  originalExcl: string;
  parts: TaxSplitPart[];
  exclAdjustment: string;
  differenceBefore: string;
  differenceAfter: string;
}

/**
 * Perform a strict 0-tolerance financial validation using decimal.js
 * @param revenueData Revenue data groups
 * @param costData Cost data groups
 * @returns Array of validation errors
 */
export function validateFinancialData(
  revenueData: Record<string, Record<string, TaxItem>>,
  costData: Record<string, Record<string, TaxItem>>
): { errors: ValidationReport[], totalDifference: string } {
  const errors: ValidationReport[] = [];
  let totalDifference = new Decimal(0);

  const validateGroup = (
    data: Record<string, Record<string, TaxItem>>,
    side: 'income' | 'expense'
  ) => {
    // Group all items by tax rate
    const taxGroups = new Map<number, { key: string, item: TaxItem, fromSplit?: boolean }[]>();

    for (const [category, items] of Object.entries(data)) {
      if (!items) continue;
      for (const [itemKey, item] of Object.entries(items)) {
        if (!item || item.incl === undefined || item.excl === undefined) continue;

        const incl = new Decimal(item.incl || 0);
        const excl = new Decimal(item.excl || 0);

        // Skip empty or completely 0 items
        if (incl.isZero() && excl.isZero()) continue;

        const rate = Number(item.tax) || 0;
        if (!taxGroups.has(rate)) {
          taxGroups.set(rate, []);
        }
        // 拆分科目：按两笔子金额逐笔参与校验（各自闭合），合计不跑
        // B/B2（合计含税 vs 不含税反推在拆分场景天然差 1 分，非错误）。
        const splitParts = Array.isArray(item.splitParts) && item.splitParts.length >= 2
          ? item.splitParts
          : null;
        if (splitParts) {
          for (const part of splitParts) {
            taxGroups.get(rate)!.push({
              key: `${category}.${itemKey}`,
              item: { incl: part.incl, excl: part.excl, tax: rate },
              fromSplit: true,
            });
          }
        } else {
          taxGroups.get(rate)!.push({ key: `${category}.${itemKey}`, item });
        }
      }
    }

    // Validate each tax group
    for (const [rate, items] of taxGroups.entries()) {
      let sumIncl = new Decimal(0);
      let sumExcl = new Decimal(0);
      let sumExpectedExcl = new Decimal(0);

      const taxDivisor = new Decimal(1).plus(new Decimal(rate).div(100));

      for (const { key, item } of items) {
        const incl = new Decimal(item.incl || 0);
        const excl = new Decimal(item.excl || 0);

        // Expected Pre-tax (预期税前) = 录入的含税金额 / (1 + 税率), rounded to 2 decimal places
        const expectedExcl = incl.div(taxDivisor).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);

        // Tail Difference (分项尾差) = 录入的不含税金额 - 预期税前金额
        const difference = excl.minus(expectedExcl);

        sumIncl = sumIncl.plus(incl);
        sumExcl = sumExcl.plus(excl);
        sumExpectedExcl = sumExpectedExcl.plus(expectedExcl);

        // B: Item-level Strict Validation
        if (!difference.isZero()) {
          totalDifference = totalDifference.plus(difference);
          errors.push({
            side,
            taxRate: rate,
            key,
            expectedExcl: expectedExcl.toFixed(2),
            actualExcl: excl.toFixed(2),
            difference: difference.toFixed(2),
            field: 'excl'
          });
        }

        // B2: 财务口径核验 —— 业务系统以不含税为准，含税展示值 = round(不含税 × (1+税率))。
        // 录入含税若与该反推值不一致（如 6% 下 1038 vs 1038.01），生成的材料会与业务系统差分。
        const expectedIncl = excl.mul(taxDivisor).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);
        const inclDifference = incl.minus(expectedIncl);
        if (!excl.isZero() && !inclDifference.isZero()) {
          totalDifference = totalDifference.plus(inclDifference);
          errors.push({
            side,
            taxRate: rate,
            key,
            expectedExcl: expectedIncl.toFixed(2),
            actualExcl: incl.toFixed(2),
            difference: inclDifference.toFixed(2),
            field: 'incl'
          });
        }
      }

      // C: Aggregate-level Strict Validation
      // C1: 倒算核验 = sum(excl) - round(sum(incl) / (1 + tax_rate), 2)
      const groupExpectedExclBySum = sumIncl.div(taxDivisor).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);
      const groupDifferenceC1 = sumExcl.minus(groupExpectedExclBySum);

      // C2: 累加核验 = sum(expectedExcl) === sum(excl)
      const groupDifferenceC2 = sumExpectedExcl.minus(sumExcl);

      // 含拆分子笔的税率组以逐行累加口径为准（C2）：round-of-sum 的 C1
      // 与逐行锚点在拆分场景可差 1 分，不作为错误。
      const hasSplit = items.some(entry => entry.fromSplit);
      if ((!hasSplit && !groupDifferenceC1.isZero()) || !groupDifferenceC2.isZero()) {
        const hasItemErrors = errors.some(e => e.side === side && e.taxRate === rate);
        if (!hasItemErrors) {
          // If no specific item error was caught, but the sum is wrong, it means there is a rounding mismatch at the group level
          // Force user to adjust.
          const groupDifference = hasSplit ? groupDifferenceC2 : groupDifferenceC1;
          const requiredExclAdjustment = groupDifference.negated();
          const splitSuggestions = hasSplit
            ? []
            : items
                .flatMap(({ key, item, fromSplit }) => {
                  if (fromSplit) return [];
                  const parts = splitInclAmountForExclDelta(
                    item.incl,
                    rate,
                    item.excl,
                    requiredExclAdjustment.toFixed(2),
                  );
                  if (!parts) return [];

                  const splitExcl = parts.reduce(
                    (sum, part) => sum.plus(new Decimal(part.excl)),
                    new Decimal(0),
                  );
                  const differenceAfter = sumExcl
                    .minus(new Decimal(item.excl || 0))
                    .plus(splitExcl)
                    .minus(groupExpectedExclBySum);
                  if (!differenceAfter.isZero()) return [];

                  return [{
                    subjectKey: key,
                    originalIncl: new Decimal(item.incl || 0).toFixed(2),
                    originalExcl: new Decimal(item.excl || 0).toFixed(2),
                    parts,
                    exclAdjustment: splitExcl.minus(new Decimal(item.excl || 0)).toFixed(2),
                    differenceBefore: groupDifference.toFixed(2),
                    differenceAfter: differenceAfter.toFixed(2),
                  }];
                })
                .sort((left, right) => {
                  const leftSmallPart = Math.min(...left.parts.map(part => part.incl));
                  const rightSmallPart = Math.min(...right.parts.map(part => part.incl));
                  return rightSmallPart - leftSmallPart || left.subjectKey.localeCompare(right.subjectKey);
                })
                .slice(0, 3);
          totalDifference = totalDifference.plus(groupDifference);
          errors.push({
            side,
            taxRate: rate,
            key: hasSplit ? `[汇总误差-公式C2]` : `[汇总误差-公式C1]`,
            expectedExcl: (hasSplit ? sumExpectedExcl : groupExpectedExclBySum).toFixed(2),
            actualExcl: sumExcl.toFixed(2),
            difference: groupDifference.toFixed(2),
            ...(splitSuggestions.length ? { splitSuggestions } : {}),
          });
        }
      }
    }
  };

  validateGroup(revenueData, 'income');
  validateGroup(costData, 'expense');

  return { errors, totalDifference: totalDifference.toFixed(2) };
}
