import Decimal from 'decimal.js';

export interface TaxItem {
  incl: string | number;
  excl: string | number;
  tax: string | number; // Note: using "tax" key to match IctLifecycle.tsx state
}

export interface ValidationReport {
  side: 'income' | 'expense';
  taxRate: number;
  key: string;
  expectedExcl: string;
  actualExcl: string;
  difference: string;
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
    const taxGroups = new Map<number, { key: string, item: TaxItem }[]>();

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
        taxGroups.get(rate)!.push({ key: `${category}.${itemKey}`, item });
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
            difference: difference.toFixed(2)
          });
        }
      }

      // C: Aggregate-level Strict Validation
      // C1: 倒算核验 = sum(excl) - round(sum(incl) / (1 + tax_rate), 2)
      const groupExpectedExclBySum = sumIncl.div(taxDivisor).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);
      const groupDifferenceC1 = sumExcl.minus(groupExpectedExclBySum);

      // C2: 累加核验 = sum(expectedExcl) === sum(excl)
      const groupDifferenceC2 = sumExpectedExcl.minus(sumExcl);

      if (!groupDifferenceC1.isZero() || !groupDifferenceC2.isZero()) {
        const hasItemErrors = errors.some(e => e.side === side && e.taxRate === rate);
        if (!hasItemErrors) {
          // If no specific item error was caught, but the sum is wrong, it means there is a rounding mismatch at the group level
          // Force user to adjust.
          totalDifference = totalDifference.plus(groupDifferenceC1);
          errors.push({
            side,
            taxRate: rate,
            key: `[汇总误差-公式C1]`,
            expectedExcl: groupExpectedExclBySum.toFixed(2),
            actualExcl: sumExcl.toFixed(2),
            difference: groupDifferenceC1.toFixed(2)
          });
        }
      }
    }
  };

  validateGroup(revenueData, 'income');
  validateGroup(costData, 'expense');

  return { errors, totalDifference: totalDifference.toFixed(2) };
}
