import Decimal from "decimal.js";

/**
 * 财务口径（方案一）下的税额换算工具。
 *
 * 口径约定：财务/业务系统以「不含税金额（保留两位小数）」为唯一锚点，
 * 含税金额一律由不含税反推：含税 = round(不含税 × (1 + 税率), 2)。
 * 因此本模块所有换算都用 decimal.js 做十进制半进位（四舍五入），
 * 避免 979.25 × 1.06 = 1038.00499…（IEEE 浮点）被错误舍入成 1038.00
 * 而财务系统显示 1038.01 的口径分歧。
 *
 * 注意 round(x ÷ (1+r)) 与 round(x × (1+r)) 不是互逆运算：
 * 含税 1038（6%）→ 不含税 979.25 → 反推含税 1038.01。
 * 这类"不可精确表示"的含税价通过 normalizeTaxPairFromIncl 归一到
 * 财务口径的不动点（round(excl×(1+r)) 再除回去必得原 excl）。
 */

const toSafeNumber = (value: unknown): number => {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : 0;
};

const taxDivisor = (taxRatePercent: unknown): Decimal =>
  new Decimal(100).plus(toSafeNumber(taxRatePercent)).div(100);

const roundMoneyDecimal = (value: Decimal): number =>
  value.toDecimalPlaces(2, Decimal.ROUND_HALF_UP).toNumber();

/** 金额四舍五入到分（十进制精确，非浮点 toFixed）。 */
export const roundMoneyHalfUp = (value: unknown): number =>
  roundMoneyDecimal(new Decimal(toSafeNumber(value)));

/** 不含税 = round(含税 ÷ (1 + 税率), 2)。录入含税价时派生不含税锚点用。 */
export const exclFromIncl = (incl: unknown, taxRatePercent: unknown): number => {
  const inclValue = toSafeNumber(incl);
  if (inclValue === 0) return 0;
  return roundMoneyDecimal(new Decimal(inclValue).div(taxDivisor(taxRatePercent)));
};

/** 含税 = round(不含税 × (1 + 税率), 2)。与财务系统展示口径逐分一致。 */
export const inclFromExcl = (excl: unknown, taxRatePercent: unknown): number => {
  const exclValue = toSafeNumber(excl);
  if (exclValue === 0) return 0;
  return roundMoneyDecimal(new Decimal(exclValue).mul(taxDivisor(taxRatePercent)));
};

export type NormalizedTaxPair = {
  /** 财务口径含税额（由不含税反推，展示与出文档一律用它）。 */
  incl: number;
  /** 不含税锚点金额。 */
  excl: number;
  /** 录入的原始含税额（四舍五入到分后）。 */
  enteredIncl: number;
  /** true 表示录入含税价在财务口径下不可精确表示，已被调整。 */
  adjusted: boolean;
};

/**
 * 以录入的含税价为起点归一到财务口径：
 * excl = round(incl ÷ (1+r))，incl' = round(excl × (1+r))。
 * incl' 是反推运算的不动点（再除回去仍得同一个 excl），
 * 归一后的税额对儿与财务系统的展示逐分一致。
 */
export const normalizeTaxPairFromIncl = (
  incl: unknown,
  taxRatePercent: unknown,
): NormalizedTaxPair => {
  const enteredIncl = roundMoneyHalfUp(incl);
  const excl = exclFromIncl(enteredIncl, taxRatePercent);
  const systemIncl = inclFromExcl(excl, taxRatePercent);
  return {
    incl: systemIncl,
    excl,
    enteredIncl,
    adjusted: systemIncl !== enteredIncl,
  };
};

/** 整数分口径的不含税还原：round(含税分 × 100 ÷ (100 + 税率))。 */
export const exclCentsFromInclCents = (
  inclCents: number,
  taxRatePercent: unknown,
): number => {
  if (!Number.isFinite(inclCents) || inclCents === 0) return 0;
  return new Decimal(inclCents)
    .div(taxDivisor(taxRatePercent))
    .toDecimalPlaces(0, Decimal.ROUND_HALF_UP)
    .toNumber();
};
