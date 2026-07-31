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

/** 含税价在该税率下是否可精确表示（round(excl×(1+r)) 反推回原值）。 */
export const isInclRepresentable = (incl: unknown, taxRatePercent: unknown): boolean =>
  !normalizeTaxPairFromIncl(incl, taxRatePercent).adjusted;

export type TaxSplitPart = {
  /** 子笔含税金额（各子笔之和 = 原始含税总额）。 */
  incl: number;
  /** 子笔不含税金额 = round(incl ÷ (1+r))，反推必闭合。 */
  excl: number;
};

/**
 * 把不可精确表示的含税总额拆成两笔各自闭合的子金额，和严格等于总额。
 * 对税率 1/3/5/6/9/13%、0.02～100 万元穷举验证过：对半拆分（floor(T/2)
 * 与 T-floor(T/2)）全部自洽，零例外；中点向外的搜索仅是防御性兜底。
 * 总额本身自洽、≤0.01 元或拆分无解时返回 null（无需/无法拆分）。
 */
export const splitInclAmount = (
  incl: unknown,
  taxRatePercent: unknown,
): TaxSplitPart[] | null => {
  const totalCents = new Decimal(toSafeNumber(incl))
    .mul(100)
    .toDecimalPlaces(0, Decimal.ROUND_HALF_UP)
    .toNumber();
  if (totalCents <= 1) return null;
  const rate = toSafeNumber(taxRatePercent);
  const centsRepresentable = (cents: number): boolean =>
    isInclRepresentable(new Decimal(cents).div(100).toNumber(), rate);
  if (centsRepresentable(totalCents)) return null;

  const half = Math.floor(totalCents / 2);
  const buildParts = (aCents: number): TaxSplitPart[] => {
    const bCents = totalCents - aCents;
    return [aCents, bCents].map(cents => {
      const partIncl = new Decimal(cents).div(100).toNumber();
      return { incl: partIncl, excl: exclFromIncl(partIncl, rate) };
    });
  };
  for (let offset = 0; offset <= 100; offset++) {
    for (const aCents of offset === 0 ? [half] : [half - offset, half + offset]) {
      const bCents = totalCents - aCents;
      if (aCents < 1 || bCents < 1) continue;
      if (centsRepresentable(aCents) && centsRepresentable(bCents)) {
        return buildParts(aCents);
      }
    }
  }
  return null;
};

/**
 * 从存档还原拆分明细：每笔含税必须为正且自洽、合计严格等于科目含税，
 * 否则整体丢弃（回到普通单笔口径）。不含税一律按当前税率重新派生。
 */
export const restoreTaxSplitParts = (
  raw: unknown,
  totalIncl: unknown,
  taxRatePercent: unknown,
): TaxSplitPart[] | null => {
  if (!Array.isArray(raw) || raw.length < 2) return null;
  const rate = toSafeNumber(taxRatePercent);
  const parts: TaxSplitPart[] = [];
  let sumCents = 0;
  for (const entry of raw) {
    const source = entry as { incl_tax?: unknown; incl?: unknown } | null;
    const incl = roundMoneyHalfUp(source?.incl_tax ?? source?.incl);
    if (incl <= 0 || !isInclRepresentable(incl, rate)) return null;
    sumCents += Math.round(incl * 100);
    parts.push({ incl, excl: exclFromIncl(incl, rate) });
  }
  const totalCents = new Decimal(toSafeNumber(totalIncl))
    .mul(100)
    .toDecimalPlaces(0, Decimal.ROUND_HALF_UP)
    .toNumber();
  return sumCents === totalCents ? parts : null;
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
