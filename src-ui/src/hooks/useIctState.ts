import { useState, useEffect, useCallback } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  buildDistributionFromModel,
  normalizeProjectYears,
  type CashflowModel,
  normalizeDistribution
} from "../lib/cashflowDistribution";

export { normalizeProjectYears };

export interface TaxItem { incl: number; tax: number; excl: number; }
export const defaultTaxItem = (tax = 6): TaxItem => ({ incl: 0, tax, excl: 0 });
export type SegmentValueMode = "ratio" | "amount";
export type SegmentFlowMode = "upfront" | "equal" | "custom";
export type SegmentSideScope = "project" | "it" | "ct" | "non_it_ct" | "mix";

export interface CashflowSegment {
  id: string;
  name: string;
  value: number;
  revenueValue: number;
  revenueTax: number;
  revenueScope: SegmentSideScope;
  costValue: number;
  costTax: number;
  costScope: SegmentSideScope;
  startYear: number;
  serviceYears: number;
  revenueMode: SegmentFlowMode;
  costMode: SegmentFlowMode;
  revenueAnnualValues: number[];
  costAnnualValues: number[];
}

export const clampCashflowYear = (value: number) => Math.max(1, Math.min(Number.isFinite(value) ? Math.trunc(value) : 1, 10));

export const getSegmentEffectiveDurationValue = (segment: CashflowSegment) => {
  const startIndex = clampCashflowYear(segment.startYear) - 1;
  return Math.max(1, Math.min(clampCashflowYear(segment.serviceYears), 10 - startIndex));
};

export const createCashflowSegment = (index: number): CashflowSegment => ({
  id: `${Date.now()}-${index}-${Math.random().toString(16).slice(2)}`,
  name: `板块${index}`,
  value: 0,
  revenueValue: 0,
  revenueTax: 6,
  revenueScope: "project",
  costValue: 0,
  costTax: 6,
  costScope: "project",
  startYear: 1,
  serviceYears: 1,
  revenueMode: "upfront",
  costMode: "upfront",
  revenueAnnualValues: [],
  costAnnualValues: [],
});

const roundMoney = (value: number) => Number((Number.isFinite(value) ? value : 0).toFixed(2));
const normalizeMoney = (value: number) => {
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric > 0 ? numeric : 0;
};

const buildSegmentFlowDistribution = (segment: CashflowSegment, mode: SegmentFlowMode, annualValues: number[] = []) => {
  const years = 10;
  const dist = Array(years).fill(0);
  const startIndex = clampCashflowYear(segment.startYear) - 1;
  const duration = Math.max(1, Math.min(clampCashflowYear(segment.serviceYears), years - startIndex));

  if (mode === "upfront") {
    dist[startIndex] = 1;
    return dist;
  }

  if (mode === "custom") {
    const values = Array.from({ length: duration }, (_, idx) => {
      const value = Number(annualValues[idx] || 0);
      return Number.isFinite(value) && value > 0 ? value : 0;
    });
    const sum = values.reduce((acc, value) => acc + value, 0);

    if (sum > 0) {
      values.forEach((value, idx) => {
        dist[startIndex + idx] = value / sum;
      });
      return dist;
    }
  }

  for (let i = startIndex; i < startIndex + duration; i++) {
    dist[i] = 1 / duration;
  }

  return dist;
};

export const buildSegmentDistribution = (segments: CashflowSegment[], side: "revenue" | "cost") => {
  const weighted = Array(10).fill(0);
  const validSegments = segments.filter(segment => Number(segment.value) > 0);
  const totalWeight = validSegments.reduce((sum, segment) => sum + Number(segment.value), 0);

  if (totalWeight <= 0) return buildDistributionFromModel("model_a", 1);

  validSegments.forEach(segment => {
    const flowMode = side === "revenue" ? segment.revenueMode : segment.costMode;
    const annualValues = side === "revenue" ? segment.revenueAnnualValues : segment.costAnnualValues;
    const segmentDist = buildSegmentFlowDistribution(segment, flowMode, annualValues);
    const weight = Number(segment.value) / totalWeight;
    segmentDist.forEach((ratio, index) => {
      weighted[index] += ratio * weight;
    });
  });

  return normalizeDistribution(weighted, 10);
};

const buildSegmentAnnualInclValues = (segment: CashflowSegment, side: "revenue" | "cost") => {
  const amount = normalizeMoney(side === "revenue" ? segment.revenueValue : segment.costValue);
  const mode = side === "revenue" ? segment.revenueMode : segment.costMode;
  const annualValues = side === "revenue" ? segment.revenueAnnualValues : segment.costAnnualValues;
  const values = Array(10).fill(0);
  const startIndex = clampCashflowYear(segment.startYear) - 1;
  const duration = Math.max(1, Math.min(clampCashflowYear(segment.serviceYears), 10 - startIndex));

  if (amount <= 0) return values;

  if (mode === "upfront") {
    values[startIndex] = amount;
    return values;
  }

  if (mode === "custom") {
    const customValues = Array.from({ length: duration }, (_, idx) => normalizeMoney(annualValues[idx] || 0));
    const customSum = customValues.reduce((sum, value) => sum + value, 0);

    if (customSum > 0) {
      customValues.forEach((value, idx) => {
        values[startIndex + idx] = value;
      });
      return values;
    }
  }

  const annualAmount = amount / duration;
  for (let i = startIndex; i < startIndex + duration; i++) {
    values[i] = annualAmount;
  }
  return values;
};

const addArrays = (target: number[], source: number[]) => {
  source.forEach((value, index) => {
    target[index] += value;
  });
};

const toExclCashflow = (inclValues: number[], taxRate: number) => {
  const rate = Number(taxRate);
  const divisor = 1 + (Number.isFinite(rate) ? rate : 0) / 100;
  return inclValues.map(value => value / divisor);
};

export const buildDirectCashflowFromSegments = (segments: CashflowSegment[]) => {
  const result = {
    rev: Array(10).fill(0),
    cost: Array(10).fill(0),
    itRev: Array(10).fill(0),
    itCost: Array(10).fill(0),
  };

  segments.forEach(segment => {
    const revenueExcl = toExclCashflow(buildSegmentAnnualInclValues(segment, "revenue"), segment.revenueTax);
    const costExcl = toExclCashflow(buildSegmentAnnualInclValues(segment, "cost"), segment.costTax);

    addArrays(result.rev, revenueExcl);
    addArrays(result.cost, costExcl);

    if (segment.revenueScope === "it") addArrays(result.itRev, revenueExcl);
    if (segment.costScope === "it") addArrays(result.itCost, costExcl);
  });

  return {
    rev: result.rev.map(roundMoney),
    cost: result.cost.map(roundMoney),
    itRev: result.itRev.map(roundMoney),
    itCost: result.itCost.map(roundMoney),
  };
};

export const distributionFromCashflow = (cashflow: number[]) => {
  const total = cashflow.reduce((sum, value) => sum + normalizeMoney(value), 0);
  if (total <= 0) return buildDistributionFromModel("model_a", 1);
  return normalizeDistribution(cashflow, 10);
};

export const cashflowPayloadValues = (cashflow: number[]) => cashflow.map(value => roundMoney(value).toFixed(2));
export const sumInclTaxItems = (items: TaxItem[]) => items.reduce((sum, item) => sum + normalizeMoney(item.incl), 0);

export const makeTaxItemFromIncl = (incl: number, tax: number): TaxItem => {
  const safeIncl = roundMoney(incl);
  const safeTax = Number.isFinite(Number(tax)) ? Number(tax) : 0;
  return {
    incl: safeIncl,
    tax: safeTax,
    excl: safeIncl === 0 ? 0 : roundMoney(safeIncl / (1 + safeTax / 100)),
  };
};

export function useIctState() {
  const [activeTab, setActiveTab] = useState<"basic" | "revenue" | "cost" | "cashflow" | "generate">("basic");

  // --- Basic State ---
  const [projName, setProjName] = useState("X项目");
  const [customerName, setCustomerName] = useState("X客户");
  const [propertyRights, setPropertyRights] = useState("客户");
  const [discountRate, setDiscountRate] = useState(0.055);
  const [projectYears, setProjectYears] = useState(1);
  const [cashflowModel, setCashflowModel] = useState<CashflowModel>("model_a");
  const [distRev, setDistRev] = useState<number[]>([1, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
  const [distCost, setDistCost] = useState<number[]>([1, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
  const [segmentValueMode, setSegmentValueMode] = useState<SegmentValueMode>("ratio");
  const [cashflowSegments, setCashflowSegments] = useState<CashflowSegment[]>([
    {
      id: "segment-1",
      name: "板块一",
      value: 50,
      revenueValue: 0,
      revenueTax: 6,
      revenueScope: "project",
      costValue: 0,
      costTax: 6,
      costScope: "project",
      startYear: 1,
      serviceYears: 3,
      revenueMode: "upfront",
      costMode: "upfront",
      revenueAnnualValues: [],
      costAnnualValues: [],
    },
    {
      id: "segment-2",
      name: "板块二",
      value: 50,
      revenueValue: 0,
      revenueTax: 6,
      revenueScope: "project",
      costValue: 0,
      costTax: 6,
      costScope: "project",
      startYear: 1,
      serviceYears: 2,
      revenueMode: "equal",
      costMode: "equal",
      revenueAnnualValues: [],
      costAnnualValues: [],
    },
  ]);
  const [projectBackground, setProjectBackground] = useState("在数字经济与制造业深度融合的国家战略推动下...");
  const [techItems, setTechItems] = useState<any[]>([
    { serviceName: '集成服务', serviceDesc: '集成实施', amount: 1, unit: '项' },
    { serviceName: '维保服务', serviceDesc: '硬件维保', amount: 1, unit: '项' }
  ]);
  const [inqVendors, setInqVendors] = useState<any[]>([
    { vendorName: '厂商A', amount: 0, taxRate: 6, remark: '最低' },
    { vendorName: '厂商B', amount: 0, taxRate: 6, remark: '' },
    { vendorName: '厂商C', amount: 0, taxRate: 6, remark: '' }
  ]);

  // --- Quick Calc State ---
  const [quickRevTotal, setQuickRevTotal] = useState<string>("");
  const [quickRevProduct, setQuickRevProduct] = useState<string>("");
  const [quickCostTotal, setQuickCostTotal] = useState<string>("");
  const [quickCostProduct, setQuickCostProduct] = useState<string>("");

  // --- Revenue State ---
  const [revIt, setRevIt] = useState({
    integration: defaultTaxItem(6), maintenance: defaultTaxItem(6),
    device_sales: defaultTaxItem(13), device_lease: defaultTaxItem(13),
    other: defaultTaxItem(6), cloud: defaultTaxItem(6),
  });
  const [revCt, setRevCt] = useState({ line: defaultTaxItem(9), product: defaultTaxItem(6) });
  const [revNonItCt, setRevNonItCt] = useState(defaultTaxItem(9));

  // --- Cost State ---
  const [costIt, setCostIt] = useState({
    device: defaultTaxItem(13), construction: defaultTaxItem(9),
    survey: defaultTaxItem(6), integration: defaultTaxItem(6),
    other: defaultTaxItem(6), maintenance: defaultTaxItem(6),
    running: defaultTaxItem(13), bidding: defaultTaxItem(6),
    design_eval: defaultTaxItem(6), audit: defaultTaxItem(6),
  });
  const [costCt, setCostCt] = useState({
    construction: defaultTaxItem(9), maintenance: defaultTaxItem(9),
    other: defaultTaxItem(6), bandwidth: defaultTaxItem(9), renewal: defaultTaxItem(9),
  });
  const [costMix, setCostMix] = useState({
    non_it_ct: defaultTaxItem(9), marketing: defaultTaxItem(6),
    channel: defaultTaxItem(6), other: defaultTaxItem(6),
  });

  // --- Templates & Ignore States ---
  const [templates, setTemplates] = useState<string[]>([]);
  const [selectedTemplate, setSelectedTemplate] = useState<string>("");

  const [reconciliationErrors, setReconciliationErrors] = useState<any[]>([]);
  const [showReconciliationModal, setShowReconciliationModal] = useState(false);
  const [currentTotalDifference, setCurrentTotalDifference] = useState("0");
  const [pendingTab, setPendingTab] = useState<{tab: string, template?: string} | null>(null);
  const [showConfirmIgnore, setShowConfirmIgnore] = useState(false);
  const [ignoredTailValue, setIgnoredTailValue] = useState<string | null>(null);
  const [ignoredDataHash, setIgnoredDataHash] = useState<string | null>(null);

  const loadTemplates = useCallback(async () => {
    try {
      const list: any = await invoke('get_available_templates', { moduleId: 'ict_lifecycle' });
      setTemplates(list);
    } catch (e) {
      console.error("加载 ICT 模板失败:", e);
    }
  }, []);

  useEffect(() => {
    loadTemplates();
  }, [loadTemplates]);

  // Sync distributions when cashflowModel or projectYears changes
  useEffect(() => {
    if (cashflowModel === 'model_d') {
      setDistRev(prev => buildDistributionFromModel(cashflowModel, projectYears, prev));
      setDistCost(prev => buildDistributionFromModel(cashflowModel, projectYears, prev));
      return;
    }

    if (cashflowModel === 'model_e') return;

    const nextDist = buildDistributionFromModel(cashflowModel, projectYears);
    setDistRev(nextDist);
    setDistCost(nextDist);
  }, [cashflowModel, projectYears]);

  const addCashflowSegment = () => {
    setCashflowSegments(prev => [...prev, createCashflowSegment(prev.length + 1)]);
  };

  const updateCashflowSegment = (id: string, key: keyof CashflowSegment, value: string | number) => {
    setCashflowSegments(prev => prev.map(segment => {
      if (segment.id !== id) return segment;
      const numericFields: Array<keyof CashflowSegment> = ["value", "revenueValue", "revenueTax", "costValue", "costTax", "startYear", "serviceYears"];
      const nextValue = numericFields.includes(key) ? Number(value) : value;
      return { ...segment, [key]: nextValue };
    }));
  };

  const updateCashflowSegmentAnnualValue = (id: string, side: "revenue" | "cost", index: number, value: number) => {
    setCashflowSegments(prev => prev.map(segment => {
      if (segment.id !== id) return segment;
      const field = side === "revenue" ? "revenueAnnualValues" : "costAnnualValues";
      const nextValues = [...(segment[field] || [])];
      nextValues[index] = Number.isFinite(value) && value > 0 ? value : 0;
      return { ...segment, [field]: nextValues };
    }));
  };

  const removeCashflowSegment = (id: string) => {
    setCashflowSegments(prev => prev.length <= 1 ? prev : prev.filter(segment => segment.id !== id));
  };

  const updateTaxItem = (groupId: string, key: string, field: "incl" | "tax" | "excl", val: number) => {
    if (ignoredDataHash !== null) {
      setIgnoredDataHash(null);
      setIgnoredTailValue(null);
    }

    const processItem = (groupState: any, setGroupState: any, targetKey: string) => {
      const item = { ...groupState[targetKey], [field]: isNaN(val) ? 0 : val };
      if (field === 'incl' || field === 'tax') {
        item.excl = item.incl === 0 ? 0 : Number((item.incl / (1 + item.tax / 100)).toFixed(2));
      } else if (field === 'excl') {
        item.incl = item.excl === 0 ? 0 : Number((item.excl * (1 + item.tax / 100)).toFixed(2));
      }
      setGroupState({ ...groupState, [targetKey]: item });
    };

    if (groupId === 'revIt') processItem(revIt, setRevIt, key);
    else if (groupId === 'revCt') {
      processItem(revCt, setRevCt, key);
      if (key === 'product') processItem(costCt, setCostCt, 'other');
      if (key === 'line') processItem(costCt, setCostCt, 'bandwidth');
    }
    else if (groupId === 'revNonItCt') {
      const item = { ...revNonItCt, [field]: isNaN(val) ? 0 : val };
      if (field === 'incl' || field === 'tax') {
        item.excl = item.incl === 0 ? 0 : Number((item.incl / (1 + item.tax / 100)).toFixed(2));
      } else if (field === 'excl') {
        item.incl = item.excl === 0 ? 0 : Number((item.excl * (1 + item.tax / 100)).toFixed(2));
      }
      setRevNonItCt(item);
    }
    else if (groupId === 'costIt') processItem(costIt, setCostIt, key);
    else if (groupId === 'costCt') processItem(costCt, setCostCt, key);
    else if (groupId === 'costMix') processItem(costMix, setCostMix, key);
  };

  return {
    activeTab, setActiveTab,
    projName, setProjName,
    customerName, setCustomerName,
    propertyRights, setPropertyRights,
    discountRate, setDiscountRate,
    projectYears, setProjectYears,
    cashflowModel, setCashflowModel,
    distRev, setDistRev,
    distCost, setDistCost,
    segmentValueMode, setSegmentValueMode,
    cashflowSegments, setCashflowSegments,
    projectBackground, setProjectBackground,
    techItems, setTechItems,
    inqVendors, setInqVendors,
    quickRevTotal, setQuickRevTotal,
    quickRevProduct, setQuickRevProduct,
    quickCostTotal, setQuickCostTotal,
    quickCostProduct, setQuickCostProduct,
    revIt, setRevIt,
    revCt, setRevCt,
    revNonItCt, setRevNonItCt,
    costIt, setCostIt,
    costCt, setCostCt,
    costMix, setCostMix,
    templates, setTemplates,
    selectedTemplate, setSelectedTemplate,
    reconciliationErrors, setReconciliationErrors,
    showReconciliationModal, setShowReconciliationModal,
    currentTotalDifference, setCurrentTotalDifference,
    pendingTab, setPendingTab,
    showConfirmIgnore, setShowConfirmIgnore,
    ignoredTailValue, setIgnoredTailValue,
    ignoredDataHash, setIgnoredDataHash,
    loadTemplates,
    addCashflowSegment,
    updateCashflowSegment,
    updateCashflowSegmentAnnualValue,
    removeCashflowSegment,
    updateTaxItem,
  };
}
