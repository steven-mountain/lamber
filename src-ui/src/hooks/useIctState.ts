import { useState, useEffect, useCallback } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  buildDistributionFromModel,
  normalizeProjectYears,
  type CashflowModel,
  normalizeDistribution
} from "../lib/cashflowDistribution";
import { ICT_SUBJECT_GROUPS, normalizeCustomSubjectName, type IctSubjectGroupId } from "../lib/ictSubjectCatalog";
import {
  createDefaultBalanceAllocationState,
  type BalanceAllocationSide,
  type BalanceAllocationRule,
  type BalanceAllocationState,
} from "../lib/ictBalanceAllocation";
import {
  normalizeCashflowCalculationSource,
  normalizeSubjectFundingPlans,
  SUBJECT_FUNDING_PLAN_MIGRATION_VERSION,
  initializeMissingSubjectFundingPlans,
  removeSubjectFundingPlan,
  syncSubjectFundingPlansToAmounts,
  upsertSubjectFundingPlan as upsertSubjectFundingPlanRecord,
  type CashflowCalculationSource,
  type SubjectFundingPlan,
  type SubjectFundingPlanLastChangeReason,
  type SubjectFundingPlans,
  type SubjectFundingSubjectRef,
} from "../lib/ictSubjectFundingPlan";

export { normalizeProjectYears };

export interface TaxItem { incl: number; tax: number; excl: number; customSubjectName?: string; billingSubjectName?: string; }

export type TaxItemInclUpdate = { groupId: string; key: string; incl: number; reason?: SubjectFundingPlanLastChangeReason };
export const defaultTaxItem = (tax = 6): TaxItem => ({ incl: 0, tax, excl: 0 });
const defaultSubjectTaxItem = (groupId: IctSubjectGroupId, key: string) => {
  const subject = ICT_SUBJECT_GROUPS[groupId]?.find(item => item.key === key);
  return defaultTaxItem(subject?.defaultTaxRate ?? 6);
};
const createDefaultRevIt = () => ({
  integration: defaultSubjectTaxItem("revIt", "integration"),
  maintenance: defaultSubjectTaxItem("revIt", "maintenance"),
  device_sales: defaultSubjectTaxItem("revIt", "device_sales"),
  device_lease: defaultSubjectTaxItem("revIt", "device_lease"),
  other: defaultSubjectTaxItem("revIt", "other"),
  cloud: defaultSubjectTaxItem("revIt", "cloud"),
});
const createDefaultRevCt = () => ({
  line: defaultSubjectTaxItem("revCt", "line"),
  product: defaultSubjectTaxItem("revCt", "product"),
});
const createDefaultRevNonItCt = () => defaultSubjectTaxItem("revNonItCt", "item");
const createDefaultCostIt = () => ({
  device: defaultSubjectTaxItem("costIt", "device"),
  construction: defaultSubjectTaxItem("costIt", "construction"),
  survey: defaultSubjectTaxItem("costIt", "survey"),
  integration: defaultSubjectTaxItem("costIt", "integration"),
  other: defaultSubjectTaxItem("costIt", "other"),
  maintenance: defaultSubjectTaxItem("costIt", "maintenance"),
  running: defaultSubjectTaxItem("costIt", "running"),
  bidding: defaultSubjectTaxItem("costIt", "bidding"),
  design_eval: defaultSubjectTaxItem("costIt", "design_eval"),
  audit: defaultSubjectTaxItem("costIt", "audit"),
});
const createDefaultCostCt = () => ({
  construction: defaultSubjectTaxItem("costCt", "construction"),
  maintenance: defaultSubjectTaxItem("costCt", "maintenance"),
  other: defaultSubjectTaxItem("costCt", "other"),
  bandwidth: defaultSubjectTaxItem("costCt", "bandwidth"),
  renewal: defaultSubjectTaxItem("costCt", "renewal"),
});
const createDefaultCostMix = () => ({
  non_it_ct: defaultSubjectTaxItem("costMix", "non_it_ct"),
  marketing: defaultSubjectTaxItem("costMix", "marketing"),
  channel: defaultSubjectTaxItem("costMix", "channel"),
  other: defaultSubjectTaxItem("costMix", "other"),
});
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

const createDefaultCashflowSegments = (): CashflowSegment[] => [
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
];

const createDefaultTechItems = () => [
  { serviceName: "集成服务", serviceDesc: "集成实施", amount: 1, unit: "项" },
  { serviceName: "维保服务", serviceDesc: "硬件维保", amount: 1, unit: "项" },
];

const createDefaultInquiryVendors = () => [
  { vendorName: "厂商A", amount: 0, taxRate: 6, remark: "最低" },
  { vendorName: "厂商B", amount: 0, taxRate: 6, remark: "" },
  { vendorName: "厂商C", amount: 0, taxRate: 6, remark: "" },
];

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
  const [cashflowSegments, setCashflowSegments] = useState<CashflowSegment[]>(createDefaultCashflowSegments);
  const [projectBackground, setProjectBackground] = useState("在数字经济与制造业深度融合的国家战略推动下...");
  const [techItems, setTechItems] = useState<any[]>(createDefaultTechItems);
  const [inqVendors, setInqVendors] = useState<any[]>(createDefaultInquiryVendors);

  const [balanceAllocation, setBalanceAllocation] = useState<BalanceAllocationState>(
    createDefaultBalanceAllocationState(),
  );
  const [subjectFundingPlans, setSubjectFundingPlansState] = useState<SubjectFundingPlans>({});
  const [cashflowCalculationSource, setCashflowCalculationSourceState] =
    useState<CashflowCalculationSource>("subject_funding_plans");
  const [subjectFundingPlanMigrationVersion, setSubjectFundingPlanMigrationVersion] =
    useState<number | undefined>(SUBJECT_FUNDING_PLAN_MIGRATION_VERSION);

  // --- Revenue State ---
  const [revIt, setRevIt] = useState(createDefaultRevIt);
  const [revCt, setRevCt] = useState(createDefaultRevCt);
  const [revNonItCt, setRevNonItCt] = useState(createDefaultRevNonItCt);

  // --- Cost State ---
  const [costIt, setCostIt] = useState(createDefaultCostIt);
  const [costCt, setCostCt] = useState(createDefaultCostCt);
  const [costMix, setCostMix] = useState(createDefaultCostMix);

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

  const setSubjectFundingPlans = (plans: SubjectFundingPlans | unknown) => {
    setSubjectFundingPlansState(normalizeSubjectFundingPlans(plans));
  };

  const setCashflowCalculationSource = (source: CashflowCalculationSource | unknown) => {
    setCashflowCalculationSourceState(normalizeCashflowCalculationSource(source));
  };

  const upsertSubjectFundingPlan = (plan: SubjectFundingPlan) => {
    setSubjectFundingPlansState(prev => upsertSubjectFundingPlanRecord(prev, plan));
  };

  const removeSubjectFundingPlanForRef = (subjectRef: SubjectFundingSubjectRef) => {
    setSubjectFundingPlansState(prev => removeSubjectFundingPlan(prev, subjectRef));
  };

  const collectPositiveFundingSubjects = (sources?: {
    revItState?: typeof revIt;
    revCtState?: typeof revCt;
    revNonItCtState?: typeof revNonItCt;
    costItState?: typeof costIt;
    costCtState?: typeof costCt;
    costMixState?: typeof costMix;
  }) => {
    const result: Array<{ subjectRef: SubjectFundingSubjectRef; amountIncl: number }> = [];
    const addRecordItems = (
      side: SubjectFundingSubjectRef["side"],
      groupId: SubjectFundingSubjectRef["groupId"],
      items: Record<string, TaxItem>,
    ) => {
      Object.entries(items).forEach(([itemKey, item]) => {
        if (item.incl > 0) {
          result.push({ subjectRef: { side, groupId, key: itemKey }, amountIncl: item.incl });
        }
      });
    };

    addRecordItems("revenue", "revIt", sources?.revItState ?? revIt);
    addRecordItems("revenue", "revCt", sources?.revCtState ?? revCt);
    const nonItCt = sources?.revNonItCtState ?? revNonItCt;
    if (nonItCt.incl > 0) {
      result.push({ subjectRef: { side: "revenue", groupId: "revNonItCt", key: "item" }, amountIncl: nonItCt.incl });
    }
    addRecordItems("cost", "costIt", sources?.costItState ?? costIt);
    addRecordItems("cost", "costCt", sources?.costCtState ?? costCt);
    addRecordItems("cost", "costMix", sources?.costMixState ?? costMix);
    return result;
  };

  const syncFundingPlansAfterAmountChange = (
    updates: Array<{ subjectRef: SubjectFundingSubjectRef; newAmountIncl: number; reason?: SubjectFundingPlanLastChangeReason }>,
    positiveSubjects: Array<{ subjectRef: SubjectFundingSubjectRef; amountIncl: number }>,
  ) => {
    if (updates.length === 0) return;
    const zeroedSubjectKeys = new Set(
      updates
        .filter(update => update.newAmountIncl <= 0)
        .map(update => `${update.subjectRef.side}:${update.subjectRef.groupId}:${update.subjectRef.key}`),
    );
    const activePositiveSubjects = positiveSubjects.filter(subject =>
      !zeroedSubjectKeys.has(`${subject.subjectRef.side}:${subject.subjectRef.groupId}:${subject.subjectRef.key}`),
    );
    setSubjectFundingPlansState(prev => {
      const syncedPlans = syncSubjectFundingPlansToAmounts(prev, updates);
      return initializeMissingSubjectFundingPlans(syncedPlans, activePositiveSubjects);
    });
    setCashflowCalculationSourceState("subject_funding_plans");
  };

  const clearFinancialSubjects = () => {
    setRevIt(createDefaultRevIt());
    setRevCt(createDefaultRevCt());
    setRevNonItCt(createDefaultRevNonItCt());
    setCostIt(createDefaultCostIt());
    setCostCt(createDefaultCostCt());
    setCostMix(createDefaultCostMix());
    setSubjectFundingPlansState({});
    setBalanceAllocation(createDefaultBalanceAllocationState());
    setIgnoredTailValue(null);
    setIgnoredDataHash(null);
    setReconciliationErrors([]);
    setShowReconciliationModal(false);
    setShowConfirmIgnore(false);
    setPendingTab(null);
    setCurrentTotalDifference("0");
    setCashflowCalculationSourceState("subject_funding_plans");
    setSubjectFundingPlanMigrationVersion(SUBJECT_FUNDING_PLAN_MIGRATION_VERSION);
    setCashflowSegments(prev => prev.map(segment => ({
      ...segment,
      revenueValue: 0,
      costValue: 0,
      revenueAnnualValues: [],
      costAnnualValues: [],
    })));
  };

  const resetProjectState = () => {
    setActiveTab("basic");
    setProjName("X项目");
    setCustomerName("X客户");
    setPropertyRights("客户");
    setDiscountRate(0.055);
    setProjectYears(1);
    setCashflowModel("model_a");
    setDistRev([1, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
    setDistCost([1, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
    setSegmentValueMode("ratio");
    setCashflowSegments(createDefaultCashflowSegments());
    setProjectBackground("在数字经济与制造业深度融合的国家战略推动下...");
    setTechItems(createDefaultTechItems());
    setInqVendors(createDefaultInquiryVendors());
    setSelectedTemplate("");
    clearFinancialSubjects();
  };

  const updateTaxItemsInclBatch = (updates: TaxItemInclUpdate[]) => {
    if (updates.length === 0) return;
    if (ignoredDataHash !== null) {
      setIgnoredDataHash(null);
      setIgnoredTailValue(null);
    }

    let nextRevIt = revIt;
    let nextRevCt = revCt;
    let nextRevNonItCt = revNonItCt;
    let nextCostIt = costIt;
    let nextCostCt = costCt;
    let nextCostMix = costMix;
    const changed = {
      revIt: false,
      revCt: false,
      revNonItCt: false,
      costIt: false,
      costCt: false,
      costMix: false,
    };

    const itemFromIncl = (current: TaxItem | undefined, incl: number): TaxItem => {
      const safeIncl = isNaN(incl) ? 0 : roundMoney(incl);
      const tax = Number(current?.tax ?? 0);
      return {
        ...(current || defaultTaxItem(tax)),
        incl: safeIncl,
        tax,
        excl: safeIncl === 0 ? 0 : roundMoney(safeIncl / (1 + tax / 100)),
      };
    };

    const setRecordItem = <T extends Record<string, TaxItem>>(
      group: T,
      key: string,
      incl: number,
    ): T => ({
      ...group,
      [key]: itemFromIncl(group[key], incl),
    } as T);

    updates.forEach(update => {
      if (update.groupId === "revIt") {
        nextRevIt = setRecordItem(nextRevIt, update.key, update.incl);
        changed.revIt = true;
      } else if (update.groupId === "revCt") {
        nextRevCt = setRecordItem(nextRevCt, update.key, update.incl);
        changed.revCt = true;
        if (update.key === "product") {
          nextCostCt = setRecordItem(nextCostCt, "other", update.incl);
          changed.costCt = true;
        }
        if (update.key === "line") {
          nextCostCt = setRecordItem(nextCostCt, "bandwidth", update.incl);
          changed.costCt = true;
        }
      } else if (update.groupId === "revNonItCt") {
        nextRevNonItCt = itemFromIncl(nextRevNonItCt, update.incl);
        changed.revNonItCt = true;
      } else if (update.groupId === "costIt") {
        nextCostIt = setRecordItem(nextCostIt, update.key, update.incl);
        changed.costIt = true;
      } else if (update.groupId === "costCt") {
        nextCostCt = setRecordItem(nextCostCt, update.key, update.incl);
        changed.costCt = true;
      } else if (update.groupId === "costMix") {
        nextCostMix = setRecordItem(nextCostMix, update.key, update.incl);
        changed.costMix = true;
      }
    });

    if (changed.revIt) setRevIt(nextRevIt);
    if (changed.revCt) setRevCt(nextRevCt);
    if (changed.revNonItCt) setRevNonItCt(nextRevNonItCt);
    if (changed.costIt) setCostIt(nextCostIt);
    if (changed.costCt) setCostCt(nextCostCt);
    if (changed.costMix) setCostMix(nextCostMix);

    const syncUpdates: Array<{ subjectRef: SubjectFundingSubjectRef; newAmountIncl: number; reason?: SubjectFundingPlanLastChangeReason }> = [];
    updates.forEach(update => {
      const side = (update.groupId === "revIt" || update.groupId === "revCt" || update.groupId === "revNonItCt")
        ? "revenue" as const : "cost" as const;
      syncUpdates.push({
        subjectRef: { side, groupId: update.groupId as SubjectFundingSubjectRef["groupId"], key: update.key },
        newAmountIncl: update.incl,
        reason: update.reason,
      });
      // CT linkage: revCt.product → costCt.other, revCt.line → costCt.bandwidth
      if (update.groupId === "revCt" && update.key === "product") {
        syncUpdates.push({ subjectRef: { side: "cost", groupId: "costCt", key: "other" }, newAmountIncl: update.incl, reason: "ct_linkage_sync" });
      }
      if (update.groupId === "revCt" && update.key === "line") {
        syncUpdates.push({ subjectRef: { side: "cost", groupId: "costCt", key: "bandwidth" }, newAmountIncl: update.incl, reason: "ct_linkage_sync" });
      }
    });
    syncFundingPlansAfterAmountChange(
      syncUpdates,
      collectPositiveFundingSubjects({
        revItState: nextRevIt,
        revCtState: nextRevCt,
        revNonItCtState: nextRevNonItCt,
        costItState: nextCostIt,
        costCtState: nextCostCt,
        costMixState: nextCostMix,
      }),
    );
  };

  const updateTaxItem = (groupId: string, key: string, field: "incl" | "tax" | "excl", val: number, reason?: SubjectFundingPlanLastChangeReason) => {
    if (ignoredDataHash !== null) {
      setIgnoredDataHash(null);
      setIgnoredTailValue(null);
    }

    // Collect effective incl amounts for funding plan sync.
    // processItem returns the resolved incl value so we can sync plans afterwards.
    const processItem = (groupState: any, setGroupState: any, targetKey: string): number => {
      const item = { ...groupState[targetKey], [field]: isNaN(val) ? 0 : val };
      if (field === 'incl' || field === 'tax') {
        item.excl = item.incl === 0 ? 0 : Number((item.incl / (1 + item.tax / 100)).toFixed(2));
      } else if (field === 'excl') {
        item.incl = item.excl === 0 ? 0 : Number((item.excl * (1 + item.tax / 100)).toFixed(2));
      }
      setGroupState({ ...groupState, [targetKey]: item });
      return Number(item.incl) || 0;
    };

    // Track incl amounts for subjects that need plan sync
    const syncUpdates: Array<{ subjectRef: SubjectFundingSubjectRef; newAmountIncl: number; reason?: SubjectFundingPlanLastChangeReason }> = [];
    const needsSync = field !== "tax";

    const sideForGroup = (gid: string) =>
      (gid === "revIt" || gid === "revCt" || gid === "revNonItCt") ? "revenue" as const : "cost" as const;

    const trackSync = (gid: string, k: string, effectiveIncl: number, overrideReason?: SubjectFundingPlanLastChangeReason) => {
      if (!needsSync) return;
      syncUpdates.push({
        subjectRef: { side: sideForGroup(gid), groupId: gid as SubjectFundingSubjectRef["groupId"], key: k },
        newAmountIncl: effectiveIncl,
        reason: overrideReason || reason,
      });
    };

    if (groupId === 'revIt') {
      trackSync(groupId, key, processItem(revIt, setRevIt, key));
    } else if (groupId === 'revCt') {
      const effectiveIncl = processItem(revCt, setRevCt, key);
      trackSync(groupId, key, effectiveIncl);
      if (key === 'product') {
        processItem(costCt, setCostCt, 'other');
        trackSync('costCt', 'other', effectiveIncl, "ct_linkage_sync");
      }
      if (key === 'line') {
        processItem(costCt, setCostCt, 'bandwidth');
        trackSync('costCt', 'bandwidth', effectiveIncl, "ct_linkage_sync");
      }
    } else if (groupId === 'revNonItCt') {
      const item = { ...revNonItCt, [field]: isNaN(val) ? 0 : val };
      if (field === 'incl' || field === 'tax') {
        item.excl = item.incl === 0 ? 0 : Number((item.incl / (1 + item.tax / 100)).toFixed(2));
      } else if (field === 'excl') {
        item.incl = item.excl === 0 ? 0 : Number((item.excl * (1 + item.tax / 100)).toFixed(2));
      }
      setRevNonItCt(item);
      trackSync(groupId, key, Number(item.incl) || 0);
    } else if (groupId === 'costIt') {
      trackSync(groupId, key, processItem(costIt, setCostIt, key));
    } else if (groupId === 'costCt') {
      trackSync(groupId, key, processItem(costCt, setCostCt, key));
    } else if (groupId === 'costMix') {
      trackSync(groupId, key, processItem(costMix, setCostMix, key));
    }

    // Apply funding plan sync in the same React batch
    syncFundingPlansAfterAmountChange(syncUpdates, collectPositiveFundingSubjects());
  };

  const updateTaxItemTextField = (
    groupId: string,
    key: string,
    field: "customSubjectName" | "billingSubjectName",
    value: string,
  ) => {
    const normalizedValue = normalizeCustomSubjectName(value);
    const nextItem = (item: TaxItem) => ({
      ...item,
      [field]: normalizedValue,
    });

	    const updateGroupItem = (groupState: Record<string, TaxItem>, setGroupState: (value: any) => void, targetKey: string) => {
	      setGroupState({
	        ...groupState,
	        [targetKey]: nextItem(
	          groupState[targetKey] || defaultSubjectTaxItem(groupId as IctSubjectGroupId, targetKey),
	        ),
	      });
	    };

    if (groupId === 'revIt') updateGroupItem(revIt, setRevIt, key);
    else if (groupId === 'revCt') {
      updateGroupItem(revCt, setRevCt, key);
      if (key === 'product') updateGroupItem(costCt, setCostCt, 'other');
      if (key === 'line') updateGroupItem(costCt, setCostCt, 'bandwidth');
    }
    else if (groupId === 'revNonItCt') setRevNonItCt(nextItem(revNonItCt));
    else if (groupId === 'costIt') updateGroupItem(costIt, setCostIt, key);
    else if (groupId === 'costCt') {
      updateGroupItem(costCt, setCostCt, key);
      if (key === 'other') updateGroupItem(revCt, setRevCt, 'product');
      if (key === 'bandwidth') updateGroupItem(revCt, setRevCt, 'line');
    }
    else if (groupId === 'costMix') updateGroupItem(costMix, setCostMix, key);
  };

  const updateTaxItemCustomSubjectName = (groupId: string, key: string, value: string) => {
    updateTaxItemTextField(groupId, key, "customSubjectName", value);
  };

  const updateTaxItemBillingSubjectName = (groupId: string, key: string, value: string) => {
    updateTaxItemTextField(groupId, key, "billingSubjectName", value);
  };

  const updateBalanceRule = (
    side: BalanceAllocationSide,
    patch: Partial<BalanceAllocationRule>,
  ) => {
    if (patch.balancingSubject !== undefined) {
      const oldSubject = balanceAllocation[side].balancingSubject;
      const newSubject = patch.balancingSubject;
      if (oldSubject && newSubject && (oldSubject.groupId !== newSubject.groupId || oldSubject.key !== newSubject.key)) {
        updateTaxItem(oldSubject.groupId, oldSubject.key, "incl", 0);
      }
    }

    setBalanceAllocation(prev => ({
      ...prev,
      [side]: {
        ...prev[side],
        ...patch,
      },
    }));
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
    balanceAllocation, setBalanceAllocation, updateBalanceRule,
    cashflowCalculationSource, setCashflowCalculationSource,
    subjectFundingPlanMigrationVersion, setSubjectFundingPlanMigrationVersion,
    subjectFundingPlans, setSubjectFundingPlans, upsertSubjectFundingPlan, removeSubjectFundingPlanForRef,
    clearFinancialSubjects, resetProjectState,
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
    updateTaxItemsInclBatch,
    updateTaxItemCustomSubjectName,
    updateTaxItemBillingSubjectName,
  };
}
