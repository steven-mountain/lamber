import { useState, useEffect, useRef, useCallback } from "react";
import { invoke } from "@tauri-apps/api/core";
import { useAiContextStore } from "../store/useAiContextStore";
import { useProjectStore } from "../store/useProjectStore";
import { AI_CONTEXT_KEY } from "../utils/aiContextKeys";
import {
  type CashflowSegment,
  type SegmentSideScope,
  type TaxItem,
  buildDirectCashflowFromSegments,
  distributionFromCashflow,
  cashflowPayloadValues,
  sumInclTaxItems,
  clampCashflowYear,
  useIctState
} from "./useIctState";
import {
  ICT_SUBJECT_DEFINITIONS,
  getSubjectExcelDisplayName,
  normalizeCustomSubjectName,
  type IctSubjectDefinition,
} from "../lib/ictSubjectCatalog";
import {
  serializeBalanceAllocationRule,
} from "../lib/ictBalanceAllocation";
import {
  applyLockedTotalStructureAmountsToState,
  applySubjectInclAmountToState,
  buildLockedTotalStructureSamplePoints,
  readSubjectInclAmount,
  type LockedTotalStructureContext,
  type ResolvedReverseCalculationContext,
  type ReverseSubjectOption,
  type ReverseSubjectState,
} from "../lib/ictReverseCalculation";
import {
  buildAnnualCashflowFromSubjectFundingPlans,
  validateSubjectFundingPlanCoverage,
  type SubjectFundingPlanCoverageSubject,
} from "../lib/ictSubjectFundingPlan";

// Labels mapping
const cashflowModelLabels: Record<string, string> = {
  model_a: "模型 A (100%首年)",
  model_b: "模型 B (按周期等额)",
  model_c: "模型 C (首年95%, 末年5%)",
  model_d: "模型 D (自定义比例)",
  model_e: "模型 E (分板块计划)",
};

const formatDistribution = (arr: number[]) => {
  return "[" + arr.map(v => (v * 100).toFixed(1) + "%").join(", ") + "]";
};

const formatCurrency = (v: number) => new Intl.NumberFormat('zh-CN', { style: 'currency', currency: 'CNY' }).format(v);
const formatPercent = (v: number) => (v * 100).toFixed(2) + "%";
const METRIC_EPSILON = 0.0001;
const MONEY_EPSILON = 0.004;
const roundMoney = (value: number) => Number((Number.isFinite(value) ? value : 0).toFixed(2));

type ModelEAmountBucket = {
  side: "revenue" | "cost";
  scope: SegmentSideScope;
  label: string;
};

type ModelEStructureTransfer = {
  bucket: ModelEAmountBucket;
  deltaIncl: number;
  sourceTax: number;
  reason: string;
};

type ModelEStructureSyncResult = {
  valid: boolean;
  segments: CashflowSegment[];
  message?: string;
  transfers: ModelEStructureTransfer[];
};

const getModelEAmountBucketForSubject = (subject: IctSubjectDefinition): ModelEAmountBucket | null => {
  if (subject.groupId === "revIt") return { side: "revenue", scope: "it", label: "收入 IT 板块" };
  if (subject.groupId === "revCt") return { side: "revenue", scope: "ct", label: "收入 CT 板块" };
  if (subject.groupId === "revNonItCt") return { side: "revenue", scope: "non_it_ct", label: "收入 非IT/CT 板块" };
  if (subject.groupId === "costIt") return { side: "cost", scope: "it", label: "投入 IT 板块" };
  if (subject.groupId === "costCt") return { side: "cost", scope: "ct", label: "投入 CT 板块" };
  if (subject.groupId === "costMix" && subject.key === "non_it_ct") return { side: "cost", scope: "non_it_ct", label: "投入 非IT/CT 板块" };
  if (subject.groupId === "costMix") return { side: "cost", scope: "mix", label: "投入 综合类板块" };
  return null;
};

const getPairedCostSubjectForRevenueSubject = (subject: IctSubjectDefinition) => {
  if (subject.groupId !== "revCt") return null;
  if (subject.key === "product") {
    return ICT_SUBJECT_DEFINITIONS.find(item => item.subjectCode === "cost_ct_other") || null;
  }
  if (subject.key === "line") {
    return ICT_SUBJECT_DEFINITIONS.find(item => item.subjectCode === "cost_ct_bandwidth") || null;
  }
  return null;
};

const isBalanceRuleConfigured = (rule: ReturnType<typeof useIctState>["balanceAllocation"]["investment"]) =>
  Boolean(rule.enabled && rule.totalInclAmount !== null && rule.balancingSubject);

const serializeTaxItemForPayload = (item: TaxItem) => {
  const customSubjectName = normalizeCustomSubjectName(item.customSubjectName);
  const billingSubjectName = normalizeCustomSubjectName(item.billingSubjectName);
  return {
    incl_tax: String(item.incl),
    tax_rate: String(item.tax),
    ...(customSubjectName ? { custom_subject_name: customSubjectName } : {}),
    ...(billingSubjectName ? { billing_subject_name: billingSubjectName } : {}),
  };
};

export function useIctCalculations(state: ReturnType<typeof useIctState>) {
  const updateData = useAiContextStore(stateStore => stateStore.updateBusinessData);
  const syncTimerRef = useRef<NodeJS.Timeout | null>(null);

  // --- Calculation Results ---
  const [cashflowTable, setCashflowTable] = useState<any[]>([]);
  const [metrics, setMetrics] = useState<any>({
    npv: 0, npv_rate: 0, margin_rate: 0, dynamic_payback: "--", irr: "--",
    it_npv: 0, it_npv_rate: 0, it_margin_rate: 0
  });

  // --- Selection Fee Calc State ---
  const [selQuote, setSelQuote] = useState<string>("");
  const [selMarkup, setSelMarkup] = useState<string>("50");
  const [selActualCost, setSelActualCost] = useState<string>("");
  const [selFee, setSelFee] = useState<string>("");
  const [selLimit, setSelLimit] = useState<string>("");

  // --- Smart Reverse State ---
  const [revMode, setRevMode] = useState<"cost" | "revenue">("cost");
  const [revTargetType, setRevTargetType] = useState<"margin" | "npv_rate">("margin");
  const [revTargetValue, setRevTargetValue] = useState<string>("0.15");
  const [revSubjectRefKey, setRevSubjectRefKey] = useState<string>("");

  // Helper selectors
  const directSegmentCashflow = state.cashflowSegments ? buildDirectCashflowFromSegments(state.cashflowSegments) : { rev: [], cost: [], itRev: [], itCost: [] };
  const revenueInclTotal = sumInclTaxItems([...Object.values(state.revIt), ...Object.values(state.revCt), state.revNonItCt]);
  const costInclTotal = sumInclTaxItems([...Object.values(state.costIt), ...Object.values(state.costCt), ...Object.values(state.costMix)]);
  const segmentRevenueInclTotal = state.cashflowSegments.reduce((sum, segment) => sum + (segment.revenueValue || 0), 0);
  const segmentCostInclTotal = state.cashflowSegments.reduce((sum, segment) => sum + (segment.costValue || 0), 0);

  const effectiveDistRev = state.distRev;
  const effectiveDistCost = state.distCost;

  const buildSubjectFundingCoverageSubjects = (sources?: {
    revItState?: typeof state.revIt;
    revCtState?: typeof state.revCt;
    revNonItCtState?: typeof state.revNonItCt;
    costItState?: typeof state.costIt;
    costCtState?: typeof state.costCt;
    costMixState?: typeof state.costMix;
  }): SubjectFundingPlanCoverageSubject[] => {
    const revItSource = sources?.revItState ?? state.revIt;
    const revCtSource = sources?.revCtState ?? state.revCt;
    const revNonItCtSource = sources?.revNonItCtState ?? state.revNonItCt;
    const costItSource = sources?.costItState ?? state.costIt;
    const costCtSource = sources?.costCtState ?? state.costCt;
    const costMixSource = sources?.costMixState ?? state.costMix;

    const resolveItem = (subject: IctSubjectDefinition): TaxItem | null => {
      if (subject.groupId === "revIt") return revItSource[subject.key as keyof typeof revItSource] || null;
      if (subject.groupId === "revCt") return revCtSource[subject.key as keyof typeof revCtSource] || null;
      if (subject.groupId === "revNonItCt") return revNonItCtSource;
      if (subject.groupId === "costIt") return costItSource[subject.key as keyof typeof costItSource] || null;
      if (subject.groupId === "costCt") return costCtSource[subject.key as keyof typeof costCtSource] || null;
      if (subject.groupId === "costMix") return costMixSource[subject.key as keyof typeof costMixSource] || null;
      return null;
    };

    return ICT_SUBJECT_DEFINITIONS.map(subject => {
      const item = resolveItem(subject);
      return {
        subjectRef: {
          side: subject.side,
          groupId: subject.groupId,
          key: subject.key,
        },
        displayName: getSubjectExcelDisplayName(subject, item),
        subjectAmountIncl: Number(item?.incl ?? 0),
        taxRate: Number(item?.tax ?? 0),
        isItScope: subject.groupId === "revIt" || subject.groupId === "costIt",
      };
    });
  };

  const subjectFundingCoverageSubjects = buildSubjectFundingCoverageSubjects();
  const subjectFundingCoverage = validateSubjectFundingPlanCoverage(
    subjectFundingCoverageSubjects,
    state.subjectFundingPlans,
  );
  const subjectFundingAnnualCashflow = buildAnnualCashflowFromSubjectFundingPlans(
    subjectFundingCoverageSubjects,
    state.subjectFundingPlans,
  );
  const subjectFundingCalculationBlocked =
    state.cashflowCalculationSource === "subject_funding_plans" && !subjectFundingCoverage.valid;

  const buildInputDataPayload = (options?: {
    segments?: CashflowSegment[];
    revItState?: typeof state.revIt;
    revCtState?: typeof state.revCt;
    revNonItCtState?: typeof state.revNonItCt;
    costItState?: typeof state.costIt;
    costCtState?: typeof state.costCt;
    costMixState?: typeof state.costMix;
  }) => {
    const segmentsForPayload = options?.segments ?? state.cashflowSegments;
    const directCashflowForPayload = buildDirectCashflowFromSegments(segmentsForPayload);
    const revDistributionForPayload = state.cashflowModel === 'model_e' && state.segmentValueMode === "amount"
      ? distributionFromCashflow(directCashflowForPayload.rev)
      : effectiveDistRev;
    const costDistributionForPayload = state.cashflowModel === 'model_e' && state.segmentValueMode === "amount"
      ? distributionFromCashflow(directCashflowForPayload.cost)
      : effectiveDistCost;
    const revItForPayload = options?.revItState ?? state.revIt;
    const revCtForPayload = options?.revCtState ?? state.revCt;
    const revNonItCtForPayload = options?.revNonItCtState ?? state.revNonItCt;
    const costItForPayload = options?.costItState ?? state.costIt;
    const costCtForPayload = options?.costCtState ?? state.costCt;
    const costMixForPayload = options?.costMixState ?? state.costMix;
    const calculationSource = state.cashflowCalculationSource || "legacy_model";
    const subjectRowsForPayload = buildSubjectFundingCoverageSubjects({
      revItState: revItForPayload,
      revCtState: revCtForPayload,
      revNonItCtState: revNonItCtForPayload,
      costItState: costItForPayload,
      costCtState: costCtForPayload,
      costMixState: costMixForPayload,
    });
    const coverageForPayload = validateSubjectFundingPlanCoverage(subjectRowsForPayload, state.subjectFundingPlans);
    const annualCashflowForPayload = buildAnnualCashflowFromSubjectFundingPlans(subjectRowsForPayload, state.subjectFundingPlans);
    const useSubjectFundingCashflow = calculationSource === "subject_funding_plans" && coverageForPayload.valid;
    const useLegacyDirectSegments = calculationSource === "legacy_model"
      && state.cashflowModel === 'model_e'
      && state.segmentValueMode === "amount";

    return {
      project_name: state.projName,
      customer_name: state.customerName,
      property_rights: state.propertyRights,
      discount_rate: String(state.discountRate),
      project_years: state.projectYears,
      cashflow_model: state.cashflowModel,
      cashflow_calculation_source: calculationSource,
      rev_distribution: revDistributionForPayload,
      cost_distribution: costDistributionForPayload,
      cashflow_segment_value_mode: state.segmentValueMode,
      cashflow_segments: state.cashflowModel === 'model_e' ? segmentsForPayload : [],
      project_background: state.projectBackground,
      revenue_balance_rule: serializeBalanceAllocationRule(state.balanceAllocation.revenue),
      investment_balance_rule: serializeBalanceAllocationRule(state.balanceAllocation.investment),
      subject_funding_plans: state.subjectFundingPlans,
      rev_cashflow_excl: useSubjectFundingCashflow
        ? cashflowPayloadValues(annualCashflowForPayload.annualRevenueExcl)
        : useLegacyDirectSegments
          ? cashflowPayloadValues(directCashflowForPayload.rev)
          : null,
      cost_cashflow_excl: useSubjectFundingCashflow
        ? cashflowPayloadValues(annualCashflowForPayload.annualCostExcl)
        : useLegacyDirectSegments
          ? cashflowPayloadValues(directCashflowForPayload.cost)
          : null,
      it_rev_cashflow_excl: useSubjectFundingCashflow
        ? cashflowPayloadValues(annualCashflowForPayload.annualItRevenueExcl)
        : useLegacyDirectSegments
          ? cashflowPayloadValues(directCashflowForPayload.itRev)
          : null,
      it_cost_cashflow_excl: useSubjectFundingCashflow
        ? cashflowPayloadValues(annualCashflowForPayload.annualItCostExcl)
        : useLegacyDirectSegments
          ? cashflowPayloadValues(directCashflowForPayload.itCost)
          : null,
      ignore_tail_difference: state.ignoredTailValue !== null,
      tail_difference_value: state.ignoredTailValue || "0",
      rev_it_integration: serializeTaxItemForPayload(revItForPayload.integration),
      rev_it_maintenance: serializeTaxItemForPayload(revItForPayload.maintenance),
      rev_it_device_sales: serializeTaxItemForPayload(revItForPayload.device_sales),
      rev_it_device_lease: serializeTaxItemForPayload(revItForPayload.device_lease),
      rev_it_other: serializeTaxItemForPayload(revItForPayload.other),
      rev_it_cloud: serializeTaxItemForPayload(revItForPayload.cloud),
      rev_ct_line: serializeTaxItemForPayload(revCtForPayload.line),
      rev_ct_product: serializeTaxItemForPayload(revCtForPayload.product),
      rev_non_it_ct: serializeTaxItemForPayload(revNonItCtForPayload),
      cost_it_device: serializeTaxItemForPayload(costItForPayload.device),
      cost_it_construction: serializeTaxItemForPayload(costItForPayload.construction),
      cost_it_survey: serializeTaxItemForPayload(costItForPayload.survey),
      cost_it_integration: serializeTaxItemForPayload(costItForPayload.integration),
      cost_it_other: serializeTaxItemForPayload(costItForPayload.other),
      cost_it_maintenance: serializeTaxItemForPayload(costItForPayload.maintenance),
      cost_it_running: serializeTaxItemForPayload(costItForPayload.running),
      cost_it_bidding: serializeTaxItemForPayload(costItForPayload.bidding),
      cost_it_design_eval: serializeTaxItemForPayload(costItForPayload.design_eval),
      cost_it_audit: serializeTaxItemForPayload(costItForPayload.audit),
      cost_ct_construction: serializeTaxItemForPayload(costCtForPayload.construction),
      cost_ct_maintenance: serializeTaxItemForPayload(costCtForPayload.maintenance),
      cost_ct_other: serializeTaxItemForPayload(costCtForPayload.other),
      cost_ct_bandwidth: serializeTaxItemForPayload(costCtForPayload.bandwidth),
      cost_ct_renewal: serializeTaxItemForPayload(costCtForPayload.renewal),
      cost_non_it_ct: serializeTaxItemForPayload(costMixForPayload.non_it_ct),
      cost_mix_marketing: serializeTaxItemForPayload(costMixForPayload.marketing),
      cost_mix_channel: serializeTaxItemForPayload(costMixForPayload.channel),
      cost_mix_other: serializeTaxItemForPayload(costMixForPayload.other),
    };
  };

  const getInputDataPayload = () => buildInputDataPayload();

  const performCalculation = useCallback(async () => {
    if (state.cashflowCalculationSource === "subject_funding_plans" && !subjectFundingCoverage.valid) {
      console.warn("Subject funding plan coverage is invalid. Keeping previous calculation result.");
      return;
    }

    try {
      const res: any = await invoke('calculate_ict_benefit', { input: getInputDataPayload() });
      if (res) {
        setCashflowTable(res.cashflow);
        setMetrics(res);
      }
    } catch (e) {
      console.error(e);
    }
  }, [state, subjectFundingCoverage]);

  // Recalculate whenever state variables change
  useEffect(() => {
    performCalculation();
  }, [
    state.revIt, state.revCt, state.revNonItCt,
    state.costIt, state.costCt, state.costMix,
    state.projectYears, state.discountRate, state.cashflowModel,
    state.distRev, state.distCost, state.segmentValueMode, state.cashflowSegments,
    state.cashflowCalculationSource, state.subjectFundingPlans,
    performCalculation
  ]);

  // --- AI Context Sync ---
  const buildAiContextPayload = useCallback((includeCalculated = false, overrides?: { metrics?: any; cashflow?: any[]; extra?: Record<string, any> }) => ({
    monetary_unit: '元',
    projectId: useProjectStore.getState().currentProject?.id ?? null,
    currency: 'CNY',
    ...getInputDataPayload(),
    project_background: state.projectBackground,
    metrics: includeCalculated ? (overrides?.metrics ?? metrics) : null,
    cashflow: includeCalculated ? (overrides?.cashflow ?? cashflowTable) : [],
    ...(overrides?.extra ?? {}),
  }), [state, metrics, cashflowTable]);

  useEffect(() => {
    updateData(AI_CONTEXT_KEY.ICT_CORE, buildAiContextPayload(false));
  }, [
    state.revIt, state.revCt, state.revNonItCt,
    state.costIt, state.costCt, state.costMix,
    state.projectBackground, state.projName, state.customerName, state.propertyRights,
    state.discountRate, state.projectYears, state.cashflowModel,
    state.distRev, state.distCost, state.segmentValueMode, state.cashflowSegments,
    state.ignoredTailValue, state.balanceAllocation, state.cashflowCalculationSource, state.subjectFundingPlans, updateData, buildAiContextPayload
  ]);

  useEffect(() => {
    // Debounced sync (500ms)
    if (syncTimerRef.current) clearTimeout(syncTimerRef.current);

    syncTimerRef.current = setTimeout(() => {
      const payload = buildAiContextPayload(true);
      updateData(AI_CONTEXT_KEY.ICT_CORE, payload);
      console.log("AI Context Synced: ICT");
    }, 500);

    return () => {
      if (syncTimerRef.current) clearTimeout(syncTimerRef.current);
    };
  }, [
    state.revIt, state.revCt, state.revNonItCt,
    state.costIt, state.costCt, state.costMix,
    state.projectBackground, metrics, cashflowTable, state.projName, state.customerName,
    state.propertyRights, state.discountRate, state.projectYears, state.cashflowModel,
    state.distRev, state.distCost, state.segmentValueMode, state.cashflowSegments,
    state.ignoredTailValue, state.balanceAllocation, state.cashflowCalculationSource, state.subjectFundingPlans, updateData, buildAiContextPayload
  ]);

  const handleSelFeeChange = async (type: 'quote' | 'markup' | 'limit', val: string) => {
    if (type === 'quote') setSelQuote(val);
    if (type === 'markup') setSelMarkup(val);
    if (type === 'limit') setSelLimit(val);

    const currentQuote = type === 'quote' ? val : selQuote;
    const currentMarkup = type === 'markup' ? val : selMarkup;
    const currentLimit = type === 'limit' ? val : selLimit;

    try {
      if (type === 'quote' || type === 'markup') {
        const res: any = await invoke('calculate_selection_fee', { quote: currentQuote || "0", markup: currentMarkup || "0" });
        setSelLimit(res.final_limit);
        setSelActualCost(res.actual_cost);
        setSelFee(res.selection_fee);
      } else if (type === 'limit') {
        const res: any = await invoke('reverse_calculate_selection_fee', { limit: currentLimit || "0", markup: currentMarkup || "0" });
        setSelQuote(res.quote);
        setSelActualCost(res.actual_cost);
        setSelFee(res.selection_fee);
      }
    } catch(e) {
      console.error("甄选限价计算失败:", e);
    }
  };

  const applySelectionLimit = () => {
    if (selLimit) {
      state.updateTaxItem('costIt', 'integration', 'incl', Number(selLimit));
      if (selFee) {
        state.updateTaxItem('costIt', 'bidding', 'incl', Number(selFee));
      }
    }
  };

  const selectReverseSegmentIndex = (segments: CashflowSegment[], side: "revenue" | "cost") => {
    const scopeKey = side === "revenue" ? "revenueScope" : "costScope";
    const valueKey = side === "revenue" ? "revenueValue" : "costValue";

    const firstIndexMatching = (predicate: (segment: CashflowSegment) => boolean) => segments.findIndex(predicate);
    const priorityChecks = [
      (segment: CashflowSegment) => segment[scopeKey] === "it" && segment[valueKey] <= 0,
      (segment: CashflowSegment) => segment[scopeKey] === "it" && segment.name.includes("集成"),
      (segment: CashflowSegment) => segment[scopeKey] === "it",
      (segment: CashflowSegment) => segment.name.includes("集成") && segment[valueKey] <= 0,
      (segment: CashflowSegment) => segment[valueKey] <= 0,
    ];

    for (const check of priorityChecks) {
      const index = firstIndexMatching(check);
      if (index >= 0) return index;
    }

    return segments.length > 0 ? 0 : -1;
  };

  const scaleCustomAnnualValuesForReverse = (segment: CashflowSegment, side: "revenue" | "cost", amount: number) => {
    const mode = side === "revenue" ? segment.revenueMode : segment.costMode;
    const currentValues = side === "revenue" ? segment.revenueAnnualValues : segment.costAnnualValues;

    if (mode !== "custom") return currentValues;

    const duration = Math.max(1, Math.min(clampCashflowYear(segment.serviceYears), 10 - (clampCashflowYear(segment.startYear) - 1)));
    const values = Array.from({ length: duration }, (_, idx) => Number(currentValues[idx] || 0));
    const sum = values.reduce((total, value) => total + value, 0);

    if (sum > 0) return values.map(value => Number((amount * value / sum).toFixed(2)));
    return Array.from({ length: duration }, () => Number((amount / duration).toFixed(2)));
  };

  const applyReverseValueToSegments = (
    segments: CashflowSegment[],
    side: "revenue" | "cost",
    segmentIndex: number,
    amount: number,
    tax: number
  ) => segments.map((segment, index) => {
    if (index !== segmentIndex) return segment;

    if (side === "revenue") {
      return {
        ...segment,
        revenueValue: Number(amount.toFixed(2)),
        revenueTax: tax,
        revenueScope: "it" as const,
        revenueAnnualValues: scaleCustomAnnualValuesForReverse(segment, side, amount),
      };
    }

    return {
      ...segment,
      costValue: Number(amount.toFixed(2)),
      costTax: tax,
      costScope: "it" as const,
      costAnnualValues: scaleCustomAnnualValuesForReverse(segment, side, amount),
    };
  });

  const applyModelETransferToBucket = (
    segments: CashflowSegment[],
    transfer: ModelEStructureTransfer,
  ): ModelEStructureSyncResult => {
    const delta = roundMoney(transfer.deltaIncl);
    if (Math.abs(delta) <= MONEY_EPSILON) {
      return { valid: true, segments, transfers: [transfer] };
    }

    const scopeKey = transfer.bucket.side === "revenue" ? "revenueScope" : "costScope";
    const valueKey = transfer.bucket.side === "revenue" ? "revenueValue" : "costValue";
    const taxKey = transfer.bucket.side === "revenue" ? "revenueTax" : "costTax";
    const annualValuesKey = transfer.bucket.side === "revenue" ? "revenueAnnualValues" : "costAnnualValues";
    const matchingIndexes = segments
      .map((segment, index) => ({ segment, index }))
      .filter(({ segment }) => segment[scopeKey] === transfer.bucket.scope)
      .map(({ index }) => index);

    if (matchingIndexes.length === 0) {
      return {
        valid: false,
        segments,
        transfers: [transfer],
        message: `${transfer.bucket.label}没有可同步的分板块金额计划，无法执行 model_e 结构反算。`,
      };
    }

    const currentTotal = roundMoney(matchingIndexes.reduce((sum, index) => sum + Number(segments[index][valueKey] || 0), 0));
    if (delta < 0 && currentTotal + delta < -MONEY_EPSILON) {
      return {
        valid: false,
        segments,
        transfers: [transfer],
        message: `${transfer.bucket.label}当前金额不足，结构调整会导致板块金额为负。`,
      };
    }

    const nextSegments = [...segments];
    if (delta > 0) {
      const positiveIndexes = matchingIndexes.filter(index => Number(segments[index][valueKey] || 0) > MONEY_EPSILON);
      const targetIndexes = positiveIndexes.length > 0 ? positiveIndexes : [matchingIndexes[0]];
      const baseTotal = positiveIndexes.length > 0
        ? positiveIndexes.reduce((sum, index) => sum + Number(segments[index][valueKey] || 0), 0)
        : 0;
      let remaining = delta;
      targetIndexes.forEach((index, order) => {
        const currentValue = Number(nextSegments[index][valueKey] || 0);
        const share = order === targetIndexes.length - 1
          ? remaining
          : roundMoney(delta * (baseTotal > 0 ? currentValue / baseTotal : 1 / targetIndexes.length));
        remaining = roundMoney(remaining - share);
        const nextAmount = roundMoney(currentValue + share);
        nextSegments[index] = {
          ...nextSegments[index],
          [valueKey]: nextAmount,
          [taxKey]: currentValue > MONEY_EPSILON ? nextSegments[index][taxKey] : transfer.sourceTax,
          [annualValuesKey]: scaleCustomAnnualValuesForReverse(nextSegments[index], transfer.bucket.side, nextAmount),
        };
      });
      return { valid: true, segments: nextSegments, transfers: [transfer] };
    }

    let remainingDecrease = Math.abs(delta);
    const positiveIndexes = matchingIndexes.filter(index => Number(segments[index][valueKey] || 0) > MONEY_EPSILON);
    if (positiveIndexes.length === 0) {
      return {
        valid: false,
        segments,
        transfers: [transfer],
        message: `${transfer.bucket.label}没有可减少的正金额板块，无法执行本次结构转移。`,
      };
    }

    positiveIndexes.forEach((index, order) => {
      const currentValue = Number(nextSegments[index][valueKey] || 0);
      const share = order === positiveIndexes.length - 1
        ? remainingDecrease
        : roundMoney(Math.abs(delta) * currentValue / currentTotal);
      const appliedDecrease = Math.min(currentValue, share);
      remainingDecrease = roundMoney(remainingDecrease - appliedDecrease);
      const nextAmount = roundMoney(currentValue - appliedDecrease);
      nextSegments[index] = {
        ...nextSegments[index],
        [valueKey]: nextAmount,
        [annualValuesKey]: scaleCustomAnnualValuesForReverse(nextSegments[index], transfer.bucket.side, nextAmount),
      };
    });

    if (remainingDecrease > MONEY_EPSILON) {
      return {
        valid: false,
        segments,
        transfers: [transfer],
        message: `${transfer.bucket.label}年度金额计划不足，结构调整会产生负金额。`,
      };
    }

    return { valid: true, segments: nextSegments, transfers: [transfer] };
  };

  const mergeModelETransfers = (transfers: ModelEStructureTransfer[]) => {
    const merged = new Map<string, ModelEStructureTransfer>();
    transfers.forEach(transfer => {
      const key = `${transfer.bucket.side}:${transfer.bucket.scope}`;
      const existing = merged.get(key);
      if (!existing) {
        merged.set(key, { ...transfer, deltaIncl: roundMoney(transfer.deltaIncl) });
        return;
      }
      merged.set(key, {
        ...existing,
        deltaIncl: roundMoney(existing.deltaIncl + transfer.deltaIncl),
        reason: `${existing.reason}; ${transfer.reason}`,
      });
    });
    return Array.from(merged.values()).filter(transfer => Math.abs(transfer.deltaIncl) > MONEY_EPSILON);
  };

  const applyModelEStructureTransfer = (params: {
    segments: CashflowSegment[];
    structure: LockedTotalStructureContext;
    candidateTargetIncl: number;
    candidateBalancingIncl: number;
  }): ModelEStructureSyncResult => {
    const { segments, structure, candidateTargetIncl, candidateBalancingIncl } = params;
    const targetBucket = getModelEAmountBucketForSubject(structure.targetSubject);
    const balancingBucket = getModelEAmountBucketForSubject(structure.balancingSubject);
    if (!targetBucket || !balancingBucket) {
      return {
        valid: false,
        segments,
        transfers: [],
        message: "当前结构反算组合无法映射到分板块金额计划，请更换反算科目或差额承接科目。",
      };
    }

    const rawTransfers: ModelEStructureTransfer[] = [
      {
        bucket: targetBucket,
        deltaIncl: roundMoney(candidateTargetIncl - structure.beforeTargetInclAmount),
        sourceTax: Number(structure.targetItem?.tax ?? 0),
        reason: structure.targetDisplayName,
      },
      {
        bucket: balancingBucket,
        deltaIncl: roundMoney(candidateBalancingIncl - structure.beforeBalancingInclAmount),
        sourceTax: Number(structure.balancingItem?.tax ?? 0),
        reason: structure.balancingDisplayName,
      },
    ];

    const appendPairedCostTransfer = (
      subject: IctSubjectDefinition,
      beforeIncl: number,
      afterIncl: number,
      sourceTax: number,
      reason: string,
    ) => {
      const pairedCostSubject = getPairedCostSubjectForRevenueSubject(subject);
      if (!pairedCostSubject) return;
      const pairedBucket = getModelEAmountBucketForSubject(pairedCostSubject);
      if (!pairedBucket) return;
      rawTransfers.push({
        bucket: pairedBucket,
        deltaIncl: roundMoney(afterIncl - beforeIncl),
        sourceTax,
        reason: `${reason}联动投入`,
      });
    };

    appendPairedCostTransfer(
      structure.targetSubject,
      structure.beforeTargetInclAmount,
      candidateTargetIncl,
      Number(structure.targetItem?.tax ?? 0),
      structure.targetDisplayName,
    );
    appendPairedCostTransfer(
      structure.balancingSubject,
      structure.beforeBalancingInclAmount,
      candidateBalancingIncl,
      Number(structure.balancingItem?.tax ?? 0),
      structure.balancingDisplayName,
    );

    let nextSegments = segments;
    const transfers = mergeModelETransfers(rawTransfers);
    for (const transfer of transfers) {
      const result = applyModelETransferToBucket(nextSegments, transfer);
      if (!result.valid) {
        return { ...result, transfers };
      }
      nextSegments = result.segments;
    }

    return { valid: true, segments: nextSegments, transfers };
  };

  const structureHasRevenueToCostLink = (structure: LockedTotalStructureContext) =>
    Boolean(
      getPairedCostSubjectForRevenueSubject(structure.targetSubject)
      || getPairedCostSubjectForRevenueSubject(structure.balancingSubject),
    );

  const getCurrentReverseSubjectState = (): ReverseSubjectState => ({
    revIt: state.revIt,
    revCt: state.revCt,
    revNonItCt: state.revNonItCt,
    costIt: state.costIt,
    costCt: state.costCt,
    costMix: state.costMix,
  });

  const buildReverseCandidate = (
    selectedSubject: ReverseSubjectOption,
    amount: number,
    modelEAmountSegmentIndex: number | null,
  ) => {
    const nextSubjectState = applySubjectInclAmountToState(
      getCurrentReverseSubjectState(),
      selectedSubject.subject,
      amount,
    );
    const nextSegments = modelEAmountSegmentIndex === null
      ? state.cashflowSegments
      : applyReverseValueToSegments(
          state.cashflowSegments,
          selectedSubject.subject.side,
          modelEAmountSegmentIndex,
          amount,
          Number(selectedSubject.item?.tax ?? 0),
        );
    const payload = buildInputDataPayload({
      segments: nextSegments,
      revItState: nextSubjectState.revIt as typeof state.revIt,
      revCtState: nextSubjectState.revCt as typeof state.revCt,
      revNonItCtState: nextSubjectState.revNonItCt as typeof state.revNonItCt,
      costItState: nextSubjectState.costIt as typeof state.costIt,
      costCtState: nextSubjectState.costCt as typeof state.costCt,
      costMixState: nextSubjectState.costMix as typeof state.costMix,
    });

    return { nextSubjectState, nextSegments, payload };
  };

  const buildLockedTotalStructureCandidate = (
    structure: LockedTotalStructureContext,
    targetAmount: number,
  ) => {
    const safeTargetAmount = roundMoney(Math.max(0, Math.min(structure.reallocatablePoolInclAmount, targetAmount)));
    const balancingAmount = roundMoney(structure.reallocatablePoolInclAmount - safeTargetAmount);
    const modelEAmountMode = state.cashflowModel === "model_e" && state.segmentValueMode === "amount";
    const modelESync = modelEAmountMode
      ? applyModelEStructureTransfer({
          segments: state.cashflowSegments,
          structure,
          candidateTargetIncl: safeTargetAmount,
          candidateBalancingIncl: balancingAmount,
        })
      : { valid: true, segments: state.cashflowSegments, transfers: [] };
    if (!modelESync.valid) {
      return {
        valid: false,
        message: modelESync.message || "当前 model_e 分板块金额计划无法同步本次结构候选。",
        targetAmount: safeTargetAmount,
        balancingAmount,
        nextSubjectState: getCurrentReverseSubjectState(),
        nextSegments: state.cashflowSegments,
        payload: null,
        modelETransfers: modelESync.transfers,
      };
    }
    const nextSubjectState = applyLockedTotalStructureAmountsToState(
      getCurrentReverseSubjectState(),
      structure.targetSubject,
      safeTargetAmount,
      structure.balancingSubject,
      balancingAmount,
    );
    const payload = buildInputDataPayload({
      segments: modelESync.segments,
      revItState: nextSubjectState.revIt as typeof state.revIt,
      revCtState: nextSubjectState.revCt as typeof state.revCt,
      revNonItCtState: nextSubjectState.revNonItCt as typeof state.revNonItCt,
      costItState: nextSubjectState.costIt as typeof state.costIt,
      costCtState: nextSubjectState.costCt as typeof state.costCt,
      costMixState: nextSubjectState.costMix as typeof state.costMix,
    });

    return {
      valid: true,
      targetAmount: safeTargetAmount,
      balancingAmount,
      nextSubjectState,
      nextSegments: modelESync.segments,
      payload,
      modelETransfers: modelESync.transfers,
    };
  };

  const getMetricValue = (result: any) => {
    const metricValue = Number(revTargetType === "margin" ? result.margin_rate : result.npv_rate);
    return Number.isFinite(metricValue) ? metricValue : 0;
  };

  const performLockedTotalStructureReverseCalculation = async (
    selectedSubject: ReverseSubjectOption,
    structure: LockedTotalStructureContext,
    target: number,
  ) => {
    const modelEAmountMode = state.cashflowModel === "model_e" && state.segmentValueMode === "amount";
    if (
      modelEAmountMode
      && structure.side === "revenue"
      && structureHasRevenueToCostLink(structure)
      && isBalanceRuleConfigured(state.balanceAllocation.investment)
    ) {
      return alert("当前反算科目会联动调整投入金额，但投入侧同时启用了总额锁定与差额承接。该组合涉及双侧联动结构调整，当前暂不支持，请先清空一侧承接规则或选择其他反算科目。");
    }

    if (structure.targetSubject.subjectCode === structure.balancingSubject.subjectCode) {
      return alert("结构反算目标科目不能与差额承接科目相同。");
    }
    if (structure.reallocatablePoolInclAmount < 0) {
      return alert("当前锁定总额小于固定科目合计，无法执行结构反算。");
    }

    const evaluate = async (targetAmount: number) => {
      const candidate = buildLockedTotalStructureCandidate(structure, targetAmount);
      if (!candidate.valid || !candidate.payload) {
        return {
          ...candidate,
          result: null,
          metricValue: 0,
        };
      }
      const result: any = await invoke("calculate_ict_benefit", { input: candidate.payload });
      return {
        ...candidate,
        result,
        metricValue: getMetricValue(result),
      };
    };

    try {
      if (state.ignoredDataHash !== null) {
        state.setIgnoredDataHash(null);
        state.setIgnoredTailValue(null);
      }

      const sampleAmounts = buildLockedTotalStructureSamplePoints(
        structure.reallocatablePoolInclAmount,
        structure.beforeTargetInclAmount,
      );
      const samplePoints = [];
      for (const amount of sampleAmounts) {
        samplePoints.push(await evaluate(amount));
      }

      const validSamplePoints = samplePoints.filter(point => point.valid && point.payload && point.result);
      if (validSamplePoints.length === 0) {
        const firstInvalid = samplePoints.find(point => !point.valid);
        return alert(firstInvalid?.message || "当前分板块金额计划无法支持任何结构反算候选点，请调整板块金额计划后再试。");
      }

      const metricValues = validSamplePoints.map(point => point.metricValue);
      const minMetric = Math.min(...metricValues);
      const maxMetric = Math.max(...metricValues);
      const targetName = revTargetType === "margin" ? "目标毛利润率" : "目标净现值率";

      if (maxMetric - minMetric < METRIC_EPSILON) {
        return alert(`当前结构调整对${targetName}不敏感：可重分配池在 ${formatCurrency(structure.reallocatablePoolInclAmount)} 内变化时，指标仅从 ${formatPercent(minMetric)} 到 ${formatPercent(maxMetric)}。请调整现金流、税率或目标科目后再试。`);
      }

      if (target < minMetric - METRIC_EPSILON || target > maxMetric + METRIC_EPSILON) {
        return alert(`当前锁定总额结构下无法达到目标值。${targetName}可达范围约为 ${formatPercent(minMetric)} - ${formatPercent(maxMetric)}，当前目标为 ${formatPercent(target)}。`);
      }

      const solutions: Array<{ targetAmount: number; point: Awaited<ReturnType<typeof evaluate>> }> = [];
      validSamplePoints.forEach(point => {
        if (Math.abs(point.metricValue - target) <= METRIC_EPSILON) {
          solutions.push({ targetAmount: point.targetAmount, point });
        }
      });

      for (let i = 0; i < validSamplePoints.length - 1; i++) {
        const left = validSamplePoints[i];
        const right = validSamplePoints[i + 1];
        const leftDiff = left.metricValue - target;
        const rightDiff = right.metricValue - target;
        if (leftDiff === 0 || rightDiff === 0 || leftDiff * rightDiff > 0) continue;

        let low = left;
        let high = right;
        let best = Math.abs(leftDiff) <= Math.abs(rightDiff) ? left : right;
        const increasing = right.metricValue >= left.metricValue;

        for (let step = 0; step < 45; step++) {
          const midAmount = roundMoney((low.targetAmount + high.targetAmount) / 2);
          const mid = await evaluate(midAmount);
          if (!mid.valid || !mid.payload || !mid.result) {
            break;
          }
          if (Math.abs(mid.metricValue - target) < Math.abs(best.metricValue - target)) {
            best = mid;
          }
          if (Math.abs(mid.metricValue - target) <= METRIC_EPSILON || Math.abs(high.targetAmount - low.targetAmount) <= MONEY_EPSILON) {
            best = mid;
            break;
          }
          if (increasing) {
            if (mid.metricValue < target) low = mid;
            else high = mid;
          } else if (mid.metricValue > target) {
            low = mid;
          } else {
            high = mid;
          }
        }
        solutions.push({ targetAmount: best.targetAmount, point: best });
      }

      if (solutions.length === 0) {
        return alert(`已采样当前可重分配区间，但没有找到可稳定收敛到目标值的区间。${targetName}可达范围约为 ${formatPercent(minMetric)} - ${formatPercent(maxMetric)}。`);
      }

      const bestSolution = solutions.reduce((best, current) => (
        Math.abs(current.targetAmount - structure.beforeTargetInclAmount) < Math.abs(best.targetAmount - structure.beforeTargetInclAmount)
          ? current
          : best
      ), solutions[0]);

      const finalPoint = await evaluate(bestSolution.targetAmount);
      if (!finalPoint.valid || !finalPoint.payload || !finalPoint.result) {
        return alert(finalPoint.message || "最终结构反算候选无法同步到分板块金额计划，已停止写入。");
      }
      const totalCheck = roundMoney(structure.fixedOtherInclAmount + finalPoint.targetAmount + finalPoint.balancingAmount);
      if (Math.abs(totalCheck - structure.totalInclAmount) > MONEY_EPSILON) {
        return alert("结构反算结果未能保持同侧含税总金额不变，已停止写入。");
      }
      if (finalPoint.targetAmount < -MONEY_EPSILON || finalPoint.balancingAmount < -MONEY_EPSILON) {
        return alert("结构反算结果出现负金额，已停止写入。");
      }

      state.updateTaxItemsInclBatch([
        { groupId: structure.targetSubject.groupId, key: structure.targetSubject.key, incl: finalPoint.targetAmount },
        { groupId: structure.balancingSubject.groupId, key: structure.balancingSubject.key, incl: finalPoint.balancingAmount },
      ]);
      if (modelEAmountMode) {
        state.setCashflowSegments(finalPoint.nextSegments);
      }
      state.setActiveTab(structure.side === "revenue" ? "revenue" : "cost");

      setCashflowTable(finalPoint.result.cashflow);
      setMetrics(finalPoint.result);
      updateData(AI_CONTEXT_KEY.ICT_CORE, buildAiContextPayload(true, {
        metrics: finalPoint.result,
        cashflow: finalPoint.result.cashflow,
        extra: {
          ...finalPoint.payload,
          reverse_calculation: {
            mode: "locked_total_structure",
            side: structure.side,
            target_type: revTargetType,
            target_value: revTargetValue,
            target_subject_ref: selectedSubject.ref,
            target_subject_name: structure.targetDisplayName,
            target_before_amount: structure.beforeTargetInclAmount,
            target_after_amount: finalPoint.targetAmount,
            balancing_subject: {
              subjectCode: structure.balancingSubject.subjectCode,
              groupId: structure.balancingSubject.groupId,
              key: structure.balancingSubject.key,
            },
            balancing_subject_name: structure.balancingDisplayName,
            balancing_before_amount: structure.beforeBalancingInclAmount,
            balancing_after_amount: finalPoint.balancingAmount,
            total_incl_amount: structure.totalInclAmount,
            metric_after: finalPoint.metricValue,
            model_e_amount_mode: modelEAmountMode,
            model_e_transfers: modelEAmountMode
              ? finalPoint.modelETransfers.map(transfer => ({
                  side: transfer.bucket.side,
                  scope: transfer.bucket.scope,
                  delta_incl: transfer.deltaIncl,
                  reason: transfer.reason,
                }))
              : [],
          },
        },
      }));

      const modelESuccessText = modelEAmountMode
        ? "\n本次结构调整已同步更新分板块现金流金额计划。"
        : "";

      alert(
        modelESuccessText +
        `结构反算完成：${structure.sideLabel}含税总金额保持 ${formatCurrency(structure.totalInclAmount)} 不变。\n` +
        `${targetName}：目标 ${formatPercent(target)}，当前 ${formatPercent(finalPoint.metricValue)}\n` +
        `目标科目：“${structure.targetDisplayName}” ${formatCurrency(structure.beforeTargetInclAmount)} -> ${formatCurrency(finalPoint.targetAmount)}\n` +
        `承接科目：“${structure.balancingDisplayName}” ${formatCurrency(structure.beforeBalancingInclAmount)} -> ${formatCurrency(finalPoint.balancingAmount)}\n` +
        `固定同侧科目合计：${formatCurrency(structure.fixedOtherInclAmount)}；可重分配池：${formatCurrency(structure.reallocatablePoolInclAmount)}`
      );
    } catch (e) {
      alert("结构反算失败: " + e);
    }
  };

  const performReverseCalculation = async (
    selectedSubject: ReverseSubjectOption | null,
    reverseContext?: ResolvedReverseCalculationContext,
  ) => {
    const target = Number(revTargetValue);
    if (!Number.isFinite(target)) return alert("请输入有效的目标值。");
    if (!selectedSubject) return alert("请选择需要反算的计费科目。");

    if (reverseContext?.mode === "blocked") return alert(reverseContext.message);
    if (reverseContext?.mode === "locked_total_structure") {
      return performLockedTotalStructureReverseCalculation(selectedSubject, reverseContext.structure, target);
    }

    const modelEAmountMode = state.cashflowModel === 'model_e' && state.segmentValueMode === "amount";
    const segmentIndex = modelEAmountMode
      ? selectReverseSegmentIndex(state.cashflowSegments, selectedSubject.subject.side)
      : -1;
    if (modelEAmountMode && segmentIndex < 0) {
      return alert("请先在分板块资金计划中至少保留一个板块，再使用智能反算。");
    }

    const buildCandidate = (amount: number) =>
      buildReverseCandidate(selectedSubject, amount, modelEAmountMode ? segmentIndex : null);

    const evaluate = async (amount: number) => {
      const candidate = buildCandidate(amount);
      const result: any = await invoke('calculate_ict_benefit', { input: candidate.payload });
      const metricValue = Number(revTargetType === "margin" ? result.margin_rate : result.npv_rate);
      return {
        ...candidate,
        result,
        metricValue: Number.isFinite(metricValue) ? metricValue : 0,
      };
    };

    try {
      if (state.ignoredDataHash !== null) {
        state.setIgnoredDataHash(null);
        state.setIgnoredTailValue(null);
      }

      const zeroPoint = await evaluate(0);
      if (revMode === "cost" && zeroPoint.metricValue < target) {
        return alert(`当前收入和其他成本条件下，即使“${selectedSubject.displayName}”为 0，也无法达到目标值。`);
      }

      let low = 0;
      let high = 10_000_000_000;

      if (revMode === "revenue") {
        const highPoint = await evaluate(high);
        if (highPoint.metricValue < target) {
          return alert(`当前成本和现金流条件下，即使“${selectedSubject.displayName}”达到反算上限，也无法达到目标值。`);
        }
      }

      for (let i = 0; i < 70; i++) {
        const mid = (low + high) / 2;
        const point = await evaluate(mid);

        if (revMode === "revenue") {
          if (point.metricValue < target) low = mid;
          else high = mid;
        } else if (point.metricValue > target) {
          low = mid;
        } else {
          high = mid;
        }
      }

      const finalAmount = Number((revMode === "revenue" ? high : low).toFixed(2));
      const beforeAmount = readSubjectInclAmount(getCurrentReverseSubjectState(), selectedSubject.ref);
      const finalCandidate = buildCandidate(finalAmount);
      const refreshed: any = await invoke('calculate_ict_benefit', { input: finalCandidate.payload });

      if (modelEAmountMode) {
        state.setCashflowSegments(finalCandidate.nextSegments);
      }
      state.updateTaxItem(
        selectedSubject.subject.groupId,
        selectedSubject.subject.key,
        "incl",
        finalAmount,
      );

      if (revMode === "revenue") {
        state.setActiveTab("revenue");
      } else {
        if (selectedSubject.subject.subjectCode === "cost_it_integration") {
          handleSelFeeChange('limit', String(finalAmount));
        }
        state.setActiveTab("cost");
      }

      setCashflowTable(refreshed.cashflow);
      setMetrics(refreshed);
      updateData(AI_CONTEXT_KEY.ICT_CORE, buildAiContextPayload(true, {
        metrics: refreshed,
        cashflow: refreshed.cashflow,
        extra: {
          ...finalCandidate.payload,
          reverse_calculation: {
            mode: revMode,
            target_type: revTargetType,
            target_value: revTargetValue,
            result: finalAmount,
            before_amount: beforeAmount,
            subject_ref: selectedSubject.ref,
            subject_name: selectedSubject.displayName,
            cashflow_segment: modelEAmountMode ? finalCandidate.nextSegments[segmentIndex]?.name : null,
            model_e_amount_mode: modelEAmountMode,
          },
        },
      }));

      const distText = modelEAmountMode
        ? (() => {
            const finalDirectCashflow = buildDirectCashflowFromSegments(finalCandidate.nextSegments);
            return revMode === 'revenue'
              ? formatDistribution(distributionFromCashflow(finalDirectCashflow.rev))
              : formatDistribution(distributionFromCashflow(finalDirectCashflow.cost));
          })()
        : (revMode === 'revenue' ? formatDistribution(effectiveDistRev) : formatDistribution(effectiveDistCost));
      const targetName = revTargetType === 'margin' ? '目标毛利润率' : '目标净现值率';
      const segmentText = modelEAmountMode
        ? `\n同步板块：${finalCandidate.nextSegments[segmentIndex]?.name ?? "对应板块"}`
        : "";

      alert(
        `反算完成：${formatCurrency(finalAmount)}\n` +
        `目标：${targetName} ≥ ${formatPercent(target)}\n` +
        `反算科目：“${selectedSubject.displayName}”\n` +
        `反算前金额：${formatCurrency(beforeAmount)}\n` +
        `该结果为该科目的含税金额，已按当前资金收付模型重新生成现金流。${segmentText}\n` +
        `当前资金收付模型：${cashflowModelLabels[state.cashflowModel]}\n` +
        `年度分布：${distText}`
      );
    } catch (e) {
      alert("反推失败: " + e);
    }
  };

  return {
    cashflowTable,
    metrics,
    selQuote, setSelQuote,
    selMarkup, setSelMarkup,
    selActualCost,
    selFee,
    selLimit, setSelLimit,
    revMode, setRevMode,
    revTargetType, setRevTargetType,
    revTargetValue, setRevTargetValue,
    revSubjectRefKey, setRevSubjectRefKey,
    directSegmentCashflow,
    revenueInclTotal,
    costInclTotal,
    segmentRevenueInclTotal,
    segmentCostInclTotal,
    subjectFundingCoverage,
    subjectFundingAnnualCashflow,
    subjectFundingCalculationBlocked,
    performCalculation,
    handleSelFeeChange,
    applySelectionLimit,
    performReverseCalculation,
    buildInputDataPayload,
  };
}
