import { useState, useEffect, useRef, useCallback } from "react";
import { invoke } from "@tauri-apps/api/core";
import { useAiContextStore } from "../store/useAiContextStore";
import { AI_CONTEXT_KEY } from "../utils/aiContextKeys";
import {
  type CashflowSegment,
  buildDirectCashflowFromSegments,
  distributionFromCashflow,
  cashflowPayloadValues,
  sumInclTaxItems,
  makeTaxItemFromIncl,
  clampCashflowYear,
  useIctState
} from "./useIctState";

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

  // Helper selectors
  const directSegmentCashflow = state.cashflowSegments ? buildDirectCashflowFromSegments(state.cashflowSegments) : { rev: [], cost: [], itRev: [], itCost: [] };
  const revenueInclTotal = sumInclTaxItems([...Object.values(state.revIt), ...Object.values(state.revCt), state.revNonItCt]);
  const costInclTotal = sumInclTaxItems([...Object.values(state.costIt), ...Object.values(state.costCt), ...Object.values(state.costMix)]);
  const segmentRevenueInclTotal = state.cashflowSegments.reduce((sum, segment) => sum + (segment.revenueValue || 0), 0);
  const segmentCostInclTotal = state.cashflowSegments.reduce((sum, segment) => sum + (segment.costValue || 0), 0);

  const effectiveDistRev = state.distRev;
  const effectiveDistCost = state.distCost;

  const buildInputDataPayload = (options?: { segments?: CashflowSegment[]; revItState?: typeof state.revIt; costItState?: typeof state.costIt }) => {
    const segmentsForPayload = options?.segments ?? state.cashflowSegments;
    const directCashflowForPayload = buildDirectCashflowFromSegments(segmentsForPayload);
    const revDistributionForPayload = state.cashflowModel === 'model_e' && state.segmentValueMode === "amount"
      ? distributionFromCashflow(directCashflowForPayload.rev)
      : effectiveDistRev;
    const costDistributionForPayload = state.cashflowModel === 'model_e' && state.segmentValueMode === "amount"
      ? distributionFromCashflow(directCashflowForPayload.cost)
      : effectiveDistCost;
    const revItForPayload = options?.revItState ?? state.revIt;
    const costItForPayload = options?.costItState ?? state.costIt;

    return {
      project_name: state.projName,
      customer_name: state.customerName,
      property_rights: state.propertyRights,
      discount_rate: String(state.discountRate),
      project_years: state.projectYears,
      cashflow_model: state.cashflowModel,
      rev_distribution: revDistributionForPayload,
      cost_distribution: costDistributionForPayload,
      cashflow_segment_value_mode: state.segmentValueMode,
      cashflow_segments: state.cashflowModel === 'model_e' ? segmentsForPayload : [],
      rev_cashflow_excl: state.cashflowModel === 'model_e' && state.segmentValueMode === "amount" ? cashflowPayloadValues(directCashflowForPayload.rev) : null,
      cost_cashflow_excl: state.cashflowModel === 'model_e' && state.segmentValueMode === "amount" ? cashflowPayloadValues(directCashflowForPayload.cost) : null,
      it_rev_cashflow_excl: state.cashflowModel === 'model_e' && state.segmentValueMode === "amount" ? cashflowPayloadValues(directCashflowForPayload.itRev) : null,
      it_cost_cashflow_excl: state.cashflowModel === 'model_e' && state.segmentValueMode === "amount" ? cashflowPayloadValues(directCashflowForPayload.itCost) : null,
      ignore_tail_difference: state.ignoredTailValue !== null,
      tail_difference_value: state.ignoredTailValue || "0",
      rev_it_integration: { incl_tax: String(revItForPayload.integration.incl), tax_rate: String(revItForPayload.integration.tax) },
      rev_it_maintenance: { incl_tax: String(revItForPayload.maintenance.incl), tax_rate: String(revItForPayload.maintenance.tax) },
      rev_it_device_sales: { incl_tax: String(revItForPayload.device_sales.incl), tax_rate: String(revItForPayload.device_sales.tax) },
      rev_it_device_lease: { incl_tax: String(revItForPayload.device_lease.incl), tax_rate: String(revItForPayload.device_lease.tax) },
      rev_it_other: { incl_tax: String(revItForPayload.other.incl), tax_rate: String(revItForPayload.other.tax) },
      rev_it_cloud: { incl_tax: String(revItForPayload.cloud.incl), tax_rate: String(revItForPayload.cloud.tax) },
      rev_ct_line: { incl_tax: String(state.revCt.line.incl), tax_rate: String(state.revCt.line.tax) },
      rev_ct_product: { incl_tax: String(state.revCt.product.incl), tax_rate: String(state.revCt.product.tax) },
      rev_non_it_ct: { incl_tax: String(state.revNonItCt.incl), tax_rate: String(state.revNonItCt.tax) },
      cost_it_device: { incl_tax: String(costItForPayload.device.incl), tax_rate: String(costItForPayload.device.tax) },
      cost_it_construction: { incl_tax: String(costItForPayload.construction.incl), tax_rate: String(costItForPayload.construction.tax) },
      cost_it_survey: { incl_tax: String(costItForPayload.survey.incl), tax_rate: String(costItForPayload.survey.tax) },
      cost_it_integration: { incl_tax: String(costItForPayload.integration.incl), tax_rate: String(costItForPayload.integration.tax) },
      cost_it_other: { incl_tax: String(costItForPayload.other.incl), tax_rate: String(costItForPayload.other.tax) },
      cost_it_maintenance: { incl_tax: String(costItForPayload.maintenance.incl), tax_rate: String(costItForPayload.maintenance.tax) },
      cost_it_running: { incl_tax: String(costItForPayload.running.incl), tax_rate: String(costItForPayload.running.tax) },
      cost_it_bidding: { incl_tax: String(costItForPayload.bidding.incl), tax_rate: String(costItForPayload.bidding.tax) },
      cost_it_design_eval: { incl_tax: String(costItForPayload.design_eval.incl), tax_rate: String(costItForPayload.design_eval.tax) },
      cost_it_audit: { incl_tax: String(costItForPayload.audit.incl), tax_rate: String(costItForPayload.audit.tax) },
      cost_ct_construction: { incl_tax: String(state.costCt.construction.incl), tax_rate: String(state.costCt.construction.tax) },
      cost_ct_maintenance: { incl_tax: String(state.costCt.maintenance.incl), tax_rate: String(state.costCt.maintenance.tax) },
      cost_ct_other: { incl_tax: String(state.costCt.other.incl), tax_rate: String(state.costCt.other.tax) },
      cost_ct_bandwidth: { incl_tax: String(state.costCt.bandwidth.incl), tax_rate: String(state.costCt.bandwidth.tax) },
      cost_ct_renewal: { incl_tax: String(state.costCt.renewal.incl), tax_rate: String(state.costCt.renewal.tax) },
      cost_non_it_ct: { incl_tax: String(state.costMix.non_it_ct.incl), tax_rate: String(state.costMix.non_it_ct.tax) },
      cost_mix_marketing: { incl_tax: String(state.costMix.marketing.incl), tax_rate: String(state.costMix.marketing.tax) },
      cost_mix_channel: { incl_tax: String(state.costMix.channel.incl), tax_rate: String(state.costMix.channel.tax) },
      cost_mix_other: { incl_tax: String(state.costMix.other.incl), tax_rate: String(state.costMix.other.tax) },
    };
  };

  const getInputDataPayload = () => buildInputDataPayload();

  const performCalculation = useCallback(async () => {
    try {
      const res: any = await invoke('calculate_ict_benefit', { input: getInputDataPayload() });
      if (res) {
        setCashflowTable(res.cashflow);
        setMetrics(res);
      }
    } catch (e) {
      console.error(e);
    }
  }, [state]);

  // Recalculate whenever state variables change
  useEffect(() => {
    performCalculation();
  }, [
    state.revIt, state.revCt, state.revNonItCt,
    state.costIt, state.costCt, state.costMix,
    state.projectYears, state.discountRate, state.cashflowModel,
    state.distRev, state.distCost, state.segmentValueMode, state.cashflowSegments,
    performCalculation
  ]);

  // --- AI Context Sync ---
  const buildAiContextPayload = useCallback((includeCalculated = false, overrides?: { metrics?: any; cashflow?: any[]; extra?: Record<string, any> }) => ({
    monetary_unit: '元',
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
    state.ignoredTailValue, updateData, buildAiContextPayload
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
    state.ignoredTailValue, updateData, buildAiContextPayload
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

  const performModelEAmountReverseCalculation = async () => {
    const target = Number(revTargetValue);
    if (!Number.isFinite(target)) return alert("请输入有效的目标值！");

    const side = revMode === "revenue" ? "revenue" : "cost";
    const segmentIndex = selectReverseSegmentIndex(state.cashflowSegments, side);
    if (segmentIndex < 0) {
      return alert("请先在分板块资金计划中至少保留一个板块，再使用智能反算。");
    }

    const tax = revMode === "revenue" ? state.revIt.integration.tax : state.costIt.integration.tax;
    const buildCandidate = (amount: number) => {
      const nextSegments = applyReverseValueToSegments(state.cashflowSegments, side, segmentIndex, amount, tax);
      const nextRevIt = revMode === "revenue"
        ? { ...state.revIt, integration: makeTaxItemFromIncl(amount, tax) }
        : state.revIt;
      const nextCostIt = revMode === "cost"
        ? { ...state.costIt, integration: makeTaxItemFromIncl(amount, tax) }
        : state.costIt;
      const payload = buildInputDataPayload({
        segments: nextSegments,
        revItState: nextRevIt,
        costItState: nextCostIt,
      });

      return { nextSegments, nextRevIt, nextCostIt, payload };
    };

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
        return alert("当前收入和其他成本条件下，即使系统集成服务成本为 0，也无法达到目标值。");
      }

      let low = 0;
      let high = 10_000_000_000;
      let bestAmount = 0;

      if (revMode === "revenue") {
        const highPoint = await evaluate(high);
        if (highPoint.metricValue < target) {
          return alert("当前成本和现金流条件下，即使系统集成服务收入达到上限，也无法达到目标值。");
        }
      }

      for (let i = 0; i < 70; i++) {
        const mid = (low + high) / 2;
        const point = await evaluate(mid);
        bestAmount = mid;

        if (revMode === "revenue") {
          if (point.metricValue < target) {
            low = mid;
          } else {
            high = mid;
          }
        } else if (point.metricValue > target) {
          low = mid;
        } else {
          high = mid;
        }
      }

      const finalAmount = Number(bestAmount.toFixed(2));
      const finalCandidate = buildCandidate(finalAmount);
      const refreshed: any = await invoke('calculate_ict_benefit', { input: finalCandidate.payload });

      state.setCashflowSegments(finalCandidate.nextSegments);
      if (revMode === "revenue") {
        state.setRevIt(finalCandidate.nextRevIt);
      } else {
        state.setCostIt(finalCandidate.nextCostIt);
        handleSelFeeChange('limit', String(finalAmount));
      }

      setCashflowTable(refreshed.cashflow);
      setMetrics(refreshed);
      state.setActiveTab("basic");
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
            cashflow_segment: finalCandidate.nextSegments[segmentIndex]?.name,
            model_e_amount_mode: true,
          },
        },
      }));

      const finalDirectCashflow = buildDirectCashflowFromSegments(finalCandidate.nextSegments);
      const distText = revMode === 'revenue'
        ? formatDistribution(distributionFromCashflow(finalDirectCashflow.rev))
        : formatDistribution(distributionFromCashflow(finalDirectCashflow.cost));
      const targetName = revTargetType === 'margin' ? '毛利润率' : '净现值率';
      const reverseFieldName = revMode === 'revenue' ? '系统集成服务收入' : '系统集成服务成本';
      const segmentName = finalCandidate.nextSegments[segmentIndex]?.name ?? "对应板块";

      alert(
        `反算完成：${formatCurrency(finalAmount)}\n` +
        `目标：${targetName} ≥ ${formatPercent(target)}\n` +
        `反算字段：${reverseFieldName}\n` +
        `同步板块：${segmentName}\n` +
        `当前资金收付模型：${cashflowModelLabels[state.cashflowModel]}\n` +
        `年度分布：${distText}\n` +
        `已按该板块金额、税率和年度计划重新生成精确现金流。`
      );
    } catch (e) {
      alert("反推失败: " + e);
    }
  };

  const performReverseCalculation = async () => {
    if (!revTargetValue) return alert("请输入目标值！");
    if (state.cashflowModel === 'model_e' && state.segmentValueMode === "amount") {
      return performModelEAmountReverseCalculation();
    }
    try {
      const apiName = revMode === 'revenue' ? 'reverse_calc_ict_revenue_target' : 'reverse_calc_ict_target';
      const basePayload = getInputDataPayload();
      const valStr: string = await invoke(apiName, {
        input: basePayload,
        targetType: revTargetType,
        targetValue: String(revTargetValue)
      });

      const numVal = Number(valStr);
      const nextPayload = {
        ...basePayload,
        ...(revMode === 'revenue'
          ? { rev_it_integration: { ...basePayload.rev_it_integration, incl_tax: String(numVal) } }
          : { cost_it_integration: { ...basePayload.cost_it_integration, incl_tax: String(numVal) } })
      };

      if (revMode === 'revenue') {
        state.updateTaxItem('revIt', 'integration', 'incl', numVal);
        state.setActiveTab("revenue");
      } else {
        state.updateTaxItem('costIt', 'integration', 'incl', numVal);
        handleSelFeeChange('limit', String(numVal));
        state.setActiveTab("cost");
      }

      const refreshed: any = await invoke('calculate_ict_benefit', { input: nextPayload });
      if (refreshed) {
        setCashflowTable(refreshed.cashflow);
        setMetrics(refreshed);
        updateData(AI_CONTEXT_KEY.ICT_CORE, buildAiContextPayload(true, {
          metrics: refreshed,
          cashflow: refreshed.cashflow,
          extra: {
            ...nextPayload,
            reverse_calculation: {
              mode: revMode,
              target_type: revTargetType,
              target_value: revTargetValue,
              result: numVal,
            },
          },
        }));
      }

      const distText = revMode === 'revenue'
        ? formatDistribution(effectiveDistRev)
        : formatDistribution(effectiveDistCost);
      const targetName = revTargetType === 'margin' ? '毛利润率' : '净现值率';
      const reverseFieldName = revMode === 'revenue' ? '系统集成服务收入' : '系统集成服务成本';

      alert(
        `反算完成：${formatCurrency(numVal)}\n` +
        `目标：${targetName} ≥ ${formatPercent(Number(revTargetValue))}\n` +
        `反算字段：${reverseFieldName}\n` +
        `该结果为含税总额参数值，将按当前资金收付模型自动分摊。\n` +
        `当前资金收付模型：${cashflowModelLabels[state.cashflowModel]}\n` +
        `年度分布：${distText}\n` +
        `已自动刷新 10 年现金流推演。`
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
    directSegmentCashflow,
    revenueInclTotal,
    costInclTotal,
    segmentRevenueInclTotal,
    segmentCostInclTotal,
    performCalculation,
    handleSelFeeChange,
    applySelectionLimit,
    performReverseCalculation,
    buildInputDataPayload,
  };
}
