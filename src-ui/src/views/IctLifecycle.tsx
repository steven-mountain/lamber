import { useEffect, useState, useCallback, useMemo } from "react";
import WorkspaceHeader from "../components/WorkspaceHeader";
import { useRef } from "react";
import AppIcon from "../components/icons/AppIcon";
import TemplateForms from "./TemplateForms";
import { useNavigationStore } from "../store/useNavigationStore";
import { validateFinancialData, type ValidationReport } from "../lib/financeValidator";
import { SCHEME_STAGE_OPTIONS, getSchemeStageLabel, type SchemeStage } from "../lib/schemeStage";
import { useAiContextStore } from "../store/useAiContextStore";
import { AI_CONTEXT_KEY, buildAiContextKey } from "../utils/aiContextKeys";
import { useIctState } from "../hooks/useIctState";
import { useIctCalculations } from "../hooks/useIctCalculations";
import { IctBasicInfo } from "../components/IctBasicInfo";
import { IctCashflowTable } from "../components/IctCashflowTable";
import { IctProposalCashflowSummary } from "../components/IctProposalCashflowSummary";
import { IctMetricsDashboard } from "../components/IctMetricsDashboard";
import {
  projectService,
  type Project,
  type BenefitAnalysisScheme,
  type BenefitAnalysisSnapshot
} from "../utils/projectService";
import { useWorkspaceStore } from "../store/useWorkspaceStore";
import { useProjectStore } from "../store/useProjectStore";
import { useSaveStore } from "../store/useSaveStore";
import { useUnsavedChangesGuard } from "../hooks/useUnsavedChangesGuard";
import { domainSaveService } from "../services/domainSaveService";
import type { LifecycleStatePayload, CashflowStatePayload } from "../services/domainSaveService";
import {
  ICT_SUBJECT_DEFINITIONS,
  ICT_SUBJECT_GROUPS,
  getSubjectBillingName,
  getSubjectCustomName,
  getSubjectExcelDisplayName,
  normalizeCustomSubjectName,
  type IctSubjectDefinition
} from "../lib/ictSubjectCatalog";
import {
  evaluateBalanceRule,
  isBalanceSubjectMatch,
  normalizeBalanceAllocationState,
  serializeBalanceAllocationRule,
  serializeBalanceAllocationState,
  type BalanceAllocationSide,
  type BalanceRuleEvaluation,
  type BalanceSubjectItem,
} from "../lib/ictBalanceAllocation";
import {
  findReverseSubjectOption,
  getReverseEligibleSubjects,
  resolveReverseCalculationContext,
  getReverseSubjectRef,
  getReverseSubjectRefKey,
} from "../lib/ictReverseCalculation";
import {
  SubjectRoleActions,
  SelectedSubjectRoleSummary,
  scrollToSubject,
} from "../components/IctSubjectRoleComponents";
import IctSubjectFundingPlanEditor from "../components/IctSubjectFundingPlanEditor";
import {
  createSubjectFundingPlanId,
  normalizeSubjectFundingPlans,
  initializeMissingSubjectFundingPlans,
  migrateLegacySubjectFundingPlans,
  SUBJECT_FUNDING_PLAN_MIGRATION_VERSION,
  type SubjectFundingSubjectRef,
} from "../lib/ictSubjectFundingPlan";
import { exclFromIncl, normalizeTaxPairFromIncl, restoreTaxSplitParts, roundMoneyHalfUp, splitInclAmount, type TaxSplitPart } from "../lib/taxAmount";
import { useCalcPreferencesStore } from "../store/useCalcPreferencesStore";

const restoreCustomSubjectName = (item: any) => normalizeCustomSubjectName(item?.customSubjectName ?? item?.custom_subject_name ?? "");
const restoreBillingSubjectName = (item: any) => normalizeCustomSubjectName(item?.billingSubjectName ?? item?.billing_subject_name ?? "");

const syncRestoredSubjectNamePair = (
  leftItem: any,
  rightItem: any,
  field: "customSubjectName" | "billingSubjectName",
  restore: (item: any) => string,
) => {
  const leftName = restore(leftItem);
  const rightName = restore(rightItem);
  if (leftName && !rightName) rightItem[field] = leftName;
  if (rightName && !leftName) leftItem[field] = rightName;
};

const syncRestoredPairedSubjectNames = (leftItem: any, rightItem: any) => {
  syncRestoredSubjectNamePair(leftItem, rightItem, "customSubjectName", restoreCustomSubjectName);
  syncRestoredSubjectNamePair(leftItem, rightItem, "billingSubjectName", restoreBillingSubjectName);
};

const taxItemFinancialPart = (item: any) => ({
  incl: Number(item?.incl || 0),
  tax: Number(item?.tax || 0),
  excl: Number(item?.excl || 0),
});

const readFiniteNumber = (...values: unknown[]) => {
  for (const value of values) {
    if (value === undefined || value === null || value === "") continue;
    const numeric = Number(value);
    if (Number.isFinite(numeric)) return numeric;
  }
  return null;
};

const restoreTaxItemNumber = (item: any, snakeKey: string, camelKey: string, fallback = 0) => {
  if (!item) return fallback;
  const numeric = readFiniteNumber(item[snakeKey], item[camelKey]);
  return numeric === null ? fallback : numeric;
};

const buildFinancialStateHash = (data: any) => JSON.stringify({
  revIt: Object.fromEntries(Object.entries(data.revIt || {}).map(([key, item]) => [key, taxItemFinancialPart(item)])),
  revCt: Object.fromEntries(Object.entries(data.revCt || {}).map(([key, item]) => [key, taxItemFinancialPart(item)])),
  revNonItCt: taxItemFinancialPart(data.revNonItCt),
  costIt: Object.fromEntries(Object.entries(data.costIt || {}).map(([key, item]) => [key, taxItemFinancialPart(item)])),
  costCt: Object.fromEntries(Object.entries(data.costCt || {}).map(([key, item]) => [key, taxItemFinancialPart(item)])),
  costMix: Object.fromEntries(Object.entries(data.costMix || {}).map(([key, item]) => [key, taxItemFinancialPart(item)])),
});

const formatReverseCurrency = (value: number) =>
  new Intl.NumberFormat("zh-CN", { style: "currency", currency: "CNY" }).format(value);

const buildRestoredFundingSubjects = (sources: {
  revIt: any;
  revCt: any;
  revNonItCt: any;
  costIt: any;
  costCt: any;
  costMix: any;
}) => {
  const resolveItem = (subject: IctSubjectDefinition) => {
    if (subject.groupId === "revIt") return sources.revIt[subject.key] || null;
    if (subject.groupId === "revCt") return sources.revCt[subject.key] || null;
    if (subject.groupId === "revNonItCt") return sources.revNonItCt;
    if (subject.groupId === "costIt") return sources.costIt[subject.key] || null;
    if (subject.groupId === "costCt") return sources.costCt[subject.key] || null;
    if (subject.groupId === "costMix") return sources.costMix[subject.key] || null;
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
	      taxRate: Number(item?.tax ?? subject.defaultTaxRate),
	      isItScope: subject.groupId === "revIt" || subject.groupId === "costIt",
	    };
  });
};

export default function IctLifecycle() {
  const { activeProjectId, activeSchemeId, activeScenarioId, entrySource, ictOrigin, navigateTo } = useNavigationStore();
  const isWorkspaceReady = useWorkspaceStore(state => state.isWorkspaceReady);
  const workspaceId = useWorkspaceStore(state => state.workspaceId);
  const state = useIctState();
  const calculations = useIctCalculations(state);
  const taxInclAutoFix = useCalcPreferencesStore(s => s.taxInclAutoFix);
  const restoreSelectionFeeState = calculations.restoreSelectionFeeState;
  const markDirty = useSaveStore(saveState => saveState.markDirty);
  const clearDirty = useSaveStore(saveState => saveState.clearDirty);
  const registerSaveHandler = useSaveStore(saveState => saveState.registerSaveHandler);
  const unregisterSaveHandler = useSaveStore(saveState => saveState.unregisterSaveHandler);
  const dirtyScopes = useSaveStore(saveState => saveState.dirtyScopes);
  const isSaving = useSaveStore(saveState => saveState.isSaving);
  const lastSaveError = useSaveStore(saveState => saveState.lastSaveError);
  const { confirmOrSave } = useUnsavedChangesGuard();
  const isHydratingRef = useRef(false);
  const projectLoadRequestRef = useRef(0);

  const [projects, setProjects] = useState<Project[]>([]);
  const [activeProject, setActiveProject] = useState<Project | null>(null);
  const [schemes, setSchemes] = useState<BenefitAnalysisScheme[]>([]);
  const [activeScheme, setActiveScheme] = useState<BenefitAnalysisScheme | null>(null);
  const [activeSnapshot, setActiveSnapshot] = useState<BenefitAnalysisSnapshot | null>(null);
  const [pendingNewSchemeName, setPendingNewSchemeName] = useState<string | null>(null);

  const [showSaveAsModal, setShowSaveAsModal] = useState(false);
  const [saveAsSchemeName, setSaveAsSchemeName] = useState("");
  // 另存为新方案时要打的甄选阶段标签（派生"甄选后"方案时使用），普通另存为为 null。
  const [saveAsSchemeStage, setSaveAsSchemeStage] = useState<SchemeStage | null>(null);
  const [showSelectProjectModal, setShowSelectProjectModal] = useState(false);
  const [fundingPlanFocus, setFundingPlanFocus] = useState<{ planId: string; token: number } | null>(null);
  const [cashflowSourceExpanded, setCashflowSourceExpanded] = useState(false);

  const buildHydrationInput = useCallback((baseInput: any, cashflowState: any) => {
    const merged = { ...(baseInput || {}) };
    const assumptions = cashflowState?.assumptionsJson || {};
    const paymentModel = cashflowState?.paymentModelJson || {};
    const sectorCashflow = cashflowState?.sectorCashflowJson || {};

    if (cashflowState?.cashflowModel || paymentModel.cashflowModel) {
      merged.cashflow_model = cashflowState?.cashflowModel || paymentModel.cashflowModel;
    }
    if (paymentModel.revDistribution) merged.rev_distribution = paymentModel.revDistribution;
    if (paymentModel.costDistribution) merged.cost_distribution = paymentModel.costDistribution;
    if (paymentModel.segmentValueMode) merged.cashflow_segment_value_mode = paymentModel.segmentValueMode;
    if (sectorCashflow.cashflowSegments) merged.cashflow_segments = sectorCashflow.cashflowSegments;
    if (assumptions.projectYears !== undefined) merged.project_years = assumptions.projectYears;
    if (assumptions.discountRate !== undefined) merged.discount_rate = String(assumptions.discountRate);
    if (assumptions.balanceAllocation || assumptions.balance_allocation) {
      const restoredBalanceAllocation = normalizeBalanceAllocationState(
        assumptions.balanceAllocation || assumptions.balance_allocation,
      );
      merged.revenue_balance_rule = serializeBalanceAllocationRule(restoredBalanceAllocation.revenue);
      merged.investment_balance_rule = serializeBalanceAllocationRule(restoredBalanceAllocation.investment);
    }
    if (assumptions.subjectFundingPlans || assumptions.subject_funding_plans) {
      merged.subject_funding_plans = assumptions.subjectFundingPlans || assumptions.subject_funding_plans;
    }
    const restoredMigrationVersion = readFiniteNumber(
      assumptions.subjectFundingPlanMigrationVersion,
      assumptions.subject_funding_plan_migration_version,
      merged.subjectFundingPlanMigrationVersion,
      merged.subject_funding_plan_migration_version,
    );
    if (restoredMigrationVersion !== null) {
      merged.subject_funding_plan_migration_version = restoredMigrationVersion;
    }

    const applyTaxItem = (key: string, item: any) => {
      if (!item) return;
      const customSubjectName = restoreCustomSubjectName(item);
      const billingSubjectName = restoreBillingSubjectName(item);
      merged[key] = {
        incl_tax: String(item.incl ?? item.incl_tax ?? 0),
        tax_rate: String(item.tax ?? item.tax_rate ?? 0),
        ...(customSubjectName ? { custom_subject_name: customSubjectName } : {}),
        ...(billingSubjectName ? { billing_subject_name: billingSubjectName } : {}),
      };
    };

    applyTaxItem("rev_it_integration", assumptions.revIt?.integration);
    applyTaxItem("rev_it_maintenance", assumptions.revIt?.maintenance);
    applyTaxItem("rev_it_device_sales", assumptions.revIt?.device_sales);
    applyTaxItem("rev_it_device_lease", assumptions.revIt?.device_lease);
    applyTaxItem("rev_it_other", assumptions.revIt?.other);
    applyTaxItem("rev_it_cloud", assumptions.revIt?.cloud);
    applyTaxItem("rev_ct_line", assumptions.revCt?.line);
    applyTaxItem("rev_ct_product", assumptions.revCt?.product);
    applyTaxItem("rev_non_it_ct", assumptions.revNonItCt);
    applyTaxItem("cost_it_device", assumptions.costIt?.device);
    applyTaxItem("cost_it_construction", assumptions.costIt?.construction);
    applyTaxItem("cost_it_survey", assumptions.costIt?.survey);
    applyTaxItem("cost_it_integration", assumptions.costIt?.integration);
    applyTaxItem("cost_it_other", assumptions.costIt?.other);
    applyTaxItem("cost_it_maintenance", assumptions.costIt?.maintenance);
    applyTaxItem("cost_it_running", assumptions.costIt?.running);
    applyTaxItem("cost_it_bidding", assumptions.costIt?.bidding);
    applyTaxItem("cost_it_design_eval", assumptions.costIt?.design_eval);
    applyTaxItem("cost_it_audit", assumptions.costIt?.audit);
    applyTaxItem("cost_ct_construction", assumptions.costCt?.construction);
    applyTaxItem("cost_ct_maintenance", assumptions.costCt?.maintenance);
    applyTaxItem("cost_ct_other", assumptions.costCt?.other);
    applyTaxItem("cost_ct_bandwidth", assumptions.costCt?.bandwidth);
    applyTaxItem("cost_ct_renewal", assumptions.costCt?.renewal);
    applyTaxItem("cost_non_it_ct", assumptions.costMix?.non_it_ct);
    applyTaxItem("cost_mix_marketing", assumptions.costMix?.marketing);
    applyTaxItem("cost_mix_channel", assumptions.costMix?.channel);
    applyTaxItem("cost_mix_other", assumptions.costMix?.other);

    return merged;
  }, []);

  useEffect(() => {
    if (!isWorkspaceReady) {
      setProjects([]);
      return;
    }
    projectService.getProjects().then(setProjects).catch(console.error);
  }, [activeProjectId, isWorkspaceReady]);


	  const fillCalculatorState = useCallback((params: any) => {
	    if (!params) return;

	    const defaultTaxRateForSubject = (subjectCode: string) =>
	      ICT_SUBJECT_DEFINITIONS.find(subject => subject.subjectCode === subjectCode)?.defaultTaxRate ?? 6;
	    const restoreItem = (item: any, defaultTax = 6) => {
	      const incl = restoreTaxItemNumber(item, "incl_tax", "incl", 0);
	      const tax = restoreTaxItemNumber(item, "tax_rate", "tax", defaultTax);
      const explicitExcl = restoreTaxItemNumber(item, "excl_tax", "excl", Number.NaN);
      const excl = Number.isFinite(explicitExcl)
        ? explicitExcl
        : exclFromIncl(incl, tax);
      // 拆分明细随存档还原；不含税以两笔之和为准（逐行闭合口径）。
      const splitParts = restoreTaxSplitParts(item?.split_parts, incl, tax);
      return {
        incl,
        tax,
        excl: splitParts
          ? roundMoneyHalfUp(splitParts.reduce((sum, part) => sum + part.excl, 0))
          : excl,
        customSubjectName: restoreCustomSubjectName(item),
	        billingSubjectName: restoreBillingSubjectName(item),
	        ...(splitParts ? { splitParts } : {}),
	      };
	    };
	    const restoreSubjectItem = (subjectCode: string, item: any) =>
	      restoreItem(item, defaultTaxRateForSubject(subjectCode));

    if (params.project_name) state.setProjName(params.project_name);
    if (params.customer_name) state.setCustomerName(params.customer_name);
    if (params.property_rights) state.setPropertyRights(params.property_rights);
    if (params.discount_rate) state.setDiscountRate(Number(params.discount_rate));
    if (params.project_years) state.setProjectYears(params.project_years);
    if (params.cashflow_model) state.setCashflowModel(params.cashflow_model as any);
    if (params.rev_distribution) state.setDistRev(params.rev_distribution);
    if (params.cost_distribution) state.setDistCost(params.cost_distribution);
    if (params.cashflow_segment_value_mode) state.setSegmentValueMode(params.cashflow_segment_value_mode);
    if (params.cashflow_segments) state.setCashflowSegments(params.cashflow_segments);
    if (params.project_background) state.setProjectBackground(params.project_background);
    state.setBalanceAllocation(normalizeBalanceAllocationState({
      revenue: params.revenue_balance_rule ?? params.revenueBalanceRule,
      investment: params.investment_balance_rule ?? params.investmentBalanceRule,
    }));

	    const revItRestored = {
	      integration: restoreSubjectItem("rev_it_integration", params.rev_it_integration),
	      maintenance: restoreSubjectItem("rev_it_maintenance", params.rev_it_maintenance),
	      device_sales: restoreSubjectItem("rev_it_device_sales", params.rev_it_device_sales),
	      device_lease: restoreSubjectItem("rev_it_device_lease", params.rev_it_device_lease),
	      other: restoreSubjectItem("rev_it_other", params.rev_it_other),
	      cloud: restoreSubjectItem("rev_it_cloud", params.rev_it_cloud),
	    };
	    const revCtRestored = {
	      line: restoreSubjectItem("rev_ct_line", params.rev_ct_line),
	      product: restoreSubjectItem("rev_ct_product", params.rev_ct_product),
	    };
	    const revNonItCtRestored = restoreSubjectItem("rev_non_it_ct", params.rev_non_it_ct);

	    const costItRestored = {
	      device: restoreSubjectItem("cost_it_device", params.cost_it_device),
	      construction: restoreSubjectItem("cost_it_construction", params.cost_it_construction),
	      survey: restoreSubjectItem("cost_it_survey", params.cost_it_survey),
	      integration: restoreSubjectItem("cost_it_integration", params.cost_it_integration),
	      other: restoreSubjectItem("cost_it_other", params.cost_it_other),
	      maintenance: restoreSubjectItem("cost_it_maintenance", params.cost_it_maintenance),
	      running: restoreSubjectItem("cost_it_running", params.cost_it_running),
	      bidding: restoreSubjectItem("cost_it_bidding", params.cost_it_bidding),
	      design_eval: restoreSubjectItem("cost_it_design_eval", params.cost_it_design_eval),
	      audit: restoreSubjectItem("cost_it_audit", params.cost_it_audit),
	    };
	    const costCtRestored = {
	      construction: restoreSubjectItem("cost_ct_construction", params.cost_ct_construction),
	      maintenance: restoreSubjectItem("cost_ct_maintenance", params.cost_ct_maintenance),
	      other: restoreSubjectItem("cost_ct_other", params.cost_ct_other),
	      bandwidth: restoreSubjectItem("cost_ct_bandwidth", params.cost_ct_bandwidth),
	      renewal: restoreSubjectItem("cost_ct_renewal", params.cost_ct_renewal),
	    };
	    const costMixRestored = {
	      non_it_ct: restoreSubjectItem("cost_non_it_ct", params.cost_non_it_ct),
	      marketing: restoreSubjectItem("cost_mix_marketing", params.cost_mix_marketing),
	      channel: restoreSubjectItem("cost_mix_channel", params.cost_mix_channel),
	      other: restoreSubjectItem("cost_mix_other", params.cost_mix_other),
	    };

    syncRestoredPairedSubjectNames(revCtRestored.product, costCtRestored.other);
    syncRestoredPairedSubjectNames(revCtRestored.line, costCtRestored.bandwidth);

    state.setRevIt(revItRestored);
    state.setRevCt(revCtRestored);
    state.setRevNonItCt(revNonItCtRestored);
    state.setCostIt(costItRestored);
    state.setCostCt(costCtRestored);
    state.setCostMix(costMixRestored);
    state.setCashflowCalculationSource("subject_funding_plans");
    const migrationVersion = readFiniteNumber(
      params.subject_funding_plan_migration_version,
      params.subjectFundingPlanMigrationVersion,
    );
    const fundingMigration = migrateLegacySubjectFundingPlans(
      buildRestoredFundingSubjects({
        revIt: revItRestored,
        revCt: revCtRestored,
        revNonItCt: revNonItCtRestored,
        costIt: costItRestored,
        costCt: costCtRestored,
        costMix: costMixRestored,
      }),
      normalizeSubjectFundingPlans(params.subject_funding_plans ?? params.subjectFundingPlans),
      migrationVersion,
    );
    state.setSubjectFundingPlans(fundingMigration.plans);
    state.setSubjectFundingPlanMigrationVersion(
      fundingMigration.completed
        ? SUBJECT_FUNDING_PLAN_MIGRATION_VERSION
        : migrationVersion ?? undefined,
    );

    if (params.ignore_tail_difference) {
      state.setIgnoredTailValue(params.tail_difference_value || "0");
      state.setIgnoredDataHash(buildFinancialStateHash({
        revIt: revItRestored,
        revCt: revCtRestored,
        revNonItCt: revNonItCtRestored,
        costIt: costItRestored,
        costCt: costCtRestored,
        costMix: costMixRestored
      }));
    } else {
      state.setIgnoredTailValue(null);
      state.setIgnoredDataHash(null);
    }
    restoreSelectionFeeState(params);
  }, [state, restoreSelectionFeeState]);

  const loadProjectContext = useCallback(async (pId: string | null, sId?: string | null) => {
    const requestId = ++projectLoadRequestRef.current;
    const targetProjectId = pId || null;
    const targetSchemeId = sId || null;

    if (!targetProjectId) {
      localStorage.removeItem("lamber_active_project_id");
      localStorage.removeItem("lamber_active_scheme_id");
      setActiveProject(null);
      useProjectStore.getState().setCurrentProject(null);
      setSchemes([]);
      setActiveScheme(null);
      setActiveSnapshot(null);
      setPendingNewSchemeName(null);
      state.resetProjectState();
      restoreSelectionFeeState(null);
      return;
    }
    if (!isWorkspaceReady) {
      return;
    }

    // 同一项目内切换方案时，不清空项目级/方案级展示状态：否则 activeProject 会先变 null，
    // 顶部项目卡片（红框）瞬间切成"自由测算模式"布局、数据回来后再切回来，造成一次跳动。
    // 仅在切换到不同项目时才整块清空；同项目下这些状态会在数据就绪后从旧值平滑更新到新值。
    const previousProjectId = useProjectStore.getState().currentProject?.id || null;
    const isSameProject = targetProjectId === previousProjectId;

    isHydratingRef.current = true;
    if (!isSameProject) {
      setActiveProject(null);
      setSchemes([]);
      setActiveScheme(null);
      setActiveSnapshot(null);
    }
    setPendingNewSchemeName(null);
    state.resetProjectState();
    restoreSelectionFeeState(null);

    try {
      const project = await projectService.getProject(targetProjectId);
      if (requestId !== projectLoadRequestRef.current) return;
      if (!project) {
        localStorage.removeItem("lamber_active_project_id");
        localStorage.removeItem("lamber_active_scheme_id");
        setActiveProject(null);
        useProjectStore.getState().setCurrentProject(null);
        setSchemes([]);
        setActiveScheme(null);
        setActiveSnapshot(null);
        setPendingNewSchemeName(null);
        state.resetProjectState();
        isHydratingRef.current = false;
        return;
      }

      localStorage.setItem("lamber_active_project_id", project.id);
      setActiveProject(project);
      useProjectStore.getState().setCurrentProject(project);

      const newSchemeNameLocal = localStorage.getItem("lamber_new_scheme_name");
      if (newSchemeNameLocal) {
        setPendingNewSchemeName(newSchemeNameLocal);
        localStorage.removeItem("lamber_new_scheme_name");
        setActiveScheme(null);
        setActiveSnapshot(null);

        isHydratingRef.current = true;
        state.setProjName(project.name);
        state.setCustomerName(project.customer_name);
        if (project.discount_rate > 0) state.setDiscountRate(project.discount_rate);
        if (project.project_years > 0) state.setProjectYears(project.project_years);
        if (project.cashflow_model) state.setCashflowModel(project.cashflow_model as any);
        state.setCashflowCalculationSource("subject_funding_plans");
        state.setSubjectFundingPlanMigrationVersion(SUBJECT_FUNDING_PLAN_MIGRATION_VERSION);
        state.setSubjectFundingPlans({});
        setTimeout(() => {
          isHydratingRef.current = false;
          useSaveStore.getState().clearDirtyScopes(["lifecycle", "cashflow", "benefit-analysis"]);
        }, 0);
        return;
      }

      setPendingNewSchemeName(null);

      const fullState = await domainSaveService.loadProjectFullState(project.id).catch(error => {
        console.warn("Failed to load project full state, fallback to legacy chain:", error);
        return null;
      });
      if (requestId !== projectLoadRequestRef.current) return;
      const projectSchemes = fullState?.schemes || await projectService.getSchemes(project.id);
      setSchemes(projectSchemes);
      let schemeToSelect: BenefitAnalysisScheme | null = null;
      if (targetSchemeId) {
        schemeToSelect = projectSchemes.find(s => s.id === targetSchemeId) || null;
      }
      if (!schemeToSelect) {
        schemeToSelect = projectSchemes.find(s => s.id === project.default_scheme_id) || projectSchemes[0] || null;
      }

      // 工作副本（lifecycle/cashflow state）现按 (project, scheme) 独立存储：加载选中方案时，
      // 只读取“该方案自己的草稿”，方案之间彻底隔离，不再共用项目级单例。
      // 无选中方案（自由/历史项目）才回退到项目级默认桶(fullState)与 legacy 输入。
      let draftLifecycle: any = null;
      let draftCashflow: any = null;
      if (schemeToSelect) {
        [draftLifecycle, draftCashflow] = await Promise.all([
          domainSaveService.loadLifecycleState(project.id, schemeToSelect.id).catch(() => null),
          domainSaveService.loadCashflowState(project.id, schemeToSelect.id).catch(() => null),
        ]);
        if (requestId !== projectLoadRequestRef.current) return;
      } else {
        draftLifecycle = fullState?.lifecycleState || null;
        draftCashflow = fullState?.cashflowState || null;
      }

      const currentLifecycleInput = draftLifecycle?.inputPayloadJson;
      // 选中方案时不回退到 legacy 输入：没有草稿就走该方案的快照（下方分支），避免串入项目级历史数据。
      const baseHydrationInput =
        currentLifecycleInput || (schemeToSelect ? null : fullState?.legacyLifecycleInput) || null;
      const currentHydrationInput = baseHydrationInput
        ? buildHydrationInput(baseHydrationInput, draftCashflow)
        : null;
      const preferCurrentState = Boolean(
        currentHydrationInput && (currentLifecycleInput || draftCashflow)
      );

      if (schemeToSelect) {
        setActiveScheme(schemeToSelect);
        localStorage.setItem("lamber_active_scheme_id", schemeToSelect.id);

        const snapshots = await projectService.getSnapshots(schemeToSelect.id);
        if (requestId !== projectLoadRequestRef.current) return;
        if (preferCurrentState) {
          setActiveSnapshot(snapshots[0] || null);
          isHydratingRef.current = true;
          fillCalculatorState(currentHydrationInput);
          setTimeout(() => {
            isHydratingRef.current = false;
            useSaveStore.getState().clearDirtyScopes(["lifecycle", "cashflow", "benefit-analysis"]);
          }, 0);
        } else if (snapshots.length > 0) {
          const latestSnap = snapshots.reduce((latest, current) => current.version > latest.version ? current : latest, snapshots[0]);
          setActiveSnapshot(latestSnap);
          isHydratingRef.current = true;
          fillCalculatorState(latestSnap.input_params);
          setTimeout(() => {
            isHydratingRef.current = false;
            useSaveStore.getState().clearDirtyScopes(["lifecycle", "cashflow", "benefit-analysis"]);
          }, 0);
        } else {
          setActiveSnapshot(null);
          isHydratingRef.current = true;
          state.setProjName(project.name);
          state.setCustomerName(project.customer_name);
          if (project.discount_rate > 0) state.setDiscountRate(project.discount_rate);
          if (project.project_years > 0) state.setProjectYears(project.project_years);
          if (project.cashflow_model) state.setCashflowModel(project.cashflow_model as any);
          state.setCashflowCalculationSource("subject_funding_plans");
          state.setSubjectFundingPlanMigrationVersion(SUBJECT_FUNDING_PLAN_MIGRATION_VERSION);
          state.setSubjectFundingPlans({});
          setTimeout(() => {
            isHydratingRef.current = false;
            useSaveStore.getState().clearDirtyScopes(["lifecycle", "cashflow", "benefit-analysis"]);
          }, 0);
        }
      } else {
        setActiveScheme(null);
        setActiveSnapshot(null);
        isHydratingRef.current = true;
        if (preferCurrentState) {
          fillCalculatorState(currentHydrationInput);
        } else if (fullState?.legacyLifecycleInput) {
          fillCalculatorState(fullState.legacyLifecycleInput);
        } else {
          state.setProjName(project.name);
          state.setCustomerName(project.customer_name);
          if (project.discount_rate > 0) state.setDiscountRate(project.discount_rate);
          if (project.project_years > 0) state.setProjectYears(project.project_years);
          if (project.cashflow_model) state.setCashflowModel(project.cashflow_model as any);
          state.setCashflowCalculationSource("subject_funding_plans");
          state.setSubjectFundingPlanMigrationVersion(SUBJECT_FUNDING_PLAN_MIGRATION_VERSION);
          state.setSubjectFundingPlans({});
        }
        setTimeout(() => {
          isHydratingRef.current = false;
          useSaveStore.getState().clearDirtyScopes(["lifecycle", "cashflow", "benefit-analysis"]);
        }, 0);
      }
    } catch (err) {
      if (requestId === projectLoadRequestRef.current) {
        isHydratingRef.current = false;
      }
      console.error("Failed to load project context:", err);
    }
  }, [state, restoreSelectionFeeState, fillCalculatorState, buildHydrationInput, isWorkspaceReady]);

  useEffect(() => {
    loadProjectContext(activeProjectId, activeSchemeId);
  }, [activeProjectId, activeSchemeId]);

  const getSubjectFundingBlockingMessage = useCallback((actionLabel: string) => {
    if (calculations.subjectFundingCoverage.valid) {
      return "";
    }
    const issueLines = calculations.subjectFundingCoverage.issues
      .slice(0, 8)
      .map((issue, index) => `${index + 1}. ${issue.message}`);
    const extraCount = calculations.subjectFundingCoverage.issues.length - issueLines.length;
    return [
      `当前现金流计算口径为“按科目收付款计划计算”，但覆盖校验未通过，不能${actionLabel}。`,
      "请补齐或修正科目级收付款计划后再继续。",
      ...issueLines,
      ...(extraCount > 0 ? [`另有 ${extraCount} 项问题未展开。`] : []),
    ].join("\n");
  }, [calculations.subjectFundingCoverage]);

  const buildLifecycleStatePayload = (): LifecycleStatePayload => ({
    profileJson: {
      projectName: state.projName,
      customerName: state.customerName,
      propertyRights: state.propertyRights,
    },
    parametersJson: {
      projectYears: state.projectYears,
      discountRate: state.discountRate,
      ignoreTailDifference: state.ignoredTailValue !== null,
      tailDifferenceValue: state.ignoredTailValue,
      balanceAllocation: serializeBalanceAllocationState(state.balanceAllocation),
      cashflowCalculationSource: state.cashflowCalculationSource,
      subjectFundingPlanMigrationVersion: state.subjectFundingPlanMigrationVersion,
    },
    backgroundJson: {
      projectBackground: state.projectBackground,
    },
    inputPayloadJson: calculations.buildInputDataPayload(),
  });

  const buildCashflowStatePayload = (): CashflowStatePayload => ({
    cashflowModel: state.cashflowModel,
    paymentModelJson: {
      cashflowModel: state.cashflowModel,
      revDistribution: state.distRev,
      costDistribution: state.distCost,
      segmentValueMode: state.segmentValueMode,
      cashflowCalculationSource: state.cashflowCalculationSource,
      subjectFundingPlanMigrationVersion: state.subjectFundingPlanMigrationVersion,
    },
    yearlyCashflowJson: {
      cashflowTable: calculations.cashflowTable,
      directSegmentCashflow: calculations.directSegmentCashflow,
      subjectFundingAnnualCashflow: calculations.subjectFundingAnnualCashflow,
    },
    sectorCashflowJson: {
      cashflowSegments: state.cashflowSegments,
    },
    assumptionsJson: {
      projectYears: state.projectYears,
      discountRate: state.discountRate,
      revIt: state.revIt,
      revCt: state.revCt,
      revNonItCt: state.revNonItCt,
      costIt: state.costIt,
      costCt: state.costCt,
      costMix: state.costMix,
      balanceAllocation: state.balanceAllocation,
      cashflowCalculationSource: state.cashflowCalculationSource,
      subjectFundingPlans: state.subjectFundingPlans,
      subjectFundingPlanMigrationVersion: state.subjectFundingPlanMigrationVersion,
    },
    metricsJson: calculations.metrics,
  });

  // 直接保存（保存到当前项目 / 保存当前测算）会立即重新加载项目，而 loadProjectContext
  // 优先使用 lifecycle_state / cashflow_state 进行回填。若这里只写入效益快照，
  // 重新加载时会读取到未更新的 lifecycle/cashflow 状态，导致刚刚的修改（如取消差额承接）被回退。
  // 因此这两个保存入口必须同时持久化 lifecycle 与 cashflow 状态。
  const persistLifecycleAndCashflowState = async (projectId: string, schemeId: string | null) => {
    await domainSaveService.saveLifecycleState(projectId, schemeId, buildLifecycleStatePayload());
    await domainSaveService.saveCashflowState(projectId, schemeId, buildCashflowStatePayload());
  };

  const handleSaveToSelectedProject = async (targetProjectId: string) => {
    if (!isWorkspaceReady) {
      alert("请先新建或打开工作区后再保存项目方案。");
      return;
    }
    const blockingMessage = getSubjectFundingBlockingMessage("保存测算方案");
    if (blockingMessage) {
      alert(blockingMessage);
      return;
    }

    try {
      const payload = calculations.buildInputDataPayload();
      const project = projects.find(p => p.id === targetProjectId);
      const schemeName = project ? `${project.name}_首次测算` : "测算方案";

      const updatedProj = await domainSaveService.saveBenefitAnalysis(
        targetProjectId,
        null,
        schemeName,
        payload,
        metrics,
        false
      );

      await persistLifecycleAndCashflowState(targetProjectId, updatedProj.default_scheme_id || null);

      setPendingNewSchemeName(null);
      const newSchemeId = updatedProj.default_scheme_id || null;
      clearDirty("benefit-analysis");
      clearDirty("lifecycle");
      clearDirty("cashflow");
      navigateTo("ict_lifecycle", targetProjectId, newSchemeId);
      alert("保存测算成功！已关联到项目并生成方案。");
    } catch (error) {
      console.error("保存失败:", error);
      alert("保存失败: " + error);
    }
  };

  const handleSaveToCurrent = async () => {
    if (!isWorkspaceReady) {
      alert("请先新建或打开工作区后再保存项目方案。");
      return;
    }
    if (!activeProject) return;
    const schemeId = pendingNewSchemeName ? null : (activeScheme?.id || activeProject.default_scheme_id || null);
    const schemeName = pendingNewSchemeName || activeScheme?.name || activeProject.name || "默认测算方案";
    const blockingMessage = getSubjectFundingBlockingMessage("保存当前项目指标");
    if (blockingMessage) {
      alert(blockingMessage);
      return;
    }

    try {
      const payload = calculations.buildInputDataPayload();

      const updatedProj = await domainSaveService.saveBenefitAnalysis(
        activeProject.id,
        schemeId,
        schemeName,
        payload,
        metrics,
        pendingNewSchemeName ? true : false
      );

      // 工作副本写入本次保存对应的方案桶（新建方案时即后端返回的 default_scheme_id）。
      await persistLifecycleAndCashflowState(activeProject.id, updatedProj.default_scheme_id || null);

      setPendingNewSchemeName(null);
      const newSchemeId = updatedProj.default_scheme_id || null;
      clearDirty("benefit-analysis");
      clearDirty("lifecycle");
      clearDirty("cashflow");
      await loadProjectContext(activeProject.id, newSchemeId);
      alert("保存测算成功！已更新项目指标并生成历史记录。");
    } catch (error) {
      console.error("保存失败:", error);
      alert("保存失败: " + error);
    }
  };

  const handleSaveAsNew = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isWorkspaceReady) {
      alert("请先新建或打开工作区后再保存项目方案。");
      return;
    }
    if (!activeProject || !saveAsSchemeName.trim()) return;
    const blockingMessage = getSubjectFundingBlockingMessage("另存为新方案");
    if (blockingMessage) {
      alert(blockingMessage);
      return;
    }

    try {
      const payload = calculations.buildInputDataPayload();
      const updatedProj = await domainSaveService.saveBenefitAnalysis(
        activeProject.id,
        null,
        saveAsSchemeName.trim(),
        payload,
        metrics,
        true,
        saveAsSchemeStage
      );

      // 另存为/派生的新方案，其工作副本以当前数据为起点写入新方案自己的桶（scheme_id）。
      // 每个方案的工作副本相互独立，互不覆盖。
      await persistLifecycleAndCashflowState(activeProject.id, updatedProj.default_scheme_id || null);

      setShowSaveAsModal(false);
      setSaveAsSchemeName("");
      setSaveAsSchemeStage(null);
      setPendingNewSchemeName(null);

      const newSchemeId = updatedProj.default_scheme_id || null;
      clearDirty("benefit-analysis");
      clearDirty("lifecycle");
      clearDirty("cashflow");
      await loadProjectContext(activeProject.id, newSchemeId);
      alert("另存为新方案成功！");
    } catch (error) {
      console.error("另存为失败:", error);
      alert("另存为失败: " + error);
    }
  };

  // 按更新时间倒序排列，方便在同一 stage 有多个方案时取"最近更新的一个"。
  const sortedSchemes = useMemo(
    () => [...schemes].sort((a, b) => (b.updated_at || "").localeCompare(a.updated_at || "")),
    [schemes]
  );
  // 每个阶段取最近更新的方案作为主方案。
  const preScheme = useMemo(
    () => sortedSchemes.find(s => s.stage === "pre_selection") || null,
    [sortedSchemes]
  );
  const postScheme = useMemo(
    () => sortedSchemes.find(s => s.stage === "post_selection") || null,
    [sortedSchemes]
  );
  // 未打阶段标签的方案，以及同阶段的其余方案，统一收纳进"更多方案"下拉。
  const otherSchemes = useMemo(
    () => sortedSchemes.filter(s => s.id !== preScheme?.id && s.id !== postScheme?.id),
    [sortedSchemes, preScheme, postScheme]
  );

  // 甄选结果签批表取数：不切换当前方案，直接读甄选前方案工作副本里的 IT 投入科目。
  // 甄选限价 = 甄选前方案 IT 投入；工作副本按 (project, scheme) 隔离，跨方案读取安全。
  const fetchPreSelectionCostIt = useCallback(async (): Promise<Record<string, any> | null> => {
    if (activeScheme?.stage === "pre_selection") return state.costIt;
    if (!activeProject?.id || !preScheme?.id) return null;
    try {
      const stored = await domainSaveService.loadLifecycleState(activeProject.id, preScheme.id);
      const params = stored?.inputPayloadJson;
      if (!params) return null;
      const restoreItem = (subjectCode: string, item: any) => {
        const defaultTax = ICT_SUBJECT_DEFINITIONS.find(subject => subject.subjectCode === subjectCode)?.defaultTaxRate ?? 6;
        const incl = restoreTaxItemNumber(item, "incl_tax", "incl", 0);
        const tax = restoreTaxItemNumber(item, "tax_rate", "tax", defaultTax);
        const explicitExcl = restoreTaxItemNumber(item, "excl_tax", "excl", Number.NaN);
        const excl = Number.isFinite(explicitExcl)
          ? explicitExcl
          : exclFromIncl(incl, tax);
        return {
          incl,
          tax,
          excl,
          customSubjectName: restoreCustomSubjectName(item),
          billingSubjectName: restoreBillingSubjectName(item),
        };
      };
      return {
        device: restoreItem("cost_it_device", params.cost_it_device),
        construction: restoreItem("cost_it_construction", params.cost_it_construction),
        survey: restoreItem("cost_it_survey", params.cost_it_survey),
        integration: restoreItem("cost_it_integration", params.cost_it_integration),
        other: restoreItem("cost_it_other", params.cost_it_other),
        maintenance: restoreItem("cost_it_maintenance", params.cost_it_maintenance),
        running: restoreItem("cost_it_running", params.cost_it_running),
        bidding: restoreItem("cost_it_bidding", params.cost_it_bidding),
        design_eval: restoreItem("cost_it_design_eval", params.cost_it_design_eval),
        audit: restoreItem("cost_it_audit", params.cost_it_audit),
      };
    } catch (e) {
      console.error("加载甄选前方案 IT 投入失败:", e);
      return null;
    }
  }, [activeScheme?.stage, activeProject?.id, preScheme?.id, state.costIt]);

  // 切换到指定方案：先走未保存变更确认，再复用与项目下拉一致的导航加载路径。
  const switchToScheme = async (schemeId: string) => {
    if (!activeProject || pendingNewSchemeName) return;
    if (schemeId === activeScheme?.id) return;
    const canProceed = await confirmOrSave();
    if (!canProceed) return;
    // 导航 store 的 activeSchemeId 是"返回/重进"依赖的来源，这里始终同步更新。
    // 但保存/派生方案走的是 loadProjectContext 直接加载，不会刷新 store —— 此时 store
    // 可能陈旧地仍指向目标方案：navigateTo 写入相同值不会改变 [activeProjectId,
    // activeSchemeId] 依赖，加载 effect 不触发，切换会"看似无效"。故当 store 已指向目标 id
    // 时，直接调用 loadProjectContext 兜底，保证切换必然生效。
    const navSchemeId = useNavigationStore.getState().activeSchemeId;
    navigateTo("ict_lifecycle", activeProject.id, schemeId);
    if (navSchemeId === schemeId) {
      await loadProjectContext(activeProject.id, schemeId);
    }
  };

  // 就地更新当前方案的甄选阶段标签（不派生新方案），并同步本地方案列表。
  const applyActiveSchemeStage = async (stage: SchemeStage | null) => {
    if (!activeProject || !activeScheme || pendingNewSchemeName) return;
    try {
      const updated = await projectService.updateSchemeStage(
        activeProject.id,
        activeScheme.id,
        stage
      );
      setActiveScheme(updated);
      setSchemes(prev => prev.map(s => (s.id === updated.id ? updated : s)));
    } catch (error) {
      console.error("更新甄选阶段失败:", error);
      alert("更新甄选阶段失败: " + error);
    }
  };

  // 以当前测算数据为起点，派生一个带阶段标签的新方案（复用"另存为新方案"弹窗）。
  const handleDeriveStageScheme = (stage: SchemeStage) => {
    const base = activeProject?.name || activeScheme?.name || "方案";
    setSaveAsSchemeStage(stage);
    setSaveAsSchemeName(`${base}_${getSchemeStageLabel(stage)}`);
    setShowSaveAsModal(true);
  };

  // 二段式阶段按钮点击：已存在则切换加载；不存在时按当前方案是否已标注决定"就地打标"或"派生新方案"。
  const handleStageButtonClick = async (stage: SchemeStage) => {
    if (!activeProject || pendingNewSchemeName) return;
    const target = stage === "pre_selection" ? preScheme : postScheme;
    if (target) {
      await switchToScheme(target.id);
      return;
    }
    if (!activeScheme) return;
    if (!activeScheme.stage) {
      // 当前方案尚未打阶段标签（如首次测算生成的方案）：直接标注为该阶段。
      await applyActiveSchemeStage(stage);
    } else {
      // 当前方案已属另一阶段：基于当前数据派生一个该阶段的新方案。
      handleDeriveStageScheme(stage);
    }
  };

  const {
    activeTab, setActiveTab,
    projName,
    customerName,
    projectYears,
    balanceAllocation,
    updateBalanceRule,
    subjectFundingPlans,
    upsertSubjectFundingPlan,
    revIt,
    revCt,
    revNonItCt,
    costIt,
    costCt,
    costMix,
    templates,
    selectedTemplate, setSelectedTemplate,
    reconciliationErrors, setReconciliationErrors,
    showReconciliationModal, setShowReconciliationModal,
    currentTotalDifference, setCurrentTotalDifference,
    pendingTab, setPendingTab,
    showConfirmIgnore, setShowConfirmIgnore,
    ignoredDataHash, setIgnoredDataHash,
    setIgnoredTailValue,
    loadTemplates,
    updateTaxItem,
    commitTaxItemIncl,
    splitTaxItemIncl,
    cancelTaxItemSplit,
    clearFinancialSubjects,
    updateTaxItemCustomSubjectName,
    updateTaxItemBillingSubjectName,
  } = state;

  const {
    metrics,
    selQuote,
    selMarkup,
    selActualCost,
    selFee,
    selLimit,
    selectionFeeAnchor, setSelectionFeeAnchor,
    revMode, setRevMode,
    revTargetType, setRevTargetType,
    revTargetValue, setRevTargetValue,
    revSubjectRefKey, setRevSubjectRefKey,
    subjectFundingCoverage,
    subjectFundingAnnualCashflow,
    subjectFundingCalculationBlocked,
    performReverseCalculation,
    applySelectionLimit,
    handleSelFeeChange,
  } = calculations;

  const setActiveModule = useAiContextStore(storeState => storeState.setActiveModule);

  useEffect(() => {
    setActiveModule('ict');
  }, [setActiveModule]);

  const balanceSubjectItems = useMemo<BalanceSubjectItem[]>(() => {
    const resolveItem = (subject: IctSubjectDefinition) => {
      if (subject.groupId === "revIt") return revIt[subject.key as keyof typeof revIt] || null;
      if (subject.groupId === "revCt") return revCt[subject.key as keyof typeof revCt] || null;
      if (subject.groupId === "revNonItCt") return revNonItCt;
      if (subject.groupId === "costIt") return costIt[subject.key as keyof typeof costIt] || null;
      if (subject.groupId === "costCt") return costCt[subject.key as keyof typeof costCt] || null;
      if (subject.groupId === "costMix") return costMix[subject.key as keyof typeof costMix] || null;
      return null;
    };

    return ICT_SUBJECT_DEFINITIONS.map(subject => ({
      subject,
      item: resolveItem(subject),
    }));
  }, [revIt, revCt, revNonItCt, costIt, costCt, costMix]);

  const revenueBalanceSubjects = useMemo(
    () => balanceSubjectItems.filter(row => row.subject.side === "revenue"),
    [balanceSubjectItems],
  );
  const investmentBalanceSubjects = useMemo(
    () => balanceSubjectItems.filter(row => row.subject.side === "cost"),
    [balanceSubjectItems],
  );

  const revenueBalanceEvaluation = useMemo(
    () => evaluateBalanceRule("revenue", balanceAllocation.revenue, revenueBalanceSubjects),
    [balanceAllocation.revenue, revenueBalanceSubjects],
  );
  const investmentBalanceEvaluation = useMemo(
    () => evaluateBalanceRule("investment", balanceAllocation.investment, investmentBalanceSubjects),
    [balanceAllocation.investment, investmentBalanceSubjects],
  );

  const applyBalanceRule = useCallback((evaluation: BalanceRuleEvaluation) => {
    if (!evaluation.canApply || !evaluation.balancingSubject || evaluation.autoAmount === null) return;
    const currentIncl = Number(evaluation.balancingItem?.incl ?? 0);
    if (Math.abs(currentIncl - evaluation.autoAmount) <= 0.004) return;
    updateTaxItem(
      evaluation.balancingSubject.groupId,
      evaluation.balancingSubject.key,
      "incl",
      evaluation.autoAmount,
    );
  }, [updateTaxItem]);

  useEffect(() => {
    applyBalanceRule(revenueBalanceEvaluation);
    applyBalanceRule(investmentBalanceEvaluation);
  }, [applyBalanceRule, revenueBalanceEvaluation, investmentBalanceEvaluation]);

  const blockingBalanceMessages = useMemo(
    () => [revenueBalanceEvaluation, investmentBalanceEvaluation]
      .filter(evaluation => evaluation.status === "negative" && evaluation.message)
      .map(evaluation => evaluation.message as string),
    [revenueBalanceEvaluation, investmentBalanceEvaluation],
  );

  const reverseBalanceEvaluation = revMode === "revenue"
    ? revenueBalanceEvaluation
    : investmentBalanceEvaluation;

  const reverseSubjectOptions = useMemo(
    () => getReverseEligibleSubjects(
      revMode,
      balanceSubjectItems,
      revMode === "revenue" ? revenueBalanceEvaluation : investmentBalanceEvaluation,
    ),
    [revMode, balanceSubjectItems, revenueBalanceEvaluation, investmentBalanceEvaluation],
  );

  const selectedReverseSubject = useMemo(
    () => findReverseSubjectOption(reverseSubjectOptions, revSubjectRefKey),
    [reverseSubjectOptions, revSubjectRefKey],
  );

  const reverseCalculationContext = useMemo(
    () => resolveReverseCalculationContext({
      option: selectedReverseSubject,
      subjects: balanceSubjectItems,
      sameSideBalanceEvaluation: reverseBalanceEvaluation,
    }),
    [selectedReverseSubject, balanceSubjectItems, reverseBalanceEvaluation],
  );

  const reverseContextMessage = reverseCalculationContext.mode === "blocked"
    ? reverseCalculationContext.message
    : null;

  useEffect(() => {
    if (!revSubjectRefKey) return;
    const selected = findReverseSubjectOption(reverseSubjectOptions, revSubjectRefKey);
    if (!selected || selected.disabledReason) {
      setRevSubjectRefKey("");
    }
  }, [revSubjectRefKey, reverseSubjectOptions, setRevSubjectRefKey]);

  const executeReverseCalculation = () => {
    if (blockingBalanceMessages.length > 0) {
      alert(blockingBalanceMessages.join("\n"));
      return;
    }
    if (!selectedReverseSubject) {
      alert("请先选择反算目标科目。");
      return;
    }
    if (reverseContextMessage) {
      alert(reverseContextMessage);
      return;
    }
    performReverseCalculation(selectedReverseSubject, reverseCalculationContext);
  };

  useEffect(() => {
    if (!activeProject?.id || !workspaceId) return;

    registerSaveHandler("lifecycle", async (context) => {
      if (context.workspaceId !== workspaceId || context.projectId !== activeProject.id) {
        throw new Error("项目或工作区已切换");
      }
      const inputPayload = calculations.buildInputDataPayload();
      // 工作副本写入当前方案自己的桶，保证方案间隔离。
      const lifecycleSchemeId = activeScheme?.id || activeProject.default_scheme_id || null;
      await domainSaveService.saveLifecycleState(activeProject.id, lifecycleSchemeId, {
        profileJson: {
          projectName: state.projName,
          customerName: state.customerName,
          propertyRights: state.propertyRights,
        },
        parametersJson: {
          projectYears: state.projectYears,
          discountRate: state.discountRate,
          ignoreTailDifference: state.ignoredTailValue !== null,
          tailDifferenceValue: state.ignoredTailValue,
          balanceAllocation: serializeBalanceAllocationState(state.balanceAllocation),
          cashflowCalculationSource: state.cashflowCalculationSource,
          subjectFundingPlanMigrationVersion: state.subjectFundingPlanMigrationVersion,
        },
        backgroundJson: {
          projectBackground: state.projectBackground,
        },
        inputPayloadJson: inputPayload,
      });
      return { success: true, savedScopes: ["lifecycle"] };
    });

    registerSaveHandler("cashflow", async (context) => {
      if (context.workspaceId !== workspaceId || context.projectId !== activeProject.id) {
        throw new Error("项目或工作区已切换");
      }
      const cashflowState = {
        cashflowModel: state.cashflowModel,
        paymentModelJson: {
          cashflowModel: state.cashflowModel,
          revDistribution: state.distRev,
          costDistribution: state.distCost,
          segmentValueMode: state.segmentValueMode,
          cashflowCalculationSource: state.cashflowCalculationSource,
          subjectFundingPlanMigrationVersion: state.subjectFundingPlanMigrationVersion,
        },
        yearlyCashflowJson: {
          cashflowTable: calculations.cashflowTable,
          directSegmentCashflow: calculations.directSegmentCashflow,
          subjectFundingAnnualCashflow: calculations.subjectFundingAnnualCashflow,
        },
        sectorCashflowJson: {
          cashflowSegments: state.cashflowSegments,
        },
        assumptionsJson: {
          projectYears: state.projectYears,
          discountRate: state.discountRate,
          revIt: state.revIt,
          revCt: state.revCt,
          revNonItCt: state.revNonItCt,
          costIt: state.costIt,
          costCt: state.costCt,
          costMix: state.costMix,
          balanceAllocation: state.balanceAllocation,
          cashflowCalculationSource: state.cashflowCalculationSource,
          subjectFundingPlans: state.subjectFundingPlans,
          subjectFundingPlanMigrationVersion: state.subjectFundingPlanMigrationVersion,
        },
        metricsJson: calculations.metrics,
      };
      // 工作副本写入当前方案自己的桶，保证方案间隔离。
      const cashflowSchemeId = activeScheme?.id || activeProject.default_scheme_id || null;
      await domainSaveService.saveCashflowState(activeProject.id, cashflowSchemeId, cashflowState);
      return { success: true, savedScopes: ["cashflow"] };
    });

    registerSaveHandler("benefit-analysis", async (context) => {
      if (context.workspaceId !== workspaceId || context.projectId !== activeProject.id) {
        throw new Error("项目或工作区已切换");
      }
      const blockingMessage = getSubjectFundingBlockingMessage("保存效益方案");
      if (blockingMessage) {
        throw new Error(blockingMessage);
      }
      const inputPayload = calculations.buildInputDataPayload();
      const schemeId = pendingNewSchemeName ? null : (activeScheme?.id || activeProject.default_scheme_id || null);
      const schemeName = pendingNewSchemeName || activeScheme?.name || activeProject.name || "默认测算方案";
      const updatedProject = await domainSaveService.saveBenefitAnalysis(
        activeProject.id,
        schemeId,
        schemeName,
        inputPayload,
        calculations.metrics,
        !!pendingNewSchemeName,
      );
      if (useProjectStore.getState().currentProject?.id === activeProject.id) {
        useProjectStore.getState().setCurrentProject(updatedProject);
      }
      setActiveProject(updatedProject);
      setProjects(prev => prev.map(project => project.id === updatedProject.id ? updatedProject : project));
      setPendingNewSchemeName(null);
      return { success: true, savedScopes: ["benefit-analysis"] };
    });

    return () => {
      unregisterSaveHandler("lifecycle");
      unregisterSaveHandler("cashflow");
      unregisterSaveHandler("benefit-analysis");
    };
  }, [
    activeProject?.id,
    activeProject?.default_scheme_id,
    activeProject?.name,
    activeScheme?.id,
    activeScheme?.name,
    pendingNewSchemeName,
    workspaceId,
    registerSaveHandler,
    unregisterSaveHandler,
    getSubjectFundingBlockingMessage,
    calculations,
    state,
  ]);

  useEffect(() => {
    if (isHydratingRef.current || !activeProject?.id) return;
    markDirty("lifecycle");
  }, [activeProject?.id, markDirty, state.projName, state.customerName, state.propertyRights, state.projectBackground]);

  useEffect(() => {
    if (isHydratingRef.current || !activeProject?.id) return;
    markDirty("cashflow");
  }, [
    activeProject?.id,
    markDirty,
    state.discountRate,
    state.projectYears,
    state.cashflowModel,
    state.distRev,
    state.distCost,
    state.segmentValueMode,
    state.cashflowSegments,
    state.cashflowCalculationSource,
    state.subjectFundingPlanMigrationVersion,
    state.revIt,
    state.revCt,
    state.revNonItCt,
    state.costIt,
    state.costCt,
    state.costMix,
    state.balanceAllocation,
    state.subjectFundingPlans,
  ]);

  useEffect(() => {
    if (isHydratingRef.current || !activeProject?.id) return;
    markDirty("benefit-analysis");
  }, [
    activeProject?.id,
    markDirty,
    selQuote,
    selMarkup,
    selActualCost,
    selFee,
    selLimit,
    selectionFeeAnchor,
  ]);

  const handleTabSwitch = async (tab: string, templateName?: string, forceIgnore = false) => {
    if (templateName && templateName !== selectedTemplate && dirtyScopes.includes("template-forms")) {
      const canProceed = await confirmOrSave();
      if (!canProceed) return;
    }

    if ((tab === 'cashflow' || tab === 'generate') && blockingBalanceMessages.length > 0) {
      alert(blockingBalanceMessages.join("\n"));
      return;
    }

    if (tab === 'generate') {
      const blockingMessage = getSubjectFundingBlockingMessage("进入文档生成");
      if (blockingMessage) {
        alert(blockingMessage);
        return;
      }

      // 甄选结果签批表等"签批"类文档，建议基于甄选后方案生成。非阻断提示，用户可继续。
      if (
        templateName &&
        templateName.includes("签批") &&
        activeScheme &&
        activeScheme.stage !== "post_selection"
      ) {
        const proceed = window.confirm(
          `当前方案「${activeScheme.name}」为${getSchemeStageLabel(activeScheme.stage)}阶段。\n` +
            `“${templateName.replace(/\.(docx|xlsx)$/i, "")}”通常应基于甄选后方案生成，\n` +
            `建议先切换到“甄选后”方案再生成。\n\n确定：仍然继续生成\n取消：返回切换方案`
        );
        if (!proceed) return;
      }
    }

    const currentHash = buildFinancialStateHash({ revIt, revCt, revNonItCt, costIt, costCt, costMix });

    if ((tab === 'cashflow' || tab === 'generate') && !forceIgnore) {
      if (currentHash !== ignoredDataHash) {
        const { errors, totalDifference } = validateFinancialData(
          { it: revIt, ct: revCt, non_it_ct: { item: revNonItCt } },
          { it: costIt, ct: costCt, mix: costMix }
        );
        if (errors.length > 0) {
          setReconciliationErrors(errors);
          setCurrentTotalDifference(totalDifference);
          setPendingTab({ tab, template: templateName });
          setShowReconciliationModal(true);
          return;
        }
      }
    }

    if (forceIgnore) {
      setIgnoredTailValue(currentTotalDifference);
      setIgnoredDataHash(currentHash);
    }

    setActiveTab(tab as any);
    if (templateName) {
      setSelectedTemplate(templateName);
      setActiveModule(buildAiContextKey('ict', 'template', templateName));
    } else {
      setActiveModule(AI_CONTEXT_KEY.ICT_CORE);
    }
  };

  const parseRuleAmount = (val: string) => {
    if (val === "") return null;
    const num = Number(val);
    return isNaN(num) ? null : num;
  };

  const getEvaluationForSide = (side: BalanceAllocationSide) => {
    return side === "revenue" ? revenueBalanceEvaluation : investmentBalanceEvaluation;
  };

  const isSubjectAutoBalanced = (subject: IctSubjectDefinition) => {
    const side = subject.side === "revenue" ? "revenue" : "investment";
    const rule = balanceAllocation[side];
    return rule.enabled && isBalanceSubjectMatch(rule.balancingSubject, subject);
  };

  const formatFundingControlMoney = (value: number) =>
    new Intl.NumberFormat("zh-CN", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(Number.isFinite(value) ? value : 0);

  const formatAnnualCashflowPreview = (values: number[]) => {
    const parts = values
      .map((value, index) => ({ value, year: index + 1 }))
      .filter(item => Math.abs(item.value) > 0.005)
      .map(item => `第${item.year}年 ${formatFundingControlMoney(item.value)} 元`);
    return parts.length > 0 ? parts.join("，") : "暂无有效年度金额";
  };

  const handleLocateFirstCoverageIssue = () => {
    const firstIssue = subjectFundingCoverage.issues[0];
    if (!firstIssue) return;
    const planId = createSubjectFundingPlanId(firstIssue.subjectRef);
    setFundingPlanFocus({ planId, token: Date.now() });
    scrollToSubject(
      firstIssue.subjectRef.side,
      firstIssue.subjectRef.groupId,
      firstIssue.subjectRef.key,
      activeTab,
      setActiveTab,
    );
  };

  const handleClearAllFinancialSubjects = () => {
    const confirmed = window.confirm(
      "确定要一键清空全部收入和投入吗？\n\n将清空全部收入金额、投入/支出金额、具体业务/产品名称、计费科目名称、科目资金计划、差额承接和反算目标状态；标准科目目录会保留。此操作无法撤销。"
    );
    if (!confirmed) return;

    clearFinancialSubjects();
    setRevSubjectRefKey("");
    setFundingPlanFocus(null);
  };

  const renderCashflowCalculationSourceControl = () => {
    const counts = subjectFundingCoverage.counts;

    const statusBadge = subjectFundingCalculationBlocked
      ? { cls: "bg-warning-soft text-warning-foreground", label: `${subjectFundingCoverage.issues.length} 项待处理` }
      : { cls: "bg-success-soft text-success-foreground", label: "可用于科目级计算" };

    return (
      <div className="table-card bg-card border border-border rounded-xl px-5 py-3 shadow-sm mb-6">
        {/* 收起态：仅一行状态摘要 + 展开按钮，避免长期占据 1/3 页面 */}
        <button
          type="button"
          onClick={() => setCashflowSourceExpanded(prev => !prev)}
          className="w-full flex items-center justify-between gap-3 text-left"
          aria-expanded={cashflowSourceExpanded}
        >
          <div className="flex flex-wrap items-center gap-2 min-w-0">
            <h3 className="text-sm font-extrabold text-foreground shrink-0">现金流计算口径</h3>
            <span className="rounded-md bg-primary-soft px-2 py-0.5 text-[11px] font-extrabold text-primary shrink-0">科目级资金计划</span>
            <span className={`rounded-md px-2 py-0.5 text-[11px] font-bold ${statusBadge.cls}`}>{statusBadge.label}</span>
            <span className="text-[11px] font-semibold text-secondary-foreground numeric-value shrink-0">
              收入 {counts.revenuePlannedCount}/{counts.revenueSubjectCount}{"　"}投入 {counts.costPlannedCount}/{counts.costSubjectCount}
            </span>
          </div>
          <span className="inline-flex items-center gap-1 text-[11px] font-bold text-secondary-foreground shrink-0">
            {cashflowSourceExpanded ? "收起" : "展开"}
            <AppIcon name={cashflowSourceExpanded ? "chevronUp" : "chevronDown"} size={14} />
          </span>
        </button>

        {!cashflowSourceExpanded && (
          <p className="mt-1.5 text-[11px] leading-relaxed text-secondary-foreground">
            现金流依据：科目收付款计划，按各科目税率逐年换算为不含税现金流并参与效益指标计算。
            {subjectFundingCalculationBlocked && <span className="text-warning-foreground font-semibold">{"　"}部分科目未维护计划，已临时按「第一年一次性」计入，展开可补全。</span>}
          </p>
        )}

        {cashflowSourceExpanded && (
        <div className="mt-4 flex flex-col gap-4">
        <p className="text-xs leading-relaxed text-secondary-foreground">
          现金流依据：科目收付款计划。各科目的年度收款 / 付款计划会自动汇总为项目现金流并参与效益指标计算。
        </p>

        <div className={`rounded-lg p-3 text-xs font-semibold leading-relaxed ${subjectFundingCalculationBlocked ? "bg-warning-soft text-warning-foreground" : "bg-success-soft text-success-foreground"}`}>
          {subjectFundingCalculationBlocked
            ? "部分科目尚未维护收付款计划，已按「第一年一次性」临时计入现金流；已维护的多年/按比例计划（含 IT 部分）照常逐年生效。请补全下列计划以获得精确分布。"
            : "当前正式现金流由科目级收付款计划生成，并按各科目税率逐年换算为不含税现金流。"}
        </div>

        {subjectFundingCoverage.issues.length > 0 && (
          <div className="rounded-lg bg-warning-soft/70 p-3 text-xs text-warning-foreground">
            <div className="font-extrabold mb-2">需修正的问题</div>
            <div className="flex flex-col gap-1">
              {subjectFundingCoverage.issues.slice(0, 8).map(issue => (
                <span key={`${issue.planId}-${issue.type}`}>{issue.message}</span>
              ))}
              {subjectFundingCoverage.issues.length > 8 && (
                <span>另有 {subjectFundingCoverage.issues.length - 8} 项问题未展开。</span>
              )}
            </div>
          </div>
        )}

        {subjectFundingCoverage.valid && (
          <div className="grid grid-cols-1 gap-1 text-[11px] font-semibold text-secondary-foreground">
            <div>科目计划现金流入(不含税)：{formatAnnualCashflowPreview(subjectFundingAnnualCashflow.annualRevenueExcl)}</div>
            <div>科目计划现金流出(不含税)：{formatAnnualCashflowPreview(subjectFundingAnnualCashflow.annualCostExcl)}</div>
            <div className="mt-1 opacity-70">提示：完整年度明细及穿透视图将在下方测算结果表格中展示。</div>
          </div>
        )}

        <div className="mt-3 flex flex-wrap gap-2">
          <button
                type="button"
                onClick={() => {
                  const missing: Array<{ subjectRef: SubjectFundingSubjectRef; amountIncl: number }> = [];
                  for (const subj of subjectFundingCoverage.revenueSubjects) {
                    const id = createSubjectFundingPlanId(subj.subjectRef);
                    if (!state.subjectFundingPlans[id] && subj.subjectAmountIncl > 0) {
                      missing.push({ subjectRef: subj.subjectRef, amountIncl: subj.subjectAmountIncl });
                    }
                  }
                  if (missing.length > 0) {
                    state.setSubjectFundingPlans(initializeMissingSubjectFundingPlans(state.subjectFundingPlans, missing));
                  }
                }}
                className="inline-flex items-center gap-1.5 rounded-md bg-card px-3 py-1.5 text-[11px] font-bold text-primary shadow-sm hover:bg-primary-soft transition-colors border border-border"
              >
                <AppIcon name="quickAction" size={12} />
                一键生成未配置的收入计划
              </button>
              <button
                type="button"
                onClick={() => {
                  const missing: Array<{ subjectRef: SubjectFundingSubjectRef; amountIncl: number }> = [];
                  for (const subj of subjectFundingCoverage.costSubjects) {
                    const id = createSubjectFundingPlanId(subj.subjectRef);
                    if (!state.subjectFundingPlans[id] && subj.subjectAmountIncl > 0) {
                      missing.push({ subjectRef: subj.subjectRef, amountIncl: subj.subjectAmountIncl });
                    }
                  }
                  if (missing.length > 0) {
                    state.setSubjectFundingPlans(initializeMissingSubjectFundingPlans(state.subjectFundingPlans, missing));
                  }
                }}
                className="inline-flex items-center gap-1.5 rounded-md bg-card px-3 py-1.5 text-[11px] font-bold text-primary shadow-sm hover:bg-primary-soft transition-colors border border-border"
              >
                <AppIcon name="quickAction" size={12} />
                一键生成未配置的投入计划
              </button>

          {subjectFundingCoverage.issues.length > 0 && (
            <button
              type="button"
              onClick={handleLocateFirstCoverageIssue}
              className="inline-flex items-center gap-1.5 rounded-md bg-card px-3 py-1.5 text-[11px] font-bold text-warning shadow-sm hover:bg-warning-soft transition-colors border border-border"
            >
              <AppIcon name="search" size={12} />
              定位到第一个待处理项
            </button>
          )}

          <button
            type="button"
            onClick={handleClearAllFinancialSubjects}
            className="inline-flex items-center gap-1.5 rounded-md bg-card px-3 py-1.5 text-[11px] font-bold text-destructive shadow-sm hover:bg-destructive/10 transition-colors border border-border"
          >
            <AppIcon name="delete" size={12} />
            一键清空全部收入和支出
          </button>
        </div>
        </div>
        )}
      </div>
    );
  };

  const renderBalanceControl = (side: BalanceAllocationSide) => {
    const rule = balanceAllocation[side];
    const evaluation = getEvaluationForSide(side);
    const sideLabel = side === "revenue" ? "收入" : "投入";

    const handleLocate = () => {
      if (evaluation.balancingSubject) {
        scrollToSubject(
          evaluation.balancingSubject.side,
          evaluation.balancingSubject.groupId,
          evaluation.balancingSubject.key,
          activeTab,
          setActiveTab
        );
      }
    };

    const handleClear = () => {
      updateBalanceRule(side, {
        balancingSubject: null,
        enabled: rule.totalInclAmount !== null,
      });
    };

    return (
      <div className="table-card bg-card border border-border rounded-xl p-6 shadow-sm mb-6 flex flex-col gap-4">
        <div className="flex flex-col md:flex-row md:items-start justify-between gap-4">
          <div className="flex flex-col gap-1.5 flex-1 max-w-sm">
            <label className="text-sm font-bold text-foreground">
              {sideLabel}含税总金额 (元)
            </label>
            <div className="flex items-center gap-2 h-[38px]">
              <input
                type="number"
                placeholder={`请输入${sideLabel}含税总金额`}
                className="bg-card border border-input px-3 py-2 rounded-md outline-none text-sm w-full font-semibold focus:border-ring h-full"
                value={rule.totalInclAmount === null ? "" : rule.totalInclAmount}
                onChange={e => {
                  const val = parseRuleAmount(e.target.value);
                  updateBalanceRule(side, {
                    totalInclAmount: val,
                    enabled: val !== null || rule.balancingSubject !== null,
                  });
                }}
              />
            </div>
            {side === "revenue" && rule.totalInclAmount !== null && (
              <span className="text-[10px] text-primary font-bold">
                产品含税收入最低提示：{(rule.totalInclAmount * 0.01).toFixed(2)} 元 (总金额 1%)
              </span>
            )}
          </div>

          <div className="flex flex-col gap-1.5 flex-1">
            <label className="text-sm font-bold text-foreground">
              差额承接科目
            </label>
            <div className="h-[38px] flex items-center">
              {rule.enabled && evaluation.balancingSubject ? (
                <SelectedSubjectRoleSummary
                  subject={evaluation.balancingSubject}
                  item={evaluation.balancingItem}
                  onLocate={handleLocate}
                  onClear={handleClear}
                />
              ) : (
                <span className="text-xs text-secondary-foreground font-semibold">
                  尚未指定。请在下方{sideLabel}科目列表中点击“设置角色”进行设置。
                </span>
              )}
            </div>
          </div>
        </div>

        {rule.enabled && evaluation.status === "negative" && evaluation.message && (
          <div className="text-xs text-red-500 font-bold bg-red-50 p-2.5 rounded-lg border border-red-200/30">
            {evaluation.message}
          </div>
        )}
      </div>
    );
  };

  const renderTaxGroup = (title: string, groupId: string, groupState: any, items: IctSubjectDefinition[]) => (
    <div className="table-card bg-card border border-border rounded-xl p-6 shadow-sm mb-6">
      <h3 className="font-bold text-lg mb-4">{title}</h3>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {items.map(item => {
          const itemErr = reconciliationErrors.find(e => e.key === `${groupId}.${item.key}`);
          const currentItem = groupState[item.key];
          const autoBalanced = isSubjectAutoBalanced(item);
          // 已拆分科目按两笔子金额与业务系统逐笔闭合，不再对合计做尾差检查。
          const splitParts: TaxSplitPart[] | null = currentItem?.splitParts?.length ? currentItem.splitParts : null;
          // 财务口径检查：含税录入值若不可精确表示（round(excl×(1+r)) ≠ 录入值），提示并在失焦后归一。
          const inclPair = !splitParts && Number(currentItem?.incl) > 0
            ? normalizeTaxPairFromIncl(currentItem.incl, currentItem.tax)
            : null;
          const inclAdjust = inclPair?.adjusted ? inclPair : null;
          const splitPreview = inclAdjust && !autoBalanced
            ? splitInclAmount(currentItem.incl, currentItem.tax)
            : null;
          const customSubjectName = getSubjectCustomName(currentItem);
          const billingSubjectName = getSubjectBillingName(currentItem);
          const displayName = getSubjectExcelDisplayName(item, currentItem);

          const balanceSide = item.side === "revenue" ? "revenue" : "investment";
          const rule = balanceAllocation[balanceSide];
          const isBalancing = rule.enabled && isBalanceSubjectMatch(rule.balancingSubject, item);
          const isReverseTarget = revSubjectRefKey === getReverseSubjectRefKey(getReverseSubjectRef(item));
          const fundingPlanId = createSubjectFundingPlanId({
            side: item.side,
            groupId: item.groupId,
            key: item.key,
          });
          const fundingPlan = subjectFundingPlans[fundingPlanId];

          const borderClass = isBalancing
            ? "border-l-2 border-warning/70 bg-warning-soft/30 pl-1.5"
            : isReverseTarget
            ? "border-l-2 border-primary/70 bg-primary-soft/30 pl-1.5"
            : "border-l-2 border-transparent pl-1.5";

          return (
            <div
              key={item.key}
              id={`subject-anchor-${item.side}-${groupId}-${item.key}`}
              className={`flex flex-col gap-1.5 py-2 pr-2 rounded-lg transition-all duration-300 ${borderClass}`}
            >
              <div className="flex items-center justify-between gap-2 flex-wrap min-w-0">
                <div className="flex items-center gap-2 flex-wrap min-w-0">
                  <label className="text-sm font-semibold text-secondary-foreground shrink-0">{item.standardSubjectName}</label>
                  {(customSubjectName || billingSubjectName) && (
                    <span className="text-[10px] font-medium text-primary truncate max-w-[150px] sm:max-w-[200px]" title={displayName}>
                      ({displayName})
                    </span>
                  )}
                </div>
                <SubjectRoleActions
                  subject={item}
                  item={currentItem}
                  balanceAllocation={balanceAllocation}
                  updateBalanceRule={updateBalanceRule}
                  updateTaxItem={updateTaxItem}
                  revSubjectRefKey={revSubjectRefKey}
                  setRevSubjectRefKey={setRevSubjectRefKey}
                  setRevMode={setRevMode}
                  subjects={balanceSubjectItems}
                />
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                <input
                  type="text"
                  placeholder="具体业务/产品名称"
                  className="bg-muted px-3 py-2 rounded-md outline-none text-xs focus:bg-card focus:ring-1 focus:ring-ring"
                  value={customSubjectName}
                  onChange={e => updateTaxItemCustomSubjectName(groupId, item.key, e.target.value)}
                />
                <input
                  type="text"
                  placeholder="计费科目名称（文书/计费口径）"
                  className="bg-muted px-3 py-2 rounded-md outline-none text-xs focus:bg-card focus:ring-1 focus:ring-ring"
                  value={billingSubjectName}
                  onChange={e => updateTaxItemBillingSubjectName(groupId, item.key, e.target.value)}
                />
              </div>
              <div className="flex gap-2">
                <input type="number" placeholder="含税" readOnly={autoBalanced || !!splitParts} title={autoBalanced ? "该金额由总金额自动计算" : splitParts ? "已拆分为两笔，如需修改请先取消拆分" : undefined} className={`w-full px-3 py-2 rounded-md outline-none text-sm ${autoBalanced || splitParts ? "bg-muted text-secondary-foreground cursor-not-allowed" : `bg-card border ${inclAdjust ? "border-amber-500 ring-1 ring-amber-400" : "border-input"}`}`} value={autoBalanced ? currentItem.incl : currentItem.incl === 0 ? "" : currentItem.incl} onChange={e => {
                  if (!autoBalanced && !splitParts) updateTaxItem(groupId, item.key, 'incl', Number(e.target.value));
                }} onBlur={() => {
                  if (!autoBalanced && !splitParts) commitTaxItemIncl(groupId, item.key);
                }} />
                <input type="number" placeholder="税率" className="w-20 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm" value={currentItem.tax} onChange={e => updateTaxItem(groupId, item.key, 'tax', Number(e.target.value))} />
                <input type="number" placeholder="不含税" readOnly={autoBalanced || !!splitParts} title={autoBalanced ? "该金额由总金额自动计算" : splitParts ? "已拆分为两笔，不含税为两笔之和" : undefined} className={`w-full px-3 py-2 rounded-md outline-none text-sm ${autoBalanced || splitParts ? "bg-muted text-secondary-foreground cursor-not-allowed" : `bg-card border focus:border-ring ${itemErr ? 'border-red-500 ring-1 ring-red-500' : 'border-input'}`}`} value={autoBalanced ? currentItem.excl : currentItem.excl === 0 ? "" : currentItem.excl} onChange={e => {
                  if (!autoBalanced && !splitParts) updateTaxItem(groupId, item.key, 'excl', Number(e.target.value));
                }} />
              </div>
              {autoBalanced && <span className="text-[10px] font-bold text-primary">该科目金额由总金额自动计算，税率可继续编辑。</span>}
              {inclAdjust && (
                <div className="rounded-md border border-amber-300 bg-amber-50/60 px-2.5 py-1.5 flex flex-col gap-1">
                  <div className="flex items-center gap-1.5 flex-wrap">
                    <span className="px-1.5 py-0.5 rounded bg-amber-500 text-white text-[10px] font-bold tabular-nums shrink-0">
                      尾差 {inclAdjust.incl > inclAdjust.enteredIncl ? "+" : ""}{roundMoneyHalfUp(inclAdjust.incl - inclAdjust.enteredIncl).toFixed(2)} 元
                    </span>
                    <span className="text-[10px] text-amber-700 font-semibold tabular-nums">
                      录入 {inclAdjust.enteredIncl.toFixed(2)} → 业务系统 {inclAdjust.incl.toFixed(2)}（按不含税 {inclAdjust.excl.toFixed(2)} 反推）
                    </span>
                  </div>
                  {autoBalanced ? (
                    <span className="text-[10px] text-secondary-foreground">该科目由总金额自动计算，可调整总金额分配，或在「设置 → 测算行为」开启自动修正。</span>
                  ) : (
                    <div className="flex items-center gap-2 flex-wrap">
                      <button
                        type="button"
                        className="text-[10px] font-bold px-2 py-0.5 rounded border border-amber-400 text-amber-700 hover:bg-amber-100 tabular-nums"
                        onClick={() => updateTaxItem(groupId, item.key, 'incl', inclAdjust.incl)}
                      >
                        改为 {inclAdjust.incl.toFixed(2)}
                      </button>
                      {splitPreview && (
                        <button
                          type="button"
                          className="text-[10px] font-bold px-2 py-0.5 rounded border border-amber-400 text-amber-700 hover:bg-amber-100 tabular-nums"
                          onClick={() => splitTaxItemIncl(groupId, item.key)}
                        >
                          拆分两笔 {splitPreview.map(part => part.incl.toFixed(2)).join(" + ")}
                        </button>
                      )}
                      <span className="text-[10px] text-secondary-foreground">
                        {taxInclAutoFix ? "失焦后将自动修正" : "「设置 → 测算行为」可自动修正"}
                      </span>
                    </div>
                  )}
                </div>
              )}
              {splitParts && (
                <div className="rounded-md border border-emerald-200 bg-emerald-50/60 px-2.5 py-1.5 flex flex-col gap-1">
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-[10px] font-bold text-emerald-700">已拆两笔 · 每笔与业务系统口径闭合</span>
                    <button
                      type="button"
                      className="text-[10px] underline text-primary shrink-0"
                      onClick={() => cancelTaxItemSplit(groupId, item.key)}
                    >
                      取消拆分
                    </button>
                  </div>
                  <div className="grid grid-cols-[auto_1fr_1fr] gap-x-3 gap-y-0.5 items-baseline max-w-xs">
                    <span />
                    {splitParts.map((_, index) => (
                      <span key={index} className="text-[10px] text-secondary-foreground text-right">第 {index + 1} 笔</span>
                    ))}
                    <span className="text-[10px] text-secondary-foreground">含税</span>
                    {splitParts.map((part, index) => (
                      <span key={index} className="text-sm font-bold text-emerald-700 tabular-nums text-right">{part.incl.toFixed(2)}</span>
                    ))}
                    <span className="text-[10px] text-secondary-foreground">不含税</span>
                    {splitParts.map((part, index) => (
                      <span key={index} className="text-sm font-bold text-emerald-700 tabular-nums text-right">{part.excl.toFixed(2)}</span>
                    ))}
                  </div>
                </div>
              )}
              {itemErr && <span className="text-[10px] text-red-500 font-bold">校验失败：{itemErr.field === 'incl' ? '含税与财务口径' : '不含税'}偏离 {itemErr.difference} 元，要求：{itemErr.expectedExcl} 元</span>}
              <IctSubjectFundingPlanEditor
                subject={item}
                item={currentItem}
                plan={fundingPlan}
                displayName={displayName}
                forceOpenToken={fundingPlanFocus?.planId === fundingPlanId ? fundingPlanFocus.token : 0}
                onPlanChange={upsertSubjectFundingPlan}
              />
            </div>
          );
        })}
      </div>
    </div>
  );

  const renderReverseCalculationTools = () => (
    <>
      <h3 className="font-bold text-foreground mb-4">智能反算</h3>
      <div className="bg-card border border-border p-4 rounded-xl flex flex-col gap-4 mb-6">
        <p className="text-xs leading-relaxed text-secondary-foreground">
          反算结果为含税总额参数值，系统会同步调整科目收付款计划并重新生成年度现金流。
        </p>
        <div className="flex flex-col gap-2">
          <label className="text-xs font-bold text-secondary-foreground">反算目标</label>
          {selectedReverseSubject ? (
            <div className="flex flex-col gap-2.5">
              <SelectedSubjectRoleSummary
                subject={selectedReverseSubject.subject}
                item={selectedReverseSubject.item}
                onLocate={() => {
                  scrollToSubject(
                    selectedReverseSubject.ref.side,
                    selectedReverseSubject.ref.groupId,
                    selectedReverseSubject.ref.key,
                    activeTab,
                    setActiveTab
                  );
                }}
                onClear={() => {
                  setRevSubjectRefKey("");
                }}
              />
              <div className="flex flex-col gap-1 text-[11px] text-secondary-foreground font-semibold bg-muted/30 p-2 rounded-lg">
                <div>反算方向：<span className="text-foreground">{selectedReverseSubject.ref.side === 'revenue' ? '收入' : '投入'}</span></div>
                <div>反算模式：<span className="text-foreground">{reverseCalculationContext.mode === 'locked_total_structure' ? '结构反算' : '普通反算'}</span></div>
              </div>
            </div>
          ) : (
            <div className="text-xs text-secondary-foreground font-semibold bg-muted/20 p-3 rounded-lg border border-dashed border-border/60">
              当前未指定反算目标。请在下方{activeTab === 'revenue' ? '收入' : '投入'}科目列表中点击“设置角色”进行设置。
            </div>
          )}
          {selectedReverseSubject && (
            <div className="flex flex-col gap-2 mt-1">
              {reverseContextMessage && (
                <span className="text-[11px] leading-relaxed text-warning-foreground bg-warning-soft rounded-md px-2 py-1 font-semibold">
                  {reverseContextMessage}
                </span>
              )}
              {reverseCalculationContext.mode === "locked_total_structure" && (
                <span className="text-[11px] leading-relaxed text-primary bg-primary-soft rounded-md px-2 py-1 font-semibold">
                  当前为结构反算模式：{reverseCalculationContext.structure.sideLabel}含税总金额保持 {formatReverseCurrency(reverseCalculationContext.structure.totalInclAmount)} 不变。调整“{reverseCalculationContext.structure.targetDisplayName}”时，“{reverseCalculationContext.structure.balancingDisplayName}”将自动反向补差。
                  {state.cashflowModel === "model_e" && state.segmentValueMode === "amount" ? " 分板块现金流金额计划将同步更新。" : ""}
                </span>
              )}
              {reverseCalculationContext.mode === "normal" && (
                <span className="text-[11px] leading-relaxed text-secondary-foreground bg-muted/40 rounded-md px-2 py-1 font-semibold">
                  调整该科目金额后，{selectedReverseSubject.ref.side === 'revenue' ? '收入' : '投入'}总额将随反算结果变化。
                </span>
              )}
            </div>
          )}
        </div>
        <div className="flex gap-2 bg-background p-1 border border-border rounded-lg">
          <button className={`flex-1 py-1.5 text-sm font-semibold rounded-md ${revTargetType === 'margin' ? 'bg-primary text-primary-foreground shadow-sm' : 'text-secondary-foreground'}`} onClick={() => setRevTargetType('margin')}>目标毛利润率</button>
          <button className={`flex-1 py-1.5 text-sm font-semibold rounded-md ${revTargetType === 'npv_rate' ? 'bg-primary text-primary-foreground shadow-sm' : 'text-secondary-foreground'}`} onClick={() => setRevTargetType('npv_rate')}>目标净现值率</button>
        </div>
        <div className="flex flex-col gap-1">
          <label className="text-xs font-semibold text-secondary-foreground">目标值 (如0.15代表15%)</label>
          <input type="number" step="0.0001" className="bg-card border border-input px-3 py-2 rounded-md outline-none text-sm" value={revTargetValue} onChange={e => setRevTargetValue(e.target.value)} />
        </div>
        <button
          className="flex w-full items-center justify-center gap-2 bg-primary text-primary-foreground font-bold py-2 rounded-lg shadow-sm disabled:opacity-50 disabled:cursor-not-allowed"
          disabled={!selectedReverseSubject}
          onClick={executeReverseCalculation}
        >
          <AppIcon name="reverse" size={16} /> 智能反算
        </button>
      </div>
      <h3 className="font-bold text-foreground mb-4">采购甄选费测算</h3>
      <div className="bg-card border border-border p-4 rounded-xl flex flex-col gap-3">
        <div className="flex flex-col gap-1">
           <div className="flex items-center gap-1.5">
             <label className="text-xs font-semibold text-secondary-foreground">供应商报价 (元)</label>
             <button
               type="button"
               aria-label="固定供应商报价"
               aria-pressed={selectionFeeAnchor === 'quote'}
               title="固定供应商报价"
               onClick={() => setSelectionFeeAnchor('quote')}
               className={`flex h-5 w-5 items-center justify-center rounded-full transition-colors ${selectionFeeAnchor === 'quote' ? 'bg-primary-soft' : 'bg-muted hover:bg-muted/80'}`}
             >
               <span className={`h-2.5 w-2.5 rounded-full ${selectionFeeAnchor === 'quote' ? 'bg-primary' : 'bg-secondary-foreground/35'}`} />
             </button>
           </div>
           <input type="number" aria-label="供应商报价" value={selQuote} onChange={e => handleSelFeeChange('quote', e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md text-sm outline-none" />
        </div>
        <div className="flex flex-col gap-1">
           <label className="text-xs font-semibold text-secondary-foreground">代理服务费浮动 (+)</label>
           <input type="number" aria-label="代理服务费浮动" value={selMarkup} onChange={e => handleSelFeeChange('markup', e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md text-sm outline-none" />
        </div>
        <div className="flex flex-col gap-1">
           <label className="text-xs font-semibold text-secondary-foreground">测算甄选费 / 实际测算成本</label>
           <div className="flex gap-2">
             <input type="text" aria-label="测算甄选费" disabled value={selFee} className="bg-muted/50 border border-input px-3 py-2 rounded-md text-sm w-full text-secondary-foreground" />
             <input type="text" aria-label="实际测算成本" disabled value={selActualCost} className="bg-muted/50 border border-input px-3 py-2 rounded-md text-sm w-full text-secondary-foreground" />
           </div>
        </div>
        <div className="flex flex-col gap-1 mt-2 border-t border-border pt-3">
           <div className="flex items-center gap-1.5">
             <label className="text-xs font-semibold text-primary">甄选最高限价 (反向测算入口)</label>
             <button
               type="button"
               aria-label="固定甄选最高限价"
               aria-pressed={selectionFeeAnchor === 'limit'}
               title="固定甄选最高限价"
               onClick={() => setSelectionFeeAnchor('limit')}
               className={`flex h-5 w-5 items-center justify-center rounded-full transition-colors ${selectionFeeAnchor === 'limit' ? 'bg-primary-soft' : 'bg-muted hover:bg-muted/80'}`}
             >
               <span className={`h-2.5 w-2.5 rounded-full ${selectionFeeAnchor === 'limit' ? 'bg-primary' : 'bg-secondary-foreground/35'}`} />
             </button>
           </div>
           <input type="number" aria-label="甄选最高限价" value={selLimit} onChange={e => handleSelFeeChange('limit', e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md text-sm outline-none text-foreground font-bold" />
        </div>
        <button
          onClick={applySelectionLimit}
          disabled={!selLimit}
          className="mt-2 bg-primary hover:bg-primary/95 text-primary-foreground disabled:opacity-50 disabled:cursor-not-allowed font-bold py-2.5 rounded-lg shadow-sm hover:shadow-md transition-all active:scale-[0.98] w-full text-xs flex items-center justify-center gap-1.5"
        >
          <AppIcon name="download" size={14} /> 填入集成服务
        </button>
      </div>
    </>
  );

  const pageDirty = dirtyScopes.some(scope => ["lifecycle", "cashflow", "benefit-analysis", "template-forms"].includes(scope));
  const saveStatusLabel = isSaving
    ? "保存中..."
    : lastSaveError
      ? "保存失败，请重试"
      : pageDirty
        ? "● 未保存"
      : "已保存";
  const validIctOrigin = ictOrigin
    && ictOrigin.type === "intelligent_compute"
    && ictOrigin.workspaceId === workspaceId
    && ictOrigin.projectId === activeProject?.id
    && activeProject?.project_type === "intelligent_compute"
      ? ictOrigin
      : null;
  const isTabbedDocumentTemplate = activeTab === "generate"
    && (selectedTemplate.includes("会审") || selectedTemplate.includes("立项签批表") || selectedTemplate.includes("需求导入表") || selectedTemplate.includes("甄选结果签批表"));
  const showReverseCalculationTools = activeTab === 'revenue' || activeTab === 'cost';

  return (
    <div className="flex flex-col flex-1 animate-in fade-in duration-300 h-full overflow-hidden">
      <WorkspaceHeader
        moduleId="ict_lifecycle"
        title="ICT项目全生命周期"
        backLabel={validIctOrigin ? "返回智算测算" : "返回集市"}
        onBack={async () => {
          const canProceed = await confirmOrSave();
          if (!canProceed) return;
          if (entrySource === "project_board") {
            navigateTo("project_board", activeProjectId, activeSchemeId);
          } else if (validIctOrigin) {
            navigateTo("ai_compute_quote", activeProjectId, null, activeScenarioId);
          } else {
            navigateTo("hub");
          }
        }}
        onPathChange={() => loadTemplates()}
        inlineWorkspaceControl
        contextContent={
          <div className="flex min-w-0 items-center gap-2 text-xs">
            {activeProject ? (
              <>
                {validIctOrigin && (
                  <span className="shrink-0 rounded-full bg-primary-soft px-2 py-0.5 text-[10px] font-bold text-primary">
                    来源：智算项目 {validIctOrigin.projectName}
                  </span>
                )}
                <span className="truncate font-extrabold text-foreground max-w-[180px]">{activeProject.name}</span>
                <span className="truncate text-secondary-foreground max-w-[140px]">({activeProject.customer_name})</span>
                <span className={`shrink-0 rounded-full px-2 py-0.5 text-[10px] font-bold ${
                  activeProject.benefit_status === "normal"
                    ? "bg-success-soft text-success-foreground"
                    : activeProject.benefit_status === "outdated"
                      ? "bg-warning-soft text-warning-foreground"
                      : "bg-muted text-muted-foreground"
                }`}>
                  效益状态: {activeProject.benefit_status === "normal" ? "最新" : activeProject.benefit_status === "outdated" ? "已失效" : "未测算"}
                </span>
                <span className={`shrink-0 rounded-full px-2 py-0.5 text-[10px] font-bold ${
                  lastSaveError ? "bg-destructive-soft text-destructive" : pageDirty ? "bg-warning-soft text-warning-foreground" : "bg-success-soft text-success-foreground"
                }`}>
                  {saveStatusLabel}
                </span>
                <span className="min-w-0 truncate text-secondary-foreground">
                  当前方案: <span className="font-semibold text-primary">{pendingNewSchemeName || activeScheme?.name || "默认方案"}</span>
                  {activeSnapshot ? ` (v${activeSnapshot.version})` : ""}
                </span>
              </>
            ) : (
              <span className="rounded-full bg-secondary px-2 py-0.5 text-[10px] font-bold text-secondary-foreground">
                自由测算模式
              </span>
            )}
          </div>
        }
      />

      <div className={`grid flex-1 min-h-0 min-w-0 grid-cols-[260px_minmax(0,1fr)] overflow-hidden ${showReverseCalculationTools ? "xl:grid-cols-[260px_minmax(0,1fr)_300px]" : ""}`}>
        <div className="bg-muted p-6 overflow-y-auto flex flex-col gap-4 border-r border-border">
          <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70 mb-1">测算流程</h3>
          <div className="flex flex-col gap-1">
            <button className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${activeTab === 'basic' ? 'bg-primary-soft text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`} onClick={() => handleTabSwitch("basic")}><AppIcon name="project" size={18} /> 项目概况与参数</button>
            <button className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${activeTab === 'revenue' ? 'bg-primary-soft text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`} onClick={() => handleTabSwitch("revenue")}><AppIcon name="revenue" size={18} /> 收入侧测算</button>
            <button className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${activeTab === 'cost' ? 'bg-primary-soft text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`} onClick={() => handleTabSwitch("cost")}><AppIcon name="cost" size={18} /> 投入侧测算</button>
            <button className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${activeTab === 'cashflow' ? 'bg-primary-soft text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`} onClick={() => handleTabSwitch("cashflow")}><AppIcon name="cashflow" size={18} /> 10年现金流推演</button>
          </div>
          <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70 mt-6 pt-4 border-t border-border mb-2">一键生成全流程文档</h3>
          <div className="flex flex-col gap-2">
            {templates.length === 0 ? <span className="text-xs text-secondary-foreground px-4">未找到模板文件</span> :
              templates.map(t => {
                const isActive = selectedTemplate === t && activeTab === 'generate';
                return (
                  <button
                    key={t}
                    className={`relative overflow-hidden px-4 py-3 rounded-lg text-sm flex items-start gap-2.5 transition-all text-left border-b border-border/60 shadow-sm ${isActive ? 'bg-primary-soft text-primary font-bold border-primary/20 shadow-sm' : 'text-secondary-foreground font-semibold hover:bg-primary-soft hover:text-primary'}`}
                    onClick={() => handleTabSwitch("generate", t)}
                  >
                    {isActive && <div className="absolute left-0 top-0 bottom-0 w-1 bg-primary" />}
                    <AppIcon name={t.endsWith('.xlsx') ? "spreadsheet" : "document"} size={18} className="mt-0.5" />
                    <span className="whitespace-normal break-words leading-relaxed flex-1">
                      {t.replace('.docx', '').replace('.xlsx', '').replace('.pptx', '')}
                    </span>
                  </button>
                );
              })
            }
          </div>
        </div>

        <div className="min-w-0 p-4 sm:p-6 overflow-y-auto bg-background flex flex-col">
          {activeProject ? (
            <div className="bg-card border border-border rounded-xl p-4 mb-6 grid grid-cols-1 lg:grid-cols-[minmax(0,1fr)_auto] lg:items-center gap-4 shadow-sm animate-in slide-in-from-top duration-300">
              <div className="flex min-w-0 items-start sm:items-center gap-3">
                <div className="bg-primary/10 p-2.5 rounded-lg text-primary shrink-0">
                  <AppIcon name="project" size={20} />
                </div>
                <div className="min-w-0">
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="font-extrabold text-foreground text-sm break-words">{activeProject.name}</span>
                    <span className="text-xs text-secondary-foreground">({activeProject.customer_name})</span>
                    <span className={`text-[10px] px-2 py-0.5 rounded-full font-bold border ${
                      activeProject.benefit_status === 'normal'
                        ? 'bg-success-soft text-success-foreground border-success-soft/80'
                        : activeProject.benefit_status === 'outdated'
                        ? 'bg-warning-soft text-warning-foreground border-warning-soft/80'
                        : 'bg-muted text-muted-foreground border-border'
                    }`}>
                      效益状态: {activeProject.benefit_status === 'normal' ? '最新' : activeProject.benefit_status === 'outdated' ? '已失效' : '未测算'}
                    </span>
                  </div>
                  <div className="text-xs text-secondary-foreground mt-0.5 flex items-center gap-2 min-w-0">
                    {pendingNewSchemeName ? (
                      <span className="truncate">拟新建方案: <span className="font-semibold text-primary">{pendingNewSchemeName}</span> (未保存)</span>
                    ) : (
                      <>
                        {/* 阶段切换控件放在固定的最左侧，位置不随方案名长短变化；方案名放其右侧并按需截断。 */}
                        {activeScheme && (
                          <span className="inline-flex items-center gap-1.5 shrink-0">
                            <span className="inline-flex overflow-hidden rounded-md border border-border">
                              {SCHEME_STAGE_OPTIONS.map((option, idx) => {
                                const stageScheme = option.value === "pre_selection" ? preScheme : postScheme;
                                const exists = Boolean(stageScheme);
                                const isActive = exists && activeScheme.id === stageScheme!.id;
                                const title = exists
                                  ? isActive
                                    ? `当前即为${option.label}方案`
                                    : `切换到${option.label}方案：${stageScheme!.name}`
                                  : activeScheme.stage
                                    ? `基于当前方案派生一个${option.label}方案`
                                    : `将当前方案标注为${option.label}`;
                                return (
                                  <button
                                    key={option.value}
                                    type="button"
                                    onClick={() => handleStageButtonClick(option.value)}
                                    className={`px-2 py-0.5 text-[10px] font-semibold transition-all ${idx > 0 ? "border-l border-border" : ""} ${
                                      isActive
                                        ? option.chipClass
                                        : exists
                                          ? "bg-card text-secondary-foreground hover:bg-secondary"
                                          : "bg-muted/30 text-secondary-foreground/60 hover:bg-secondary"
                                    }`}
                                    title={title}
                                  >
                                    {option.label}{!exists && "（未生成）"}
                                  </button>
                                );
                              })}
                            </span>
                            {otherSchemes.length > 0 && (
                              <select
                                value=""
                                onChange={async (e) => {
                                  const sid = e.target.value;
                                  e.target.value = "";
                                  if (sid) await switchToScheme(sid);
                                }}
                                className="bg-card border border-border rounded-md px-1.5 py-0.5 text-[10px] font-semibold text-secondary-foreground outline-none focus:border-ring cursor-pointer"
                                title="其他方案（未标注阶段 / 同阶段历史方案）"
                              >
                                <option value="" disabled>更多方案 ▾</option>
                                {otherSchemes.map(s => (
                                  <option key={s.id} value={s.id}>
                                    {s.name}
                                    {getSchemeStageLabel(s.stage) !== "未标注" ? `（${getSchemeStageLabel(s.stage)}）` : ""}
                                    {s.id === activeScheme.id ? "（当前）" : ""}
                                  </option>
                                ))}
                              </select>
                            )}
                          </span>
                        )}
                        <span className="truncate min-w-0">当前方案: <span className="font-semibold text-primary">{activeScheme?.name || "默认方案"}</span> {activeSnapshot ? `(v${activeSnapshot.version})` : ''}</span>
                      </>
                    )}
                  </div>
                </div>
              </div>
              <div className="grid w-full min-w-0 grid-cols-1 gap-2 md:grid-cols-[minmax(0,1fr)_auto_auto] lg:w-auto lg:min-w-[420px] lg:max-w-[560px]">
                <select
                  onChange={async (e) => {
                    const pid = e.target.value;
                    const canProceed = await confirmOrSave();
                    if (!canProceed) return;
                    if (pid === "free") {
                      navigateTo("ict_lifecycle", null, null);
                    } else {
                      navigateTo("ict_lifecycle", pid, null);
                    }
                  }}
                  value={activeProject.id}
                  className="w-full min-w-0 bg-card border border-input px-3 py-1.5 rounded-lg text-xs outline-none focus:border-ring font-semibold text-foreground cursor-pointer"
                >
                  <option value="free">断开关联 (进入自由测算)</option>
                  {projects.map(p => (
                    <option key={p.id} value={p.id}>{p.name} ({p.customer_name})</option>
                  ))}
                </select>

                <button
                  id="save_benefit_btn"
                  onClick={handleSaveToCurrent}
                  className="bg-primary text-primary-foreground font-bold px-4 py-2 rounded-lg text-xs hover:bg-primary/90 transition-all shadow-sm flex items-center justify-center gap-1.5 active:scale-[0.98] whitespace-nowrap"
                >
                  <AppIcon name="save" size={14} /> 保存到当前项目
                </button>
                <button
                  id="save_as_new_benefit_btn"
                  onClick={() => {
                    setSaveAsSchemeStage(null);
                    setSaveAsSchemeName(activeScheme?.name ? `${activeScheme.name}_复本` : "新方案");
                    setShowSaveAsModal(true);
                  }}
                  className="bg-card border border-border text-foreground hover:bg-secondary font-bold px-4 py-2 rounded-lg text-xs transition-all shadow-sm flex items-center justify-center gap-1.5 active:scale-[0.98] whitespace-nowrap"
                >
                  <AppIcon name="copy" size={14} /> 另存为新方案
                </button>
              </div>
            </div>
          ) : (
            <div className="bg-card border border-border rounded-xl p-4 mb-6 grid grid-cols-1 lg:grid-cols-[minmax(0,1fr)_auto] lg:items-center gap-4 shadow-sm">
              <div className="flex min-w-0 items-start sm:items-center gap-3">
                <div className="bg-secondary p-2.5 rounded-lg text-primary shrink-0">
                  <AppIcon name="project" size={20} />
                </div>
                <div className="min-w-0">
                  <div className="font-extrabold text-foreground text-sm flex items-center gap-2">
                    <span className="whitespace-nowrap">自由测算模式</span>
                    <span className="text-[10px] bg-secondary text-secondary-foreground font-bold px-2 py-0.5 rounded-full">未绑定项目</span>
                  </div>
                  <div className="text-xs leading-relaxed text-secondary-foreground mt-0.5">你可以输入参数进行效益测算。如需保存，请在右侧选择关联一个项目：</div>
                </div>
              </div>
              <div className="grid w-full min-w-0 grid-cols-1 gap-2 sm:grid-cols-[minmax(0,1fr)_auto] lg:w-auto lg:min-w-[360px] lg:max-w-[520px]">
                <select
                  onChange={async (e) => {
                    const pid = e.target.value;
                    if (pid) {
                      const canProceed = await confirmOrSave();
                      if (!canProceed) return;
                      navigateTo("ict_lifecycle", pid, null);
                    }
                  }}
                  value=""
                  className="w-full min-w-0 bg-card border border-input px-3 py-1.5 rounded-lg text-xs outline-none focus:border-ring font-semibold text-foreground cursor-pointer"
                >
                  <option value="" disabled>-- 关联已有项目 --</option>
                  {projects.map(p => (
                    <option key={p.id} value={p.id}>{p.name} ({p.customer_name})</option>
                  ))}
                </select>
                <button
                  id="save_free_benefit_btn"
                  onClick={() => setShowSelectProjectModal(true)}
                  className="bg-primary text-primary-foreground font-bold px-4 py-2 rounded-lg text-xs hover:bg-primary/90 transition-all shadow-sm flex items-center justify-center gap-1.5 active:scale-[0.98] whitespace-nowrap"
                >
                  <AppIcon name="save" size={14} /> 保存当前测算
                </button>
              </div>
            </div>
          )}

          {activeTab === "basic" && (
            <IctBasicInfo state={state} calculations={calculations} />
          )}

          {activeTab === "revenue" && (
            <div>
              {renderCashflowCalculationSourceControl()}
              {renderBalanceControl("revenue")}
              {showReverseCalculationTools && (
                <div className="mb-6 flex flex-col xl:hidden">
                  {renderReverseCalculationTools()}
                </div>
              )}
              <div className="mb-4 text-xs text-primary bg-primary-soft p-3 rounded-lg border border-primary/20">
                <span className="inline-flex items-start gap-2"><AppIcon name="info" size={16} className="mt-0.5" /> <span>提示：在「CT收入」中填写的产品或专线含税收入，将会自动【1:1平过】填入对应的「CT投入」中。</span></span>
              </div>
              {renderTaxGroup("IT/移动云收入", 'revIt', revIt, ICT_SUBJECT_GROUPS.revIt)}
              {renderTaxGroup("CT收入", 'revCt', revCt, ICT_SUBJECT_GROUPS.revCt)}
              {renderTaxGroup("非IT/CT收入", 'revNonItCt', { item: revNonItCt }, ICT_SUBJECT_GROUPS.revNonItCt)}
            </div>
          )}

          {activeTab === "cost" && (
            <div>
              {renderCashflowCalculationSourceControl()}
              {renderBalanceControl("investment")}
              {showReverseCalculationTools && (
                <div className="mb-6 flex flex-col xl:hidden">
                  {renderReverseCalculationTools()}
                </div>
              )}
              {renderTaxGroup("IT/移动云投入", 'costIt', costIt, ICT_SUBJECT_GROUPS.costIt)}
              {renderTaxGroup("CT投入", 'costCt', costCt, ICT_SUBJECT_GROUPS.costCt)}
              {renderTaxGroup("非IT/CT投入 & 综合类成本", 'costMix', costMix, ICT_SUBJECT_GROUPS.costMix)}
            </div>
          )}

          {activeTab === "cashflow" && (
            <div className="flex flex-col gap-6">
              {renderCashflowCalculationSourceControl()}
              <IctCashflowTable state={state} calculations={calculations} />
              <IctProposalCashflowSummary state={state} calculations={calculations} />
            </div>
          )}

          <div className={`flex-col gap-6 ${activeTab === "generate" ? "flex" : "hidden"} ${isTabbedDocumentTemplate ? "" : "bg-card border border-border rounded-xl p-8 shadow-sm"}`}>
            {!isTabbedDocumentTemplate && <h3 className="text-lg font-bold text-foreground">即将生成：{selectedTemplate}</h3>}
            <TemplateForms
              selectedTemplate={selectedTemplate}
              projectData={{
                basic: {proj_name: projName, customer_name: customerName, project_years: projectYears},
                cost: { it: costIt, ct: costCt, mix: costMix },
                revenue: { it: revIt, ct: revCt, non_it_ct: revNonItCt },
                selectionFee: {
                  quote: selQuote,
                  markup: selMarkup,
                  actualCost: selActualCost,
                  amount: selFee,
                  limit: selLimit,
                  anchor: selectionFeeAnchor,
                },
              }}
              metrics={metrics}
              projectBackground={state.projectBackground}
              setProjectBackground={state.setProjectBackground}
              techItems={state.techItems}
              setTechItems={state.setTechItems}
              inqVendors={state.inqVendors}
              setInqVendors={state.setInqVendors}
              outputDir={activeProject?.folder_path || undefined}
              projectId={activeProject?.id || undefined}
              currentSchemeLabel={pendingNewSchemeName || activeScheme?.name || "默认方案"}
              fetchPreSelectionCostIt={fetchPreSelectionCostIt}
              preSchemeName={preScheme?.name}
            />
          </div>

          {!isTabbedDocumentTemplate && <IctMetricsDashboard metrics={metrics} />}
        </div>

        {showReverseCalculationTools && (
          <div className="hidden min-w-0 bg-card border-l border-border p-6 xl:flex flex-col overflow-y-auto animate-in slide-in-from-right duration-200">
            {renderReverseCalculationTools()}
          </div>
        )}
      </div>

      {showReconciliationModal && (
        <div className="fixed inset-0 z-50 bg-background/80 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in">
          <div className="bg-card border border-red-500/30 rounded-xl shadow-xl w-full max-w-2xl overflow-hidden flex flex-col">
            <div className="bg-red-500/10 border-b border-red-500/20 px-6 py-4 flex items-center gap-3">
              <AppIcon name="warning" size={24} className="text-red-600" />
              <div>
                <h2 className="font-bold text-red-600 text-lg">0 容差财务核算拦截</h2>
                <p className="text-xs text-red-600/80 mt-0.5">检测到税前/税后金额转换存在微小尾差，系统已拦截保存操作。</p>
              </div>
            </div>

            <div className="p-6 overflow-y-auto max-h-[60vh] flex flex-col gap-4 bg-muted/30">
              {reconciliationErrors.map((err: ValidationReport, i: number) => (
                <div key={i} className="bg-background border border-red-200 p-4 rounded-lg flex flex-col gap-2">
                  <div className="flex justify-between items-center">
                    <span className="font-bold text-sm bg-red-100 text-red-800 px-2 py-0.5 rounded">
                      {err.side === 'income' ? '收入侧' : '投入侧'} - {err.taxRate}% 税率组
                    </span>
                    <span className="text-xs font-mono bg-muted px-2 py-0.5 rounded">{err.key}</span>
                  </div>
                  <div className="grid grid-cols-3 gap-4 mt-2 text-sm">
                    <div className="flex flex-col gap-1">
                      <span className="text-secondary-foreground text-xs">{err.field === 'incl' ? '录入含税' : '录入不含税'}</span>
                      <span className="font-bold">{err.actualExcl} 元</span>
                    </div>
                    <div className="flex flex-col gap-1">
                      <span className="text-secondary-foreground text-xs">{err.field === 'incl' ? '财务口径含税（不含税×(1+税率)）' : '预期不含税'}</span>
                      <span className="font-bold text-primary">{err.expectedExcl} 元</span>
                    </div>
                    <div className="flex flex-col gap-1">
                      <span className="text-secondary-foreground text-xs">尾差</span>
                      <span className="font-bold text-red-500">{err.difference} 元</span>
                    </div>
                  </div>
                </div>
              ))}
            </div>

            <div className="border-t border-border p-4 bg-background flex justify-end gap-3">
              {Math.abs(Number(currentTotalDifference)) <= 0.10 && (
                <button
                  onClick={() => setShowConfirmIgnore(true)}
                  className="px-6 py-2 border border-border hover:bg-muted text-secondary-foreground font-bold rounded-md shadow-sm transition-colors text-sm"
                >
                  忽略微小尾差，继续提交
                </button>
              )}
              <button
                onClick={() => setShowReconciliationModal(false)}
                className="px-6 py-2 bg-red-500 hover:bg-red-600 text-white font-bold rounded-md shadow-sm transition-colors text-sm"
              >
                返回手工平账
              </button>
            </div>
          </div>
        </div>
      )}

      {showConfirmIgnore && pendingTab && (
        <div className="fixed inset-0 z-[60] bg-background/80 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in">
          <div className="bg-card border border-border rounded-xl shadow-xl w-full max-w-md overflow-hidden flex flex-col">
            <div className="px-6 py-4 border-b border-border bg-yellow-500/10 flex items-center gap-2">
              <AppIcon name="warning" size={20} className="text-yellow-600" />
              <h2 className="font-bold text-yellow-700 text-base">确认忽略误差？</h2>
            </div>
            <div className="p-6">
              <p className="text-sm text-secondary-foreground leading-relaxed">系统检测到存在微小尾差，强制提交可能需要在后续向财务提供纸质/邮件说明。是否确认忽略误差并继续？</p>
            </div>
            <div className="border-t border-border p-4 bg-background flex justify-end gap-3">
              <button onClick={() => setShowConfirmIgnore(false)} className="px-4 py-2 border border-border hover:bg-muted font-bold rounded-md transition-colors text-sm">取消</button>
              <button onClick={() => {
                setShowConfirmIgnore(false);
                setShowReconciliationModal(false);
                handleTabSwitch(pendingTab.tab, pendingTab.template, true);
              }} className="px-4 py-2 bg-primary hover:bg-primary/90 text-primary-foreground font-bold rounded-md transition-colors text-sm">确认忽略并继续</button>
            </div>
          </div>
        </div>
      )}

      {showSaveAsModal && (
        <div className="fixed inset-0 z-[60] bg-background/80 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in">
          <form
            onSubmit={handleSaveAsNew}
            className="bg-card border border-border rounded-xl shadow-xl w-full max-w-sm overflow-hidden"
          >
            <div className="px-6 py-4 border-b border-border bg-muted/30 flex items-center justify-between">
              <h4 className="font-bold text-sm text-foreground">
                {saveAsSchemeStage ? `派生${getSchemeStageLabel(saveAsSchemeStage)}方案` : "另存为新方案"}
              </h4>
              <button
                type="button"
                onClick={() => { setShowSaveAsModal(false); setSaveAsSchemeStage(null); }}
                className="text-secondary-foreground hover:bg-secondary p-1 rounded-md"
              >
                <AppIcon name="close" size={14} />
              </button>
            </div>
            <div className="p-6">
              {saveAsSchemeStage && (
                <p className="text-xs text-secondary-foreground mb-3 leading-relaxed">
                  将以当前方案的最新数据为起点，创建一个标注为
                  <span className="font-semibold text-primary">「{getSchemeStageLabel(saveAsSchemeStage)}」</span>
                  的新方案。保存后可在此基础上按甄选实际报价调整。
                </p>
              )}
              <label className="text-xs font-semibold text-secondary-foreground block mb-1.5">方案名称 <span className="text-red-500">*</span></label>
              <input
                id="save_as_new_scheme_name_input"
                type="text"
                required
                placeholder="例如：方案 B、第二轮测算"
                value={saveAsSchemeName}
                onChange={(e) => {
                  setSaveAsSchemeName(e.target.value);
                  markDirty("benefit-analysis");
                }}
                className="bg-card border border-input px-3 py-2 rounded-lg text-xs outline-none focus:border-ring w-full"
              />
            </div>
            <div className="border-t border-border p-3 bg-muted/10 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => { setShowSaveAsModal(false); setSaveAsSchemeStage(null); }}
                className="px-3 py-1.5 border border-border hover:bg-secondary rounded-lg text-xs font-semibold text-secondary-foreground transition-all"
              >
                取消
              </button>
              <button
                id="submit_save_as_new_scheme_btn"
                type="submit"
                disabled={!saveAsSchemeName.trim()}
                className="px-3 py-1.5 bg-primary hover:bg-primary/90 disabled:opacity-50 text-white font-bold rounded-lg text-xs transition-all"
              >
                确认
              </button>
            </div>
          </form>
        </div>
      )}

      {showSelectProjectModal && (
        <div className="fixed inset-0 z-[60] bg-background/80 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in">
          <div className="bg-card border border-border rounded-xl shadow-xl w-full max-w-sm overflow-hidden">
            <div className="px-6 py-4 border-b border-border bg-muted/30 flex items-center justify-between">
              <h4 className="font-bold text-sm text-foreground">选择要保存到的项目</h4>
              <button
                type="button"
                onClick={() => setShowSelectProjectModal(false)}
                className="text-secondary-foreground hover:bg-secondary p-1 rounded-md"
              >
                <AppIcon name="close" size={14} />
              </button>
            </div>
            <div className="p-6">
              <label className="text-xs font-semibold text-secondary-foreground block mb-1.5">选择项目 <span className="text-red-500">*</span></label>
              <select
                id="save_target_project_select"
                defaultValue=""
                onChange={(e) => {
                  const val = e.target.value;
                  if (val) {
                    setShowSelectProjectModal(false);
                    handleSaveToSelectedProject(val);
                  }
                }}
                className="bg-card border border-input px-3 py-2 rounded-lg text-xs outline-none focus:border-ring w-full cursor-pointer font-semibold"
              >
                <option value="" disabled>-- 请选择项目 --</option>
                {projects.map(p => (
                  <option key={p.id} value={p.id}>{p.name} ({p.customer_name})</option>
                ))}
              </select>
            </div>
            <div className="border-t border-border p-3 bg-muted/10 flex justify-end">
              <button
                onClick={() => setShowSelectProjectModal(false)}
                className="px-3 py-1.5 border border-border hover:bg-secondary rounded-lg text-xs font-semibold text-secondary-foreground transition-all"
              >
                取消
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
