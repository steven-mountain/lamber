import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import AppIcon, { type AppIconName } from "../../components/icons/AppIcon";
import { Button } from "../../components/ui/button";
import { Input } from "../../components/ui/input";
import { useNavigationStore } from "../../store/useNavigationStore";
import { useProjectStore } from "../../store/useProjectStore";
import { useWorkspaceStore } from "../../store/useWorkspaceStore";
import { projectService, type IctResult, type Project } from "../../utils/projectService";
import { domainSaveService } from "../../services/domainSaveService";
import { intelligentComputeService } from "../../services/intelligentComputeService";
import { finalizeIctInputWithFundingPlans } from "../../lib/ictCalculationInput";
import {
  buildAiComputeQuoteOutput,
  buildAiComputeQuoteOutputFundingPlans,
  runAiComputeQuoteSensitivity,
  summarizeQuote,
} from "./calculations";
import BenefitConclusionSidebar from "./BenefitConclusionSidebar";
import CalculationItemCard from "./CalculationItemCard";
import { getFormulaParameterReferences } from "./formulaEngine";
import {
  getAiComputeDiscountRatePercent,
  getAiComputeProjectCycleYears,
  isAiComputeDiscountRateParameter,
  isAiComputeProjectCycleParameter,
  isAiComputeStableIctParameter,
} from "./fundingPlans";
import {
  buildAiComputeIctExportPayloads,
  buildAiComputeAutoSyncPreview,
  buildIntelligentComputeAggregatePreview,
  validateIntelligentComputeSources,
  type AiComputeIctExportPreview,
} from "./ictExport";
import { PARAMETER_GROUP_IDS } from "./parameterLayout";
import { createH200Blueprint } from "./presets";
import {
  buildAmountSourcePackage,
  getDefaultImportedAmountSourceName,
  normalizeAmountSourcePackage,
  sanitizeAmountSourceFileName,
  type AiComputeAmountSourcePackage,
} from "./amountSourceExchange";
import {
  buildBlueprintForAmountSource,
  canDeleteIntelligentAmountSource,
  type CreateAmountSourceBaseMode,
  type CreateAmountSourceRequest,
  getDefaultCreateAmountSourceBaseMode,
  isH200BaselineAmountSource,
  useAiComputeQuoteStore,
} from "./store";
import { buildIntelligentComputeSyncLock, getAiComputeSyncFingerprint } from "./ictSync";
import type {
  AiComputeLineItemFundingPlanMode,
  AiComputeQuoteLineItem,
	  AiComputeQuoteParameter,
	  AiComputeQuoteParameterCategory,
	  AiComputeQuoteParameterGroup,
	  AiComputeQuoteSide,
	  IntelligentAmountSource,
	} from "./types";

type QuoteWorkspaceTab = "parameters" | "revenue" | "cost" | "sensitivity";
type ParameterFilterMode = "key" | "modified" | "sensitive" | "all";
const PARAMETER_DRAG_TYPE = "application/x-lamber-ai-parameter";
const PARAMETER_GROUP_DRAG_TYPE = "application/x-lamber-ai-parameter-group";

const PARAMETER_IMPACT_BY_KEY: Record<string, string> = {
  device_count: "直接影响收入规模、投入规模与整体项目体量。",
  years: "影响累计收入、持续成本与动态回收期。",
  discount_rate: "直接覆盖 ICT 项目折现率，影响 NPV、净现值率与收益评价。",
  capital_rate: "影响 NPV 与回收期，建议结合融资方案设置。",
  gpu_service_price: "影响收入、毛利率与 NPV，建议重点关注。",
  cabinet_revenue_price: "影响机柜收入和项目毛利率。",
  cabinet_cost_price: "影响机柜持续成本与毛利率。",
  power_kw_per_device: "同时影响机柜收入与机柜成本。",
  bandwidth_revenue_price: "影响专线收入、毛利率与 NPV。",
  bandwidth_cost_price: "影响带宽成本和项目毛利率。",
  bandwidth_per_device: "影响带宽收入与成本规模。",
  machine_price: "影响主要建设投入、毛利率与 NPV。",
  maintenance_price: "影响服务期内持续运营成本。",
  network_price: "影响组网建设投入与项目毛利率。",
};

const H200_DEFAULT_PARAMETERS = createH200Blueprint().parameters;
const H200_DEFAULT_PARAMETER_BY_ID = new Map(H200_DEFAULT_PARAMETERS.map(parameter => [parameter.id, parameter]));
const H200_DEFAULT_PARAMETER_BY_KEY = new Map(H200_DEFAULT_PARAMETERS.map(parameter => [parameter.key, parameter]));

function createId(prefix: string) {
  const suffix = typeof crypto !== "undefined" && "randomUUID" in crypto
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  return `${prefix}-${suffix}`;
}

function toNumber(value: string) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function getCompactInputWidth(value: string) {
  const text = value || "单位";
  const contentWidth = Array.from(text).reduce((width, character) => (
    width + (/[\u2E80-\u9FFF]/.test(character) ? 14 : 8)
  ), 0);
  return Math.min(128, Math.max(52, contentWidth + 18));
}

function formatParameterValue(value: number) {
  return value.toLocaleString("zh-CN", {
    maximumFractionDigits: 4,
  });
}

function getParameterGroupIcon(groupId: string): AppIconName {
  if (groupId === PARAMETER_GROUP_IDS.scale) return "project";
  if (groupId === PARAMETER_GROUP_IDS.pricing) return "revenue";
  if (groupId === PARAMETER_GROUP_IDS.investment) return "cost";
  if (groupId === PARAMETER_GROUP_IDS.operations) return "settings";
  if (groupId === PARAMETER_GROUP_IDS.finance) return "npv";
  return "parameters";
}

function getLegacyParameterCategory(groupId: string): AiComputeQuoteParameterCategory {
  if (groupId === PARAMETER_GROUP_IDS.scale) return "scale";
  if (groupId === PARAMETER_GROUP_IDS.pricing) return "price";
  if (groupId === PARAMETER_GROUP_IDS.investment) return "cost";
  if (groupId === PARAMETER_GROUP_IDS.operations) return "technical";
  if (groupId === PARAMETER_GROUP_IDS.finance) return "finance";
  return "custom";
}

function isParameterModified(parameter: AiComputeQuoteParameter) {
  const baseline = H200_DEFAULT_PARAMETER_BY_ID.get(parameter.id)
    || H200_DEFAULT_PARAMETER_BY_KEY.get(parameter.key);
  if (!baseline) return true;
  return baseline.name !== parameter.name
    || baseline.key !== parameter.key
    || baseline.value !== parameter.value
    || (baseline.unit || "") !== (parameter.unit || "")
    || baseline.category !== parameter.category
    || baseline.groupId !== parameter.groupId
    || baseline.isKey !== parameter.isKey
    || (baseline.sensitivityEnabled !== false) !== (parameter.sensitivityEnabled !== false);
}

function getParameterImpact(parameter: AiComputeQuoteParameter, usage: { revenue: string[]; cost: string[] }) {
  if (PARAMETER_IMPACT_BY_KEY[parameter.key]) return PARAMETER_IMPACT_BY_KEY[parameter.key];
  const impacts = [
    usage.revenue.length > 0 ? "收入" : "",
    usage.cost.length > 0 ? "成本" : "",
  ].filter(Boolean);
  return impacts.length > 0
    ? `影响${impacts.join("与")}，调整后会实时进入正式 ICT 指标预览。`
    : "当前尚未被公式引用，可用于后续扩展报价逻辑。";
}

function getSyncStatusLabel(status: "idle" | "syncing" | "synced" | "error" | "conflict") {
  if (status === "syncing") return "同步中";
  if (status === "synced") return "已同步";
  if (status === "conflict") return "存在冲突";
  if (status === "error") return "同步失败";
  return "待同步";
}

function isIntelligentComputeVersionConflict(error: unknown) {
  const message = String(error);
  return [
    "IntelligentComputeSyncRevisionConflict",
    "IntelligentAmountSourceSetConflict",
    "IntelligentAmountSourceVersionConflict",
  ].some(marker => message.includes(marker));
}

function getSavedStatusLabel(isDirty: boolean, lastSavedAt: string | null) {
  if (isDirty) return "有未保存修改";
  if (!lastSavedAt) return "尚未保存";
  return `已保存 ${new Date(lastSavedAt).toLocaleTimeString("zh-CN", {
    hour: "2-digit",
    minute: "2-digit",
  })}`;
}

function formatWan(value: number) {
  return (value / 10000).toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function formatNumber(value: number, digits = 2) {
  return value.toLocaleString("zh-CN", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

export default function AiComputeQuoteView() {
  const navigateTo = useNavigationStore(state => state.navigateTo);
  const openIctFromIntelligentCompute = useNavigationStore(state => state.openIctFromIntelligentCompute);
  const activeProjectId = useNavigationStore(state => state.activeProjectId);
  const currentProject = useProjectStore(state => state.currentProject);
  const workspaceId = useWorkspaceStore(state => state.workspaceId);
  const store = useAiComputeQuoteStore();
  const loadQuote = store.load;
  const [projects, setProjects] = useState<Project[]>([]);
	  const [showOutput, setShowOutput] = useState(false);
	  const [showSyncDetails, setShowSyncDetails] = useState(false);
	  const [showAmountSourceManager, setShowAmountSourceManager] = useState(false);
	  const [ictExportPreview, setIctExportPreview] = useState<AiComputeIctExportPreview | null>(null);
	  const [ictResult, setIctResult] = useState<IctResult | null>(null);
	  const [syncStatus, setSyncStatus] = useState<"idle" | "syncing" | "synced" | "error" | "conflict">("idle");
	  const [ictExportError, setIctExportError] = useState<string | null>(null);
	  const syncRequestRef = useRef(0);
	  const autoSyncTimerRef = useRef<number | null>(null);
	  const lastAutoSyncFingerprintRef = useRef<string | null>(null);
	  const hasLoadedAutoSyncBaselineRef = useRef(false);
  const [activeTab, setActiveTab] = useState<QuoteWorkspaceTab>("parameters");
  const [moreMenuOpen, setMoreMenuOpen] = useState(false);
  const [showSaveAsModal, setShowSaveAsModal] = useState(false);
  const [pendingImportPackage, setPendingImportPackage] = useState<AiComputeAmountSourcePackage | null>(null);
  const [sensitivityParameterId, setSensitivityParameterId] = useState("gpu-service-price");
  const [sensitivityMin, setSensitivityMin] = useState(72000);
  const [sensitivityMax, setSensitivityMax] = useState(108000);
  const [sensitivityStep, setSensitivityStep] = useState(9000);
  const [ictSensitivityRows, setIctSensitivityRows] = useState<Array<
    ReturnType<typeof runAiComputeQuoteSensitivity>[number] & { ictResult?: IctResult }
  >>([]);
  const sensitivityRequestRef = useRef(0);

  const projectId = activeProjectId || currentProject?.id || null;
  const activeProject = projects.find(project => project.id === projectId)
    || (currentProject?.id === projectId ? currentProject : null);
  const isIntelligentProject = activeProject?.project_type === "intelligent_compute";

  useEffect(() => {
    projectService.listWorkspaceProjects()
      .then(rows => setProjects(rows.map(row => row.project)))
      .catch(() => setProjects([]));
  }, []);

	  useEffect(() => {
	    if (!projectId) {
	      void loadQuote(null);
    } else if (isIntelligentProject) {
      void loadQuote(projectId);
    } else {
	      void loadQuote(null);
	    }
	  }, [isIntelligentProject, loadQuote, projectId]);

	  useEffect(() => {
	    hasLoadedAutoSyncBaselineRef.current = false;
	    lastAutoSyncFingerprintRef.current = null;
	    if (autoSyncTimerRef.current) {
	      window.clearTimeout(autoSyncTimerRef.current);
	      autoSyncTimerRef.current = null;
	    }
	  }, [projectId]);

  useEffect(() => {
    if (!projectId || !store.isDirty || store.isLoading || store.isSaving) return;
    const timer = window.setTimeout(() => {
      const latest = useAiComputeQuoteStore.getState();
      if (latest.projectId === projectId && latest.isDirty && !latest.isLoading && !latest.isSaving) {
        void latest.save(projectId);
      }
    }, 650);
    return () => window.clearTimeout(timer);
  }, [projectId, store.blueprint, store.isDirty, store.isLoading, store.isSaving]);

  const summary = useMemo(() => summarizeQuote(store.blueprint), [store.blueprint]);
  const outputPackage = useMemo(() => buildAiComputeQuoteOutput(store.blueprint), [store.blueprint]);
  const majorRevenue = useMemo(
    () => [...store.blueprint.revenueItems]
      .filter(item => item.enabled)
      .sort((left, right) => right.amountInclTax - left.amountInclTax)[0]?.name || "--",
    [store.blueprint.revenueItems],
  );
  const majorCost = useMemo(
    () => [...store.blueprint.costItems]
      .filter(item => item.enabled)
      .sort((left, right) => right.amountInclTax - left.amountInclTax)[0]?.name || "--",
    [store.blueprint.costItems],
  );
  const sensitivityRows = useMemo(() => runAiComputeQuoteSensitivity(store.blueprint, {
    parameterId: sensitivityParameterId,
    min: sensitivityMin,
    max: sensitivityMax,
    step: sensitivityStep,
  }), [sensitivityMax, sensitivityMin, sensitivityParameterId, sensitivityStep, store.blueprint]);
	  const sensitivityBlueprint = store.blueprint;
	  const projectCycleParameter = store.blueprint.parameters.find(isAiComputeProjectCycleParameter);
	  const discountRateParameter = store.blueprint.parameters.find(isAiComputeDiscountRateParameter);

	  const buildAutoSyncFingerprint = useCallback(() => {
	    const latest = useAiComputeQuoteStore.getState();
	    if (!latest.projectState || !latest.projectId) return null;
	    const sourceFingerprint = (blueprint: ReturnType<typeof buildBlueprintForAmountSource>) =>
	      getAiComputeSyncFingerprint({
	        ...blueprint,
	        name: "",
	        description: "",
	      });
	    const selectedSource = latest.amountSources.find(source => source.id === latest.activeAmountSourceId)
	      || latest.amountSources[0];
	    if (!selectedSource) return null;
	    const selectedBlueprint = selectedSource.id === latest.activeAmountSourceId
	      ? latest.blueprint
	      : buildBlueprintForAmountSource(selectedSource, latest.projectState);
	    return JSON.stringify({
	      projectId: latest.projectId,
	      projectYears: getAiComputeProjectCycleYears(selectedBlueprint.parameters),
	      discountRate: getAiComputeDiscountRatePercent(selectedBlueprint.parameters),
	      source: {
	        id: selectedSource.id,
	        fingerprint: sourceFingerprint(selectedBlueprint),
	      },
	    });
	  }, []);

  useEffect(() => {
    const requestId = ++sensitivityRequestRef.current;
    if (!projectId || sensitivityRows.length === 0) {
      setIctSensitivityRows([]);
      return;
    }
    const timer = window.setTimeout(async () => {
      try {
        const fullState = await domainSaveService.loadProjectFullState(projectId);
        const inputs = sensitivityRows.map(row => {
          const candidate = {
            ...sensitivityBlueprint,
            parameters: sensitivityBlueprint.parameters.map(parameter =>
              parameter.id === sensitivityParameterId
                ? { ...parameter, value: row.parameterValue }
                : parameter
            ),
          };
          const preview = buildAiComputeAutoSyncPreview(candidate, fullState, projectId);
          const payloads = buildAiComputeIctExportPayloads(preview, fullState);
          return finalizeIctInputWithFundingPlans(
            payloads.lifecycleState.inputPayloadJson,
            payloads.cashflowState.assumptionsJson.subjectFundingPlans,
          ).input;
        });
        const results = await domainSaveService.calculateIctBenefitBatch(inputs);
        if (requestId !== sensitivityRequestRef.current) return;
        setIctSensitivityRows(sensitivityRows.map((row, index) => ({
          ...row,
          ictResult: results[index],
        })));
      } catch {
        if (requestId === sensitivityRequestRef.current) setIctSensitivityRows(sensitivityRows);
      }
    }, 500);
    return () => window.clearTimeout(timer);
  }, [projectId, sensitivityBlueprint, sensitivityParameterId, sensitivityRows]);

	  const renameAmountSource = useCallback(async (name: string) => {
	    if (!projectId || !name.trim()) return false;
	    useAiComputeQuoteStore.getState().renameBlueprint(name);
	    return useAiComputeQuoteStore.getState().save(projectId);
	  }, [projectId]);

	  const clearSyncPreview = useCallback(() => {
	    setIctExportPreview(null);
	    setIctExportError(null);
	    setShowSyncDetails(false);
	    setSyncStatus(status => status === "syncing" ? status : "idle");
	  }, []);

  const saveCurrentAsAmountSource = useCallback(async (name: string) => {
    const ok = await useAiComputeQuoteStore.getState().saveCurrentAsAmountSource(name);
    if (ok) {
      clearSyncPreview();
      setShowSaveAsModal(false);
    }
    return ok;
  }, [clearSyncPreview]);

  const exportCurrentAmountSource = useCallback(async () => {
    const latest = useAiComputeQuoteStore.getState();
    if (!projectId || !latest.projectState) {
      setIctExportError("请先打开智算项目后再导出金额来源。");
      return;
    }
    const source = latest.amountSources.find(item => item.id === latest.activeAmountSourceId);
    if (!source) {
      setIctExportError("当前金额来源不存在，无法导出。");
      return;
    }
    try {
      const pkg = buildAmountSourcePackage(source, latest.blueprint, latest.projectState);
      const exportedPath = await intelligentComputeService.exportAmountSourcePackage(
        projectId,
        pkg,
        sanitizeAmountSourceFileName(`智算金额来源-${pkg.source.name}`),
      );
      if (exportedPath) {
        window.alert(`金额来源已导出：${exportedPath}`);
      }
      setIctExportError(null);
    } catch (error) {
      setIctExportError(`金额来源导出失败：${String(error)}`);
    }
  }, [projectId]);

  const selectImportAmountSourcePackage = useCallback(async () => {
    try {
      const raw = await intelligentComputeService.selectAndReadAmountSourcePackage();
      if (!raw) return;
      setPendingImportPackage(normalizeAmountSourcePackage(raw));
      setIctExportError(null);
    } catch (error) {
      setIctExportError(`金额来源导入文件读取失败：${String(error)}`);
    }
  }, []);

  const confirmImportAmountSourcePackage = useCallback(async (
    pkg: AiComputeAmountSourcePackage,
    options: { name: string; applyProjectSettings: boolean },
  ) => {
    const ok = await useAiComputeQuoteStore.getState().importAmountSourcePackage(pkg, options);
    if (ok) {
      clearSyncPreview();
      setPendingImportPackage(null);
    }
    return ok;
  }, [clearSyncPreview]);

	  const syncToIct = useCallback(async (options: {
	    confirmed?: boolean;
	    silent?: boolean;
	    autoFingerprint?: string;
	  } = {}) => {
	    const { confirmed = false, silent = false, autoFingerprint } = options;
	    const currentStore = useAiComputeQuoteStore.getState();
	    if (!projectId || !activeProject || !workspaceId || !currentStore.projectState) return false;
	    const requestId = ++syncRequestRef.current;
	    setSyncStatus("syncing");
	    setIctExportError(null);
	    try {
	      if (currentStore.isDirty && !(await currentStore.save(projectId))) {
	        throw new Error("请先完成智算金额来源保存");
	      }
	      const latest = useAiComputeQuoteStore.getState();
      if (!latest.projectState) throw new Error("智算项目状态未加载");
      const selectedSource = latest.amountSources.find(source => source.id === latest.activeAmountSourceId)
        || latest.amountSources[0];
      if (!selectedSource) {
        throw new Error("请选择一个智算金额来源后才能同步");
      }
      const selectedBlueprint = selectedSource.id === latest.activeAmountSourceId
        ? latest.blueprint
        : buildBlueprintForAmountSource(selectedSource, latest.projectState);
      const selectedSources = [{
        sourceId: selectedSource.id,
        sourceName: selectedSource.name,
        blueprint: selectedBlueprint,
      }];
      const validation = validateIntelligentComputeSources(selectedSources);
      if (!validation.valid) {
        throw new Error(validation.issues.slice(0, 6).join("\n"));
      }
      const fullState = await domainSaveService.loadProjectFullState(projectId);
      const preview = buildIntelligentComputeAggregatePreview(
        selectedSources,
        fullState,
        projectId,
        latest.projectState.controlledSubjects,
      );
      if (preview.rows.length === 0) {
        throw new Error("没有可同步到 ICT 的映射科目");
      }
	      setIctExportPreview(preview);
	      if (!confirmed) {
	        if (!silent) setShowSyncDetails(true);
	        setSyncStatus("idle");
	        return true;
	      }
      const payloads = buildAiComputeIctExportPayloads(preview, fullState);
      const finalized = finalizeIctInputWithFundingPlans(
        payloads.lifecycleState.inputPayloadJson,
        payloads.cashflowState.assumptionsJson.subjectFundingPlans,
      );
      if (!finalized.coverage.valid) {
        throw new Error(finalized.coverage.issues[0]?.message || "ICT 科目资金计划校验失败");
      }
      const controlledSubjects = Object.fromEntries(
        preview.rows
          .map(row => [`${row.side}:${row.ictSubjectCode}`, {
            side: row.side,
            ictSubjectCode: row.ictSubjectCode,
            amountInclTax: row.writtenAmount,
            taxRate: row.taxRate,
            yearlyAmounts: row.yearlyAmounts,
            sourceLineItemIds: row.sourceLineItemIds,
          }]),
      );
      const syncLock = buildIntelligentComputeSyncLock(latest.projectState, latest.amountSources);
      const result = await domainSaveService.syncIntelligentComputeToIct(projectId, {
        expectedSyncRevision: syncLock.expectedSyncRevision,
        sourceVersions: syncLock.sourceVersions,
        controlledSubjects,
        lifecycleState: payloads.lifecycleState,
        cashflowState: payloads.cashflowState,
        calculationInput: finalized.input,
      });
	      if (requestId !== syncRequestRef.current) return;
	      useAiComputeQuoteStore.getState().applyProjectState(result.projectState);
	      setIctExportPreview(null);
	      setIctResult(result.ictResult);
	      setSyncStatus("synced");
	      setShowSyncDetails(false);
	      setIctExportError(null);
	      lastAutoSyncFingerprintRef.current = autoFingerprint || buildAutoSyncFingerprint();
	      return true;
	    } catch (error) {
	      if (requestId !== syncRequestRef.current) return;
	      if (isIntelligentComputeVersionConflict(error)) {
	        setSyncStatus("conflict");
	        setShowSyncDetails(true);
	        setIctExportError("同步冲突：智算或 ICT 状态已更新。请重新加载冲突并使用最新预览执行智算完全覆盖 ICT。");
	        return false;
	      }
	      if (autoFingerprint) lastAutoSyncFingerprintRef.current = autoFingerprint;
	      setSyncStatus("error");
	      setIctExportError(`同步失败：${String(error)}`);
	      return false;
	    }
	  }, [activeProject, buildAutoSyncFingerprint, projectId, workspaceId]);

	  const reloadConflictAndOverwrite = useCallback(async () => {
	    if (!projectId) return;
	    setSyncStatus("syncing");
	    setIctExportError(null);
	    await useAiComputeQuoteStore.getState().load(projectId);
	    if (!useAiComputeQuoteStore.getState().projectState) {
	      setSyncStatus("error");
	      setIctExportError("重新加载冲突状态失败，请返回项目后重新打开智算测算。");
	      return;
	    }
	    await syncToIct({ confirmed: true });
	  }, [projectId, syncToIct]);

	  useEffect(() => {
	    if (!projectId || !isIntelligentProject || store.isLoading || store.isSaving) return;
	    if (syncStatus === "syncing") return;
	    const fingerprint = buildAutoSyncFingerprint();
	    if (!fingerprint) return;
	    if (!hasLoadedAutoSyncBaselineRef.current) {
	      hasLoadedAutoSyncBaselineRef.current = true;
	      lastAutoSyncFingerprintRef.current = fingerprint;
	      return;
	    }
	    if (fingerprint === lastAutoSyncFingerprintRef.current) return;
	    if (autoSyncTimerRef.current) window.clearTimeout(autoSyncTimerRef.current);
	    autoSyncTimerRef.current = window.setTimeout(() => {
	      void syncToIct({
	        confirmed: true,
	        silent: true,
	        autoFingerprint: fingerprint,
	      });
	    }, 1200);
	    return () => {
	      if (autoSyncTimerRef.current) {
	        window.clearTimeout(autoSyncTimerRef.current);
	        autoSyncTimerRef.current = null;
	      }
	    };
	  }, [
	    buildAutoSyncFingerprint,
	    isIntelligentProject,
	    projectId,
	    store.activeAmountSourceId,
	    store.amountSources,
	    store.blueprint,
	    store.isLoading,
	    store.isSaving,
	    store.projectState,
	    syncStatus,
	    syncToIct,
	  ]);

  const openIct = () => {
    if (!projectId || !activeProject || !workspaceId) return;
    openIctFromIntelligentCompute({
      type: "intelligent_compute",
      workspaceId,
      projectId,
      projectName: activeProject.name,
      amountSourceId: store.activeAmountSourceId,
    });
  };

  return (
    <div className="flex h-full min-h-0 flex-1 flex-col overflow-hidden bg-background text-foreground">
      <header className="shrink-0 bg-card shadow-sm">
        <div className="flex min-h-14 items-center gap-3 px-5 py-2">
          <Button variant="ghost" size="sm" onClick={() => navigateTo("project_board")}>← 返回项目看板</Button>
          <div className="min-w-0">
            <h1 className="truncate text-page-title font-bold">智算测算</h1>
            <p className="truncate text-caption text-secondary-foreground">维护智算金额来源并显式同步到 ICT 测算</p>
          </div>
          <div className="relative ml-auto flex shrink-0 items-center gap-2">
	            <Button
	              variant="outline"
	              disabled={!projectId || syncStatus === "syncing"}
	              onClick={() => void syncToIct()}
            >
              <AppIcon name={syncStatus === "syncing" ? "loading" : "reverse"} className={syncStatus === "syncing" ? "animate-spin" : ""} />
              {syncStatus === "syncing" ? "同步中..." : "同步到 ICT"}
            </Button>
            <Button
              variant="outline"
              disabled={!projectId || store.isSaving}
              onClick={() => setShowSaveAsModal(true)}
            >
              <AppIcon name="copy" />另存为来源
            </Button>
            <Button onClick={() => void store.save(projectId)}>
              <AppIcon name="save" />保存金额来源
            </Button>
            <Button
              variant="outline"
              size="icon"
              aria-label="更多操作"
              aria-expanded={moreMenuOpen}
              onClick={() => setMoreMenuOpen(value => !value)}
            >
              <span className="text-lg leading-none">•••</span>
            </Button>
            {moreMenuOpen && (
              <div className="absolute right-0 top-11 z-40 w-44 rounded-lg bg-card p-1.5 shadow-lg ring-1 ring-border/70">
                <button
                  type="button"
                  className="flex w-full items-center gap-2 rounded-md px-3 py-2 text-left text-sm hover:bg-muted"
                  onClick={() => {
                    store.resetToH200();
                    setMoreMenuOpen(false);
                  }}
                >
                  <AppIcon name="reverse" />恢复 H200 默认值
                </button>
	                <button
	                  type="button"
	                  className="flex w-full items-center gap-2 rounded-md px-3 py-2 text-left text-sm hover:bg-muted"
	                  onClick={() => {
	                    setShowAmountSourceManager(true);
	                    setMoreMenuOpen(false);
	                  }}
	                >
	                  <AppIcon name="presets" />管理金额来源
	                </button>
                <button
                  type="button"
                  className="flex w-full items-center gap-2 rounded-md px-3 py-2 text-left text-sm hover:bg-muted"
                  onClick={() => {
                    void exportCurrentAmountSource();
                    setMoreMenuOpen(false);
                  }}
                >
                  <AppIcon name="exportReport" />导出当前来源
                </button>
                <button
                  type="button"
                  className="flex w-full items-center gap-2 rounded-md px-3 py-2 text-left text-sm hover:bg-muted"
                  onClick={() => {
                    void selectImportAmountSourcePackage();
                    setMoreMenuOpen(false);
                  }}
                >
                  <AppIcon name="importExcel" />导入来源
                </button>
                <button
                  type="button"
                  disabled={!canDeleteIntelligentAmountSource(store.amountSources, store.activeAmountSourceId)}
                  className="flex w-full items-center gap-2 rounded-md px-3 py-2 text-left text-sm hover:bg-muted disabled:opacity-45"
                  onClick={() => {
                    if (
                      store.activeAmountSourceId
                      && canDeleteIntelligentAmountSource(store.amountSources, store.activeAmountSourceId)
                      && window.confirm("确认删除当前智算金额来源吗？")
                    ) {
                      void (async () => {
                        if (await store.deleteAmountSource(store.activeAmountSourceId!)) clearSyncPreview();
                      })();
                    }
                    setMoreMenuOpen(false);
                  }}
                >
                  <AppIcon name="delete" />删除当前来源
                </button>
                <button
                  type="button"
                  className="flex w-full items-center gap-2 rounded-md px-3 py-2 text-left text-sm hover:bg-muted"
                  onClick={() => {
                    openIct();
                    setMoreMenuOpen(false);
                  }}
                >
                  <AppIcon name="calculator" />打开 ICT 测算
                </button>
                <button
                  type="button"
                  className="flex w-full items-center gap-2 rounded-md px-3 py-2 text-left text-sm hover:bg-muted"
	                  onClick={() => {
	                    setShowOutput(true);
                    setMoreMenuOpen(false);
                  }}
                >
                  <AppIcon name="document" />查看输出包
                </button>
                <button
                  type="button"
                  className="flex w-full items-center gap-2 rounded-md px-3 py-2 text-left text-sm hover:bg-muted disabled:cursor-not-allowed disabled:opacity-45"
                  disabled={!ictExportPreview}
                  onClick={() => {
                    setShowSyncDetails(true);
                    setMoreMenuOpen(false);
                  }}
                >
                  <AppIcon name="tableProperties" />查看同步明细
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="flex flex-wrap items-center gap-2 px-5 pb-3">
          <HeaderStatusChip label="当前项目" value={activeProject?.name || "未选择智算项目"} />
          {projectCycleParameter && (
            <label className="flex h-9 items-center gap-2 rounded-lg bg-muted/60 px-3 text-caption font-semibold text-secondary-foreground">
              项目周期
              <input
                type="number"
                min={1}
                max={10}
                value={getAiComputeProjectCycleYears(store.blueprint.parameters)}
                onChange={event => store.updateParameter(projectCycleParameter.id, { value: Number(event.target.value) })}
                className="w-12 bg-transparent text-right font-bold text-foreground outline-none"
              />
              年
            </label>
          )}
          {discountRateParameter && (
            <label className="flex h-9 items-center gap-2 rounded-lg bg-muted/60 px-3 text-caption font-semibold text-secondary-foreground">
              折现率
              <input
                type="number"
                min={0}
                max={100}
                step={0.1}
                value={getAiComputeDiscountRatePercent(store.blueprint.parameters)}
                onChange={event => store.updateParameter(discountRateParameter.id, { value: Number(event.target.value) })}
                className="w-16 bg-transparent text-right font-bold text-foreground outline-none"
              />
              %
            </label>
          )}
	          <label className="flex h-9 min-w-64 items-center gap-2 rounded-lg bg-muted/60 px-3 text-caption font-semibold text-secondary-foreground">
	            金额来源
            <select
              className="min-w-0 flex-1 bg-transparent text-sm font-bold text-foreground outline-none"
              value={store.activeAmountSourceId || ""}
              onChange={event => {
                const nextSourceId = event.target.value;
                void (async () => {
                  if (await store.setActiveAmountSource(nextSourceId)) clearSyncPreview();
                })();
              }}
            >
              {store.amountSources.map(source => (
                <option key={source.id} value={source.id}>
                  {source.name}
                </option>
	              ))}
	            </select>
	          </label>
	          <Button
	            variant="outline"
	            size="sm"
	            onClick={() => setShowAmountSourceManager(true)}
	          >
	            <AppIcon name="edit" />管理来源
	          </Button>
          <HeaderStatusChip
            label="ICT 状态"
            value={`${activeProject?.benefit_status === "normal" ? "已有正式测算" : "待配置"} · ${getSyncStatusLabel(syncStatus)}`}
            tone={syncStatus === "error" || syncStatus === "conflict" ? "warning" : syncStatus === "synced" ? "success" : undefined}
          />
          <HeaderStatusChip
            label="同步来源"
            value={store.amountSources.find(source => source.id === store.activeAmountSourceId)?.name || "--"}
          />
          <HeaderStatusChip
            label="保存状态"
            value={getSavedStatusLabel(store.isDirty, store.lastSavedAt)}
            tone={store.isDirty ? "warning" : store.lastSavedAt ? "success" : undefined}
          />
        </div>
      </header>

      <main className="flex min-h-0 flex-1 flex-col gap-3 overflow-y-auto p-4 xl:overflow-hidden">
        {projectId && activeProject && !isIntelligentProject && (
          <div className="shrink-0 rounded-lg bg-warning-soft px-3 py-2 text-caption text-warning-foreground">
            当前项目类型为 ICT，不能加载智算金额来源。请返回项目看板选择智算项目。
          </div>
        )}
        {(store.error || ictExportError) && (
          <div className="shrink-0 rounded-lg bg-destructive-soft px-3 py-2 text-caption text-destructive">
            {store.error || ictExportError}
          </div>
        )}
        <div className="grid min-h-0 flex-1 gap-3 xl:grid-cols-[minmax(0,1fr)_292px]">
          <section className="flex min-h-[560px] min-w-0 flex-col overflow-hidden rounded-xl bg-card shadow-sm xl:min-h-0">
            <QuoteWorkspaceTabs
              activeTab={activeTab}
              parameterCount={store.blueprint.parameters.length}
              revenueCount={store.blueprint.revenueItems.length}
              costCount={store.blueprint.costItems.length}
              sensitivityCount={store.blueprint.parameters.filter(parameter => parameter.sensitivityEnabled !== false).length}
              onChange={setActiveTab}
            />
            <div className="relative min-h-0 flex-1">
              <div
                id="ai-compute-tab-panel-parameters"
                role="tabpanel"
                className={`h-full overflow-y-auto p-4 ${activeTab === "parameters" ? "block" : "hidden"}`}
              >
                <ParameterPanel />
              </div>
              <div
                id="ai-compute-tab-panel-revenue"
                role="tabpanel"
                className={`h-full overflow-y-auto p-4 ${activeTab === "revenue" ? "block" : "hidden"}`}
              >
                <LineItemPanel side="revenue" title="收入计算项" syncStatus={syncStatus} />
              </div>
              <div
                id="ai-compute-tab-panel-cost"
                role="tabpanel"
                className={`h-full overflow-y-auto p-4 ${activeTab === "cost" ? "block" : "hidden"}`}
              >
                <LineItemPanel side="cost" title="成本计算项" syncStatus={syncStatus} />
              </div>
              <div
                id="ai-compute-tab-panel-sensitivity"
                role="tabpanel"
                className={`h-full overflow-y-auto p-4 ${activeTab === "sensitivity" ? "block" : "hidden"}`}
              >
                <SensitivityPanel
                  parameterId={sensitivityParameterId}
                  min={sensitivityMin}
                  max={sensitivityMax}
                  step={sensitivityStep}
                  rows={ictSensitivityRows}
                  onParameterChange={setSensitivityParameterId}
                  onMinChange={setSensitivityMin}
                  onMaxChange={setSensitivityMax}
                  onStepChange={setSensitivityStep}
                />
              </div>
            </div>
          </section>
          <BenefitConclusionSidebar
            summary={summary}
            ictResult={ictResult}
            syncStatus={syncStatus}
            majorRevenue={majorRevenue}
            majorCost={majorCost}
            outputCount={outputPackage.length}
            syncDetailsAvailable={Boolean(ictExportPreview)}
            onShowOutput={() => setShowOutput(true)}
            onShowSyncDetails={() => setShowSyncDetails(true)}
          />
        </div>
      </main>

      {showSaveAsModal && (
        <SaveAsAmountSourceModal
          currentName={store.blueprint.name}
          busy={store.isSaving}
          onClose={() => setShowSaveAsModal(false)}
          onSave={saveCurrentAsAmountSource}
        />
      )}
      {pendingImportPackage && store.projectState && (
        <ImportAmountSourceModal
          pkg={pendingImportPackage}
          currentProjectYears={store.projectState.projectYears}
          currentDiscountRate={store.projectState.discountRate}
          busy={store.isSaving}
          onClose={() => setPendingImportPackage(null)}
          onImport={confirmImportAmountSourcePackage}
        />
      )}
      {showOutput && (
        <OutputPackageModal
          onClose={() => setShowOutput(false)}
	          onRequestIctOutput={() => {
	            setShowOutput(false);
	            void syncToIct();
	          }}
	        />
	      )}
	      {showAmountSourceManager && (
	        <AmountSourceManagerModal
	          sources={store.amountSources}
	          activeSourceId={store.activeAmountSourceId}
	          onClose={() => setShowAmountSourceManager(false)}
	          onRename={renameAmountSource}
            onSelect={async sourceId => {
              const ok = await useAiComputeQuoteStore.getState().setActiveAmountSource(sourceId);
              if (ok) clearSyncPreview();
              return ok;
            }}
	          onCreate={async request => {
	            const ok = await useAiComputeQuoteStore.getState().createAmountSource(request);
              if (ok) clearSyncPreview();
              return ok;
	          }}
            onDelete={async sourceId => {
              const ok = await useAiComputeQuoteStore.getState().deleteAmountSource(sourceId);
              if (ok) clearSyncPreview();
              return ok;
            }}
	        />
	      )}
	      {showSyncDetails && ictExportPreview && (
        <IctExportConfirmModal
          preview={ictExportPreview}
          error={ictExportError}
          syncing={syncStatus === "syncing"}
          requiresReload={syncStatus === "conflict"}
	          onConfirm={() => void (
	            syncStatus === "conflict"
	              ? reloadConflictAndOverwrite()
	              : syncToIct({ confirmed: true })
	          )}
          onCancel={() => {
            setShowSyncDetails(false);
            setIctExportError(null);
          }}
        />
      )}
    </div>
  );
	}

function SaveAsAmountSourceModal({
  currentName,
  busy,
  onClose,
  onSave,
}: {
  currentName: string;
  busy: boolean;
  onClose: () => void;
  onSave: (name: string) => Promise<boolean>;
}) {
  const [name, setName] = useState(`${currentName || "当前金额来源"} 副本`);
  const [error, setError] = useState<string | null>(null);

  const handleSave = async () => {
    const nextName = name.trim();
    if (!nextName) {
      setError("请填写金额来源名称。");
      return;
    }
    setError(null);
    const ok = await onSave(nextName);
    if (!ok) {
      setError("当前测算另存失败，请检查项目状态后重试。");
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-foreground/20 p-4 backdrop-blur-sm">
      <div className="w-full max-w-md rounded-xl bg-card p-5 shadow-xl ring-1 ring-border/60">
        <div className="mb-4 flex items-center justify-between gap-3">
          <div>
            <h2 className="text-section-title font-bold">当前测算另存为来源</h2>
            <p className="text-caption text-secondary-foreground">复制当前画布，不覆盖原金额来源。</p>
          </div>
          <Button variant="ghost" size="sm" disabled={busy} onClick={onClose}>关闭</Button>
        </div>
        <div className="rounded-lg bg-muted/45 p-3.5">
          <div className="mb-2 text-caption font-extrabold text-foreground">新来源名称</div>
          <Input
            value={name}
            disabled={busy}
            onChange={event => setName(event.target.value)}
            className="bg-card"
          />
        </div>
        {error && (
          <div className="mt-3 rounded-lg bg-destructive-soft px-3 py-2 text-caption font-semibold text-destructive">
            {error}
          </div>
        )}
        <div className="mt-4 flex justify-end gap-2">
          <Button variant="outline" disabled={busy} onClick={onClose}>取消</Button>
          <Button disabled={busy || !name.trim()} onClick={() => void handleSave()}>
            <AppIcon name={busy ? "loading" : "copy"} className={busy ? "animate-spin" : ""} />
            {busy ? "另存中..." : "另存为来源"}
          </Button>
        </div>
      </div>
    </div>
  );
}

function ImportAmountSourceModal({
  pkg,
  currentProjectYears,
  currentDiscountRate,
  busy,
  onClose,
  onImport,
}: {
  pkg: AiComputeAmountSourcePackage;
  currentProjectYears: number;
  currentDiscountRate: number;
  busy: boolean;
  onClose: () => void;
  onImport: (pkg: AiComputeAmountSourcePackage, options: { name: string; applyProjectSettings: boolean }) => Promise<boolean>;
}) {
  const [name, setName] = useState(getDefaultImportedAmountSourceName(pkg));
  const [applyProjectSettings, setApplyProjectSettings] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const handleImport = async () => {
    const nextName = name.trim();
    if (!nextName) {
      setError("请填写导入后的金额来源名称。");
      return;
    }
    setError(null);
    const ok = await onImport(pkg, { name: nextName, applyProjectSettings });
    if (!ok) {
      setError("金额来源导入失败，请检查当前项目状态后重试。");
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-foreground/20 p-4 backdrop-blur-sm">
      <div className="w-full max-w-xl rounded-xl bg-card p-5 shadow-xl ring-1 ring-border/60">
        <div className="mb-4 flex items-center justify-between gap-3">
          <div>
            <h2 className="text-section-title font-bold">导入金额来源</h2>
            <p className="text-caption text-secondary-foreground">导入后生成一个新来源，并设为当前同步来源。</p>
          </div>
          <Button variant="ghost" size="sm" disabled={busy} onClick={onClose}>关闭</Button>
        </div>

        <div className="grid gap-3 sm:grid-cols-2">
          <div className="rounded-lg bg-muted/45 p-3.5">
            <div className="text-caption font-extrabold text-foreground">文件来源</div>
            <div className="mt-2 text-sm font-bold text-foreground">{pkg.source.name}</div>
            <div className="mt-1 text-caption text-secondary-foreground">
              导出时间：{new Date(pkg.exportedAt).toLocaleString("zh-CN")}
            </div>
          </div>
          <div className="rounded-lg bg-muted/45 p-3.5">
            <div className="text-caption font-extrabold text-foreground">项目参数</div>
            <div className="mt-2 space-y-1 text-caption text-secondary-foreground">
              <div>当前项目：{currentProjectYears} 年 / {formatNumber(currentDiscountRate * 100)}%</div>
              <div>导入文件：{pkg.projectSettings.projectYears} 年 / {formatNumber(pkg.projectSettings.discountRate * 100)}%</div>
            </div>
          </div>
        </div>

        <div className="mt-3 rounded-lg bg-muted/45 p-3.5">
          <div className="mb-2 text-caption font-extrabold text-foreground">导入后名称</div>
          <Input
            value={name}
            disabled={busy}
            onChange={event => setName(event.target.value)}
            className="bg-card"
          />
          <label className="mt-3 flex items-center gap-2 text-caption font-semibold text-secondary-foreground">
            <input
              type="checkbox"
              checked={applyProjectSettings}
              disabled={busy}
              onChange={event => setApplyProjectSettings(event.target.checked)}
            />
            使用导入文件的项目周期和折现率
          </label>
        </div>

        {error && (
          <div className="mt-3 rounded-lg bg-destructive-soft px-3 py-2 text-caption font-semibold text-destructive">
            {error}
          </div>
        )}
        <div className="mt-4 flex justify-end gap-2">
          <Button variant="outline" disabled={busy} onClick={onClose}>取消</Button>
          <Button disabled={busy || !name.trim()} onClick={() => void handleImport()}>
            <AppIcon name={busy ? "loading" : "importExcel"} className={busy ? "animate-spin" : ""} />
            {busy ? "导入中..." : "导入为新来源"}
          </Button>
        </div>
      </div>
    </div>
  );
}

function AmountSourceManagerModal({
  sources,
  activeSourceId,
  onClose,
  onRename,
  onSelect,
  onCreate,
  onDelete,
}: {
  sources: IntelligentAmountSource[];
  activeSourceId: string | null;
  onClose: () => void;
  onRename: (name: string) => Promise<boolean>;
  onSelect: (sourceId: string) => Promise<boolean>;
  onCreate: (request: CreateAmountSourceRequest) => Promise<boolean>;
  onDelete: (sourceId: string) => Promise<boolean>;
}) {
  const activeSource = sources.find(source => source.id === activeSourceId) || null;
  const [renameName, setRenameName] = useState(activeSource?.name || "");
  const [newName, setNewName] = useState("");
  const [baseValue, setBaseValue] = useState<CreateAmountSourceBaseMode | `source:${string}`>(
    getDefaultCreateAmountSourceBaseMode(),
  );
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    setRenameName(activeSource?.name || "");
    setBaseValue(getDefaultCreateAmountSourceBaseMode());
  }, [activeSource?.id, activeSource?.name]);

  const resolveBase = (): Pick<CreateAmountSourceRequest, "baseMode" | "baseSourceId"> => {
    if (baseValue.startsWith("source:")) {
      return { baseMode: "source", baseSourceId: baseValue.slice("source:".length) };
    }
    return { baseMode: baseValue as CreateAmountSourceBaseMode, baseSourceId: null };
  };

  const handleRename = async () => {
    const name = renameName.trim();
    if (!activeSource || !name) {
      setError("请填写当前金额来源名称。");
      return;
    }
    setBusy(true);
    setError(null);
    const ok = await onRename(name);
    setBusy(false);
    if (!ok) {
      setError("金额来源重命名失败，请稍后重试。");
    }
  };

  const handleCreate = async () => {
    const name = newName.trim();
    if (!name) {
      setError("请填写新金额来源名称。");
      return;
    }
    const base = resolveBase();
    if (base.baseMode === "source" && !base.baseSourceId) {
      setError("请选择一个已有金额来源作为基底。");
      return;
    }
    setBusy(true);
    setError(null);
    const ok = await onCreate({
      name,
      ...base,
    });
    setBusy(false);
    if (ok) {
      onClose();
      return;
    }
    setError("金额来源创建失败，请检查当前来源是否已保存。");
  };

  const handleSelect = async (sourceId: string) => {
    if (sourceId === activeSourceId) return;
    setBusy(true);
    setError(null);
    const ok = await onSelect(sourceId);
    setBusy(false);
    if (!ok) {
      setError("切换金额来源失败，请先检查当前来源是否已保存。");
    }
  };

  const handleDelete = async (source: IntelligentAmountSource) => {
    if (!canDeleteIntelligentAmountSource(sources, source.id)) {
      setError(isH200BaselineAmountSource(source, sources)
        ? "H200 标准基准来源不能删除，可复制后基于副本报价。"
        : "当前金额来源不能删除。");
      return;
    }
    if (!window.confirm(`确认删除金额来源“${source.name}”吗？`)) return;
    setBusy(true);
    setError(null);
    const ok = await onDelete(source.id);
    setBusy(false);
    if (!ok) {
      setError("金额来源删除失败，请稍后重试。");
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-foreground/20 p-4 backdrop-blur-sm">
      <div className="w-full max-w-2xl rounded-xl bg-card p-5 shadow-xl ring-1 ring-border/60">
        <div className="mb-4 flex items-center justify-between gap-3">
          <div>
            <h2 className="text-section-title font-bold">金额来源管理</h2>
            <p className="text-caption text-secondary-foreground">当前项目共 {sources.length} 个来源</p>
          </div>
          <Button variant="ghost" size="sm" onClick={onClose}>关闭</Button>
        </div>

        {error && (
          <div className="mb-3 rounded-lg bg-destructive-soft px-3 py-2 text-caption font-semibold text-destructive">
            {error}
          </div>
        )}

        <section className="rounded-lg bg-muted/45 p-3.5">
          <div className="mb-2 flex items-center justify-between gap-3">
            <div className="text-caption font-extrabold text-foreground">来源列表</div>
            <div className="text-caption font-semibold text-secondary-foreground">
              当前同步来源：{activeSource?.name || "--"}
            </div>
          </div>
          <div className="space-y-2">
            {sources.map(source => {
              const active = source.id === activeSourceId;
              const protectedBaseline = isH200BaselineAmountSource(source, sources);
              const canDelete = canDeleteIntelligentAmountSource(sources, source.id);
              return (
                <div key={source.id} className="flex flex-wrap items-center gap-2 rounded-lg bg-card px-3 py-2">
                  <div className="min-w-0 flex-1">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="truncate text-sm font-bold text-foreground">{source.name}</span>
                      {active && (
                        <span className="rounded-full bg-primary-soft px-2 py-0.5 text-[11px] font-bold text-primary">
                          当前来源
                        </span>
                      )}
                      {protectedBaseline && (
                        <span className="rounded-full bg-secondary px-2 py-0.5 text-[11px] font-bold text-secondary-foreground">
                          H200 基准
                        </span>
                      )}
                    </div>
                    <div className="mt-0.5 text-caption text-secondary-foreground">
                      {protectedBaseline ? "系统保护的标准基准，可复制但不可删除" : "普通报价来源，可设为当前或删除"}
                    </div>
                  </div>
                  <Button
                    variant={active ? "secondary" : "outline"}
                    size="sm"
                    disabled={busy || active}
                    onClick={() => void handleSelect(source.id)}
                  >
                    <AppIcon name={active ? "check" : "reverse"} />{active ? "已选中" : "设为当前"}
                  </Button>
                  <Button
                    variant="ghost"
                    size="sm"
                    disabled={busy || !canDelete}
                    title={protectedBaseline ? "H200 标准基准来源不可删除" : sources.length <= 1 ? "至少保留一个金额来源" : "删除来源"}
                    onClick={() => void handleDelete(source)}
                  >
                    <AppIcon name="delete" />删除
                  </Button>
                </div>
              );
            })}
          </div>
        </section>

        <section className="mt-3 rounded-lg bg-muted/45 p-3.5">
          <div className="mb-2 text-caption font-extrabold text-foreground">当前来源名称</div>
          <div className="flex gap-2">
            <Input
              value={renameName}
              disabled={!activeSource || busy}
              onChange={event => setRenameName(event.target.value)}
              className="bg-card"
            />
            <Button disabled={!activeSource || busy || !renameName.trim()} onClick={handleRename}>
              <AppIcon name="save" />保存名称
            </Button>
          </div>
        </section>

        <section className="mt-3 rounded-lg bg-muted/45 p-3.5">
          <div className="mb-2 text-caption font-extrabold text-foreground">新建金额来源</div>
          <div className="grid gap-2 sm:grid-cols-[minmax(0,1fr)_180px]">
            <Input
              value={newName}
              disabled={busy}
              placeholder="金额来源名称"
              onChange={event => setNewName(event.target.value)}
              className="bg-card"
            />
            <select
              value={baseValue}
              disabled={busy}
              onChange={event => setBaseValue(event.target.value as typeof baseValue)}
              className="h-10 rounded-md bg-card px-3 text-sm font-semibold outline-none ring-1 ring-input focus:ring-ring"
            >
              <option value="h200">H200 标准</option>
              <option value="blank">空白来源</option>
              {activeSource && <option value="current">当前来源</option>}
              {sources.map(source => (
                <option key={source.id} value={`source:${source.id}`}>
                  {source.name}
                </option>
              ))}
            </select>
          </div>
          <div className="mt-3 flex flex-wrap items-center justify-between gap-3">
            <div className="text-caption font-semibold text-secondary-foreground">
              新建后将设为当前唯一 ICT 同步来源
            </div>
            <Button disabled={busy || !newName.trim()} onClick={handleCreate}>
              <AppIcon name="copy" />新建来源
            </Button>
          </div>
        </section>
      </div>
    </div>
  );
}

function QuoteWorkspaceTabs({
  activeTab,
  parameterCount,
  revenueCount,
  costCount,
  sensitivityCount,
  onChange,
}: {
  activeTab: QuoteWorkspaceTab;
  parameterCount: number;
  revenueCount: number;
  costCount: number;
  sensitivityCount: number;
  onChange: (tab: QuoteWorkspaceTab) => void;
}) {
  const tabs: Array<{ id: QuoteWorkspaceTab; label: string; count: number; marker: string }> = [
    { id: "parameters", label: "参数区", count: parameterCount, marker: "ƒ" },
    { id: "revenue", label: "收入计算项", count: revenueCount, marker: "↗" },
    { id: "cost", label: "成本计算项", count: costCount, marker: "↘" },
    { id: "sensitivity", label: "敏感性分析", count: sensitivityCount, marker: "≋" },
  ];

  return (
    <div
      className="flex shrink-0 items-center gap-1 overflow-x-auto bg-muted/25 px-4 pt-2"
      style={{ borderBottom: "0.5px solid hsl(var(--border))" }}
      role="tablist"
      aria-label="智算金额来源编辑区"
    >
      {tabs.map(tab => {
        const active = activeTab === tab.id;
        return (
          <button
            key={tab.id}
            type="button"
            role="tab"
            aria-selected={active}
            aria-controls={`ai-compute-tab-panel-${tab.id}`}
            onClick={() => onChange(tab.id)}
            className={`mb-2 flex shrink-0 items-center gap-2 rounded-lg px-4 py-2.5 text-sm transition-colors ${
              active
                ? "bg-primary-soft/80 font-extrabold text-primary"
                : "font-semibold text-secondary-foreground hover:bg-card/60 hover:text-foreground"
            }`}
          >
            <span aria-hidden="true" className="text-base">{tab.marker}</span>
            {tab.label}
            <span className={`numeric-value rounded-full px-2 py-0.5 text-[11px] ${
              active ? "bg-primary-soft text-primary" : "bg-secondary text-secondary-foreground"
            }`}>
              {tab.count}
            </span>
          </button>
        );
      })}
    </div>
  );
}

function HeaderStatusChip({
  label,
  value,
  tone,
}: {
  label: string;
  value: string;
  tone?: "success" | "warning";
}) {
  return (
    <div className="flex h-9 min-w-0 items-center gap-2 rounded-lg bg-muted/60 px-3">
      <span className="shrink-0 text-[11px] font-semibold text-secondary-foreground">{label}</span>
      <span className={`truncate text-xs font-bold ${
        tone === "success"
          ? "text-success-foreground"
          : tone === "warning"
            ? "text-warning-foreground"
            : "text-foreground"
      }`}>
        {value}
      </span>
    </div>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return <label className="grid gap-1 text-caption font-semibold text-secondary-foreground">{label}{children}</label>;
}

function ParameterPanel() {
  const store = useAiComputeQuoteStore();
  const [filterMode, setFilterMode] = useState<ParameterFilterMode>("key");
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedParameterId, setSelectedParameterId] = useState("gpu-service-price");
  const [createMenuOpen, setCreateMenuOpen] = useState(false);
  const [parameterDialogGroupId, setParameterDialogGroupId] = useState<string | null>(null);
  const [categoryDialog, setCategoryDialog] = useState<{
    mode: "create" | "edit";
    group?: AiComputeQuoteParameterGroup;
  } | null>(null);
  const usageByParameter = useMemo(() => {
    const usage = new Map<string, { revenue: string[]; cost: string[] }>();
    const register = (item: AiComputeQuoteLineItem) => {
      getFormulaParameterReferences(item.formula).forEach(parameterId => {
        const current = usage.get(parameterId) || { revenue: [], cost: [] };
        current[item.side].push(item.name);
        usage.set(parameterId, current);
      });
    };
    store.blueprint.revenueItems.forEach(register);
    store.blueprint.costItems.forEach(register);
    return usage;
  }, [store.blueprint.costItems, store.blueprint.revenueItems]);
  const matchesSearch = (parameter: AiComputeQuoteParameter) => {
    const query = searchTerm.trim().toLocaleLowerCase();
    if (!query) return true;
    return [parameter.name, parameter.key, parameter.unit || ""]
      .some(value => value.toLocaleLowerCase().includes(query));
  };
  const matchesMode = (parameter: AiComputeQuoteParameter) => {
    if (filterMode === "key") return parameter.isKey === true;
    if (filterMode === "modified") return isParameterModified(parameter);
    if (filterMode === "sensitive") return parameter.sensitivityEnabled !== false;
    return true;
  };
  const visibleParameters = store.blueprint.parameters.filter(parameter =>
    matchesMode(parameter) && matchesSearch(parameter)
  );
  const selectedParameter = store.blueprint.parameters.find(parameter => parameter.id === selectedParameterId)
    || visibleParameters[0]
    || store.blueprint.parameters[0];
  const canReorder = filterMode === "all" && searchTerm.trim() === "";
  const normalAllView = filterMode === "all" && searchTerm.trim() === "";
  const visibleGroups = store.blueprint.parameterGroups.flatMap((group, groupIndex) => {
    const allGroupParameters = store.blueprint.parameters.filter(parameter => parameter.groupId === group.id);
    const groupParameters = visibleParameters.filter(parameter => parameter.groupId === group.id);
    const emptyUnclassified = group.id === PARAMETER_GROUP_IDS.unclassified && allGroupParameters.length === 0;
    if (emptyUnclassified || (!normalAllView && groupParameters.length === 0)) return [];
    return [{
      group,
      groupIndex,
      parameters: groupParameters,
      totalParameterCount: allGroupParameters.length,
    }];
  });
  const projectCycleParameter = store.blueprint.parameters.find(isAiComputeProjectCycleParameter);
  const projectCycleYears = getAiComputeProjectCycleYears(store.blueprint.parameters);
  const discountRateParameter = store.blueprint.parameters.find(isAiComputeDiscountRateParameter);
  const discountRatePercent = getAiComputeDiscountRatePercent(store.blueprint.parameters);

  return (
    <section className="mx-auto w-full max-w-[1180px]">
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 className="text-section-title">参数区</h2>
          <p className="text-caption text-secondary-foreground">
            按业务类别组织参数；切换“全部参数”后可拖拽调整类别和参数顺序。
          </p>
        </div>
        <div className="relative">
          <Button size="sm" onClick={() => setCreateMenuOpen(value => !value)}>
            新增<AppIcon name={createMenuOpen ? "chevronUp" : "chevronDown"} size={14} />
          </Button>
          {createMenuOpen && (
            <div className="absolute right-0 top-10 z-40 w-36 rounded-lg bg-card p-1.5 shadow-lg ring-1 ring-border">
              <button
                type="button"
                className="w-full rounded-md px-3 py-2 text-left text-xs font-semibold hover:bg-muted"
                onClick={() => {
                  setParameterDialogGroupId(store.blueprint.parameterGroups[0]?.id || null);
                  setCreateMenuOpen(false);
                }}
              >
                新增参数
              </button>
              <button
                type="button"
                className="w-full rounded-md px-3 py-2 text-left text-xs font-semibold hover:bg-muted"
                onClick={() => {
                  setCategoryDialog({ mode: "create" });
                  setCreateMenuOpen(false);
                }}
              >
                新增类别
              </button>
            </div>
          )}
        </div>
      </div>

      <div className="mb-4 flex flex-col gap-3 rounded-xl bg-primary-soft/55 p-4 shadow-sm lg:flex-row lg:items-center">
        <div className="flex min-w-0 flex-1 items-center gap-3">
          <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-lg bg-card text-primary shadow-sm">
            <AppIcon name="cashflow" size={20} />
          </div>
          <div className="min-w-0">
            <div className="text-sm font-extrabold text-primary">ICT 核心参数</div>
            <div className="text-caption text-secondary-foreground">
              项目周期和折现率由智算统一控制并覆盖 ICT；年度计划仍按实际收付款安排维护。
            </div>
          </div>
        </div>
        <div className="flex flex-wrap gap-2">
          <label className="flex shrink-0 items-center gap-2 rounded-lg bg-card px-3 py-2 shadow-sm">
            <span className="text-xs font-bold text-secondary-foreground">周期年限</span>
            <Input
              aria-label="智算项目周期"
              className="numeric-value h-8 w-20 text-center font-extrabold"
              type="number"
              min={1}
              max={10}
              step={1}
              value={projectCycleYears}
              disabled={!projectCycleParameter}
              onChange={event => {
                if (projectCycleParameter) {
                  store.updateParameter(projectCycleParameter.id, { value: toNumber(event.target.value) });
                }
              }}
            />
            <span className="text-xs font-bold text-secondary-foreground">年</span>
          </label>
          <label className="flex shrink-0 items-center gap-2 rounded-lg bg-card px-3 py-2 shadow-sm">
            <span className="text-xs font-bold text-secondary-foreground">项目折现率</span>
            <Input
              aria-label="智算项目折现率"
              className="numeric-value h-8 w-24 text-center font-extrabold"
              type="number"
              min={0}
              max={100}
              step={0.1}
              value={discountRatePercent}
              disabled={!discountRateParameter}
              onChange={event => {
                if (discountRateParameter) {
                  store.updateParameter(discountRateParameter.id, { value: toNumber(event.target.value) });
                }
              }}
            />
            <span className="text-xs font-bold text-secondary-foreground">%</span>
          </label>
        </div>
      </div>

      <div className="mb-4 flex flex-wrap items-center gap-2">
        {([
          ["key", "仅看关键参数"],
          ["modified", "仅看已修改"],
          ["sensitive", "仅看敏感性参数"],
          ["all", "全部参数"],
        ] as Array<[ParameterFilterMode, string]>).map(([mode, label]) => (
          <button
            key={mode}
            type="button"
            className={`h-9 rounded-lg px-3 text-xs font-bold transition-colors ${
              filterMode === mode
                ? "bg-primary-soft text-primary ring-1 ring-primary/20"
                : "bg-muted/65 text-secondary-foreground hover:bg-muted hover:text-foreground"
            }`}
            onClick={() => setFilterMode(mode)}
          >
            {label}
          </button>
        ))}
        <label className="relative ml-auto min-w-56 flex-1 sm:max-w-72">
          <AppIcon name="search" className="pointer-events-none absolute left-3 top-2.5 text-secondary-foreground" size={16} />
          <Input
            className="h-9 pl-9 text-sm"
            value={searchTerm}
            placeholder="搜索中文名、字段 key 或单位"
            onChange={event => setSearchTerm(event.target.value)}
          />
        </label>
      </div>

      {!canReorder && (
        <div className="mb-3 rounded-lg bg-muted/45 px-3 py-2 text-[11px] text-secondary-foreground">
          当前为筛选视图。参数仍可通过设置菜单调整类别；拖拽排序请切换到“全部参数”并清空搜索。
        </div>
      )}

      <div className="space-y-3">
        {visibleGroups.map(({ group, groupIndex, parameters, totalParameterCount }) => (
          <ParameterGroupSection
            key={group.id}
            group={group}
            groupIndex={groupIndex}
            groupCount={store.blueprint.parameterGroups.length}
            parameters={parameters}
            totalParameterCount={totalParameterCount}
            allGroups={store.blueprint.parameterGroups}
            usageByParameter={usageByParameter}
            selectedParameterId={selectedParameter?.id || null}
            canReorder={canReorder}
            forceOpen={Boolean(searchTerm.trim()) || filterMode !== "key"}
            onSelect={setSelectedParameterId}
            onAddParameter={groupId => setParameterDialogGroupId(groupId)}
            onEditCategory={groupValue => setCategoryDialog({ mode: "edit", group: groupValue })}
          />
        ))}
      </div>
      {visibleGroups.length === 0 && (
        <div className="rounded-lg bg-muted/55 px-4 py-10 text-center text-sm text-secondary-foreground">
          没有符合当前筛选条件的参数。
        </div>
      )}

      {parameterDialogGroupId && (
        <ParameterEditorModal
          initialGroupId={parameterDialogGroupId}
          onClose={() => setParameterDialogGroupId(null)}
          onCreated={parameterId => {
            setSelectedParameterId(parameterId);
            setFilterMode("all");
            setParameterDialogGroupId(null);
          }}
        />
      )}
      {categoryDialog && (
        <ParameterGroupEditorModal
          mode={categoryDialog.mode}
          group={categoryDialog.group}
          onClose={() => setCategoryDialog(null)}
          onSaved={() => {
            setFilterMode("all");
            setCategoryDialog(null);
          }}
        />
      )}
    </section>
  );
}

function ParameterGroupSection({
  group,
  groupIndex,
  groupCount,
  parameters,
  totalParameterCount,
  allGroups,
  usageByParameter,
  selectedParameterId,
  canReorder,
  forceOpen,
  onSelect,
  onAddParameter,
  onEditCategory,
}: {
  group: AiComputeQuoteParameterGroup;
  groupIndex: number;
  groupCount: number;
  parameters: AiComputeQuoteParameter[];
  totalParameterCount: number;
  allGroups: AiComputeQuoteParameterGroup[];
  usageByParameter: Map<string, { revenue: string[]; cost: string[] }>;
  selectedParameterId: string | null;
  canReorder: boolean;
  forceOpen: boolean;
  onSelect: (parameterId: string) => void;
  onAddParameter: (groupId: string) => void;
  onEditCategory: (group: AiComputeQuoteParameterGroup) => void;
}) {
  const store = useAiComputeQuoteStore();
  const [expanded, setExpanded] = useState(
    group.id !== PARAMETER_GROUP_IDS.operations && group.id !== PARAMETER_GROUP_IDS.finance,
  );
  const open = forceOpen || expanded;
  const handleDrop = (event: React.DragEvent<HTMLElement>) => {
    if (!canReorder) return;
    event.preventDefault();
    event.stopPropagation();
    const parameterId = event.dataTransfer.getData(PARAMETER_DRAG_TYPE);
    if (parameterId) {
      store.moveParameter(parameterId, group.id);
      return;
    }
    const groupId = event.dataTransfer.getData(PARAMETER_GROUP_DRAG_TYPE);
    if (groupId && groupId !== group.id) store.reorderParameterGroup(groupId, group.id);
  };

  return (
    <section
      className={`rounded-lg bg-card shadow-sm ring-1 ring-border/70 ${
        open ? "shadow-sm" : "transition-colors hover:bg-muted/20"
      }`}
      style={open ? { boxShadow: "inset 3px 0 0 hsl(var(--primary) / 0.82), 0 1px 2px rgb(15 23 42 / 0.05)" } : undefined}
      onDragOver={event => {
        if (canReorder) event.preventDefault();
      }}
      onDrop={handleDrop}
    >
      <div className="flex min-h-12 items-center gap-2 px-3 py-2">
        <button
          type="button"
          draggable={canReorder}
          className={`flex h-8 w-6 items-center justify-center rounded text-muted-foreground ${
            canReorder ? "cursor-grab hover:bg-muted active:cursor-grabbing" : "cursor-not-allowed opacity-35"
          }`}
          aria-label={`拖拽排序类别 ${group.name}`}
          title={canReorder ? "拖拽调整类别顺序" : "在全部参数视图中可拖拽"}
          onDragStart={event => {
            event.dataTransfer.effectAllowed = "move";
            event.dataTransfer.setData(PARAMETER_GROUP_DRAG_TYPE, group.id);
          }}
        >
          ⋮⋮
        </button>
        <button
          type="button"
          className="flex min-w-0 flex-1 items-center gap-3 text-left"
          title={group.description}
          aria-expanded={open}
          onClick={() => setExpanded(value => !value)}
        >
          <AppIcon name={getParameterGroupIcon(group.id)} size={18} />
          <span className="shrink-0 text-sm font-bold">{group.name}</span>
          {!open && parameters.length > 0 && (
            <span className="flex min-w-0 flex-1 items-center gap-4 overflow-hidden">
              {parameters.slice(0, 3).map(parameter => (
                <span key={parameter.id} className="flex min-w-0 items-baseline gap-1 whitespace-nowrap text-xs">
                  <span className="truncate text-secondary-foreground">{parameter.name}</span>
                  <span className="numeric-value font-bold text-foreground">{formatParameterValue(parameter.value)}</span>
                  <span className="text-secondary-foreground">{parameter.unit}</span>
                </span>
              ))}
            </span>
          )}
          <span className="numeric-value ml-auto rounded-md bg-muted px-2 py-0.5 text-[10px] font-bold text-secondary-foreground">
            {parameters.length}
          </span>
          <AppIcon name={open ? "chevronUp" : "chevronDown"} size={16} />
        </button>
        <Button
          type="button"
          variant="ghost"
          size="sm"
          className="h-8 px-2"
          onClick={() => onAddParameter(group.id)}
        >
          + 参数
        </Button>
        <details className="relative">
          <summary
            className="flex h-8 w-8 cursor-pointer list-none items-center justify-center rounded-md text-secondary-foreground hover:bg-muted hover:text-foreground [&::-webkit-details-marker]:hidden"
            aria-label={`管理类别 ${group.name}`}
            title="类别设置"
          >
            •••
          </summary>
          <div className="absolute right-0 top-9 z-40 w-40 rounded-lg bg-card p-1.5 shadow-lg ring-1 ring-border">
            <button
              type="button"
              className="w-full rounded-md px-3 py-2 text-left text-xs font-semibold hover:bg-muted"
              onClick={() => onEditCategory(group)}
            >
              重命名类别
            </button>
            <button
              type="button"
              disabled={groupIndex === 0}
              className="w-full rounded-md px-3 py-2 text-left text-xs font-semibold hover:bg-muted disabled:opacity-40"
              onClick={() => store.moveParameterGroupByOffset(group.id, -1)}
            >
              上移类别
            </button>
            <button
              type="button"
              disabled={groupIndex === groupCount - 1}
              className="w-full rounded-md px-3 py-2 text-left text-xs font-semibold hover:bg-muted disabled:opacity-40"
              onClick={() => store.moveParameterGroupByOffset(group.id, 1)}
            >
              下移类别
            </button>
            <button
              type="button"
              disabled={group.builtin || totalParameterCount > 0}
              className="w-full rounded-md px-3 py-2 text-left text-xs font-semibold text-destructive hover:bg-destructive-soft disabled:text-muted-foreground disabled:opacity-40"
              title={group.builtin ? "内置类别不可删除" : totalParameterCount > 0 ? "请先移出类别中的参数" : "删除类别"}
              onClick={() => store.removeParameterGroup(group.id)}
            >
              删除类别
            </button>
          </div>
        </details>
      </div>
      {open && (
        <div
          className="grid gap-3 px-4 pb-4"
          style={{ gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))" }}
          onDragOver={event => {
            if (canReorder) event.preventDefault();
          }}
          onDrop={handleDrop}
        >
          {parameters.map((parameter, parameterIndex) => (
            <ParameterCard
              key={parameter.id}
              parameter={parameter}
              parameterIndex={parameterIndex}
              siblingCount={parameters.length}
              allGroups={allGroups}
              usage={usageByParameter.get(parameter.id) || { revenue: [], cost: [] }}
              selected={selectedParameterId === parameter.id}
              canReorder={canReorder}
              onSelect={() => onSelect(parameter.id)}
            />
          ))}
          {parameters.length === 0 && (
            <button
              type="button"
              className="min-h-20 rounded-lg bg-muted/30 text-xs font-semibold text-secondary-foreground ring-1 ring-border/70 hover:bg-muted/55"
              onClick={() => onAddParameter(group.id)}
            >
              + 在“{group.name}”中新增参数
            </button>
          )}
        </div>
      )}
    </section>
  );
}

function ParameterCard({
  parameter,
  parameterIndex,
  siblingCount,
  allGroups,
  usage,
  selected,
  canReorder,
  onSelect,
}: {
  parameter: AiComputeQuoteParameter;
  parameterIndex: number;
  siblingCount: number;
  allGroups: AiComputeQuoteParameterGroup[];
  usage: { revenue: string[]; cost: string[] };
  selected: boolean;
  canReorder: boolean;
  onSelect: () => void;
}) {
  const store = useAiComputeQuoteStore();
  const [descriptionOpen, setDescriptionOpen] = useState(false);
  const modified = isParameterModified(parameter);
  const handleDrop = (event: React.DragEvent<HTMLElement>) => {
    if (!canReorder) return;
    event.preventDefault();
    event.stopPropagation();
    const parameterId = event.dataTransfer.getData(PARAMETER_DRAG_TYPE);
    if (parameterId && parameterId !== parameter.id && parameter.groupId) {
      store.moveParameter(parameterId, parameter.groupId, parameter.id);
    }
  };

  return (
    <article
      className={`group relative rounded-lg bg-background/55 px-3 py-2.5 transition-colors ${
        selected ? "ring-2 ring-primary/25" : "ring-1 ring-border/75 hover:bg-card hover:ring-primary/20"
      }`}
      onClick={onSelect}
      onMouseEnter={onSelect}
      onDragOver={event => {
        if (canReorder) event.preventDefault();
      }}
      onDrop={handleDrop}
    >
      <div className="flex min-w-0 items-start gap-2">
        <button
          type="button"
          draggable={canReorder}
          className={`mt-0.5 flex h-5 w-4 shrink-0 items-center justify-center rounded text-[10px] text-muted-foreground ${
            canReorder ? "cursor-grab opacity-0 hover:bg-muted group-hover:opacity-100 active:cursor-grabbing" : "hidden"
          }`}
          aria-label={`拖拽排序参数 ${parameter.name}`}
          title="拖拽调整参数位置"
          onDragStart={event => {
            event.stopPropagation();
            event.dataTransfer.effectAllowed = "move";
            event.dataTransfer.setData(PARAMETER_DRAG_TYPE, parameter.id);
          }}
        >
          ⋮⋮
        </button>
        <Input
          aria-label={`${parameter.name}参数名称`}
          className="h-6 min-w-0 flex-1 border-0 bg-transparent px-0 py-0 text-xs font-bold shadow-none hover:border-0 focus-visible:border-0 focus-visible:ring-0"
          value={parameter.name}
          onFocus={onSelect}
          onChange={event => store.updateParameter(parameter.id, { name: event.target.value })}
        />
        <span className={`shrink-0 rounded-md px-2 py-0.5 text-[10px] font-bold ${
          modified ? "bg-destructive-soft text-destructive" : "bg-muted text-secondary-foreground"
        }`}>
          {modified ? "已修改" : "默认值"}
        </span>
        <button
          type="button"
          className={`flex h-5 w-5 shrink-0 items-center justify-center rounded transition-colors ${
            descriptionOpen ? "bg-primary-soft text-primary" : "text-muted-foreground hover:bg-muted"
          }`}
          aria-label={`${descriptionOpen ? "收起" : "查看"}参数说明 ${parameter.name}`}
          title="参数说明"
          onClick={event => {
            event.stopPropagation();
            setDescriptionOpen(value => !value);
          }}
        >
          <AppIcon name="info" size={13} />
        </button>
        <details className="relative" onClick={event => event.stopPropagation()}>
          <summary
            className="flex h-5 w-5 cursor-pointer list-none items-center justify-center rounded text-muted-foreground hover:bg-muted [&::-webkit-details-marker]:hidden"
            aria-label={`编辑参数设置 ${parameter.name}`}
            title="参数设置"
          >
            <AppIcon name="edit" size={13} />
          </summary>
          <div className="absolute right-0 top-7 z-30 w-72 rounded-lg bg-card p-3 shadow-lg ring-1 ring-border">
            <Field label="字段 key">
              <Input
                className="h-8 font-mono text-xs"
                value={parameter.key}
                disabled={isAiComputeStableIctParameter(parameter)}
                onChange={event => store.updateParameter(parameter.id, { key: event.target.value })}
              />
            </Field>
            {isAiComputeStableIctParameter(parameter) && (
              <div className="mt-1 text-[10px] font-semibold text-secondary-foreground">
                {isAiComputeProjectCycleParameter(parameter)
                  ? "项目周期字段 key 固定为 `years`，用于同步 ICT 项目周期。"
                  : "项目折现率字段 key 固定为 `discount_rate`，页面按百分数输入并同步 ICT。"}
              </div>
            )}
            <Field label="所属类别">
              <select
                className="h-8 rounded-md bg-card px-2 text-xs ring-1 ring-input"
                value={parameter.groupId || ""}
                onChange={event => store.moveParameter(parameter.id, event.target.value)}
              >
                {allGroups.map(group => <option key={group.id} value={group.id}>{group.name}</option>)}
              </select>
            </Field>
            <label className="mt-3 flex items-center gap-2 text-xs font-semibold text-secondary-foreground">
              <input
                type="checkbox"
                checked={parameter.isKey === true}
                onChange={event => store.updateParameter(parameter.id, { isKey: event.target.checked })}
              />
              关键参数
            </label>
            <label className="mt-2 flex items-center gap-2 text-xs font-semibold text-secondary-foreground">
              <input
                type="checkbox"
                checked={parameter.sensitivityEnabled !== false}
                onChange={event => store.updateParameter(parameter.id, { sensitivityEnabled: event.target.checked })}
              />
              参与敏感性分析
            </label>
            <div className="mt-3 grid grid-cols-2 gap-1">
              <Button
                type="button"
                variant="ghost"
                size="sm"
                disabled={parameterIndex === 0}
                onClick={() => store.moveParameterByOffset(parameter.id, -1)}
              >
                上移
              </Button>
              <Button
                type="button"
                variant="ghost"
                size="sm"
                disabled={parameterIndex === siblingCount - 1}
                onClick={() => store.moveParameterByOffset(parameter.id, 1)}
              >
                下移
              </Button>
              <Button
                type="button"
                variant="ghost"
                size="sm"
                disabled={isAiComputeStableIctParameter(parameter)}
                onClick={() => store.duplicateParameter(parameter.id, createId("parameter"))}
              >
                <AppIcon name="copy" size={14} />复制
              </Button>
              <Button
                type="button"
                variant="ghost"
                size="sm"
                className="hover:bg-destructive-soft hover:text-destructive"
                disabled={parameter.locked}
                onClick={() => store.removeParameter(parameter.id)}
              >
                <AppIcon name="delete" size={14} />删除
              </Button>
            </div>
          </div>
        </details>
      </div>
      <div className="mt-1 flex h-8 items-center gap-1.5">
        <Input
          aria-label={`${parameter.name}数值`}
          className="numeric-value h-8 min-w-0 flex-1 border-0 bg-transparent px-0 py-0 text-base font-bold shadow-none hover:border-0 focus-visible:border-0 focus-visible:ring-0"
          type="number"
          min={isAiComputeProjectCycleParameter(parameter)
            ? 1
            : isAiComputeDiscountRateParameter(parameter)
              ? 0
              : undefined}
          max={isAiComputeProjectCycleParameter(parameter)
            ? 10
            : isAiComputeDiscountRateParameter(parameter)
              ? 100
              : undefined}
          step={isAiComputeProjectCycleParameter(parameter)
            ? 1
            : isAiComputeDiscountRateParameter(parameter)
              ? 0.1
              : undefined}
          value={parameter.value}
          onFocus={onSelect}
          onChange={event => store.updateParameter(parameter.id, { value: toNumber(event.target.value) })}
        />
        <Input
          aria-label={`${parameter.name}单位`}
          className="h-8 min-w-0 shrink-0 whitespace-nowrap border-0 bg-transparent px-0 py-0 text-right text-xs font-semibold shadow-none hover:border-0 focus-visible:border-0 focus-visible:ring-0"
          style={{ width: `${getCompactInputWidth(parameter.unit || "")}px` }}
          value={parameter.unit || ""}
          placeholder="单位"
          onFocus={onSelect}
          onChange={event => store.updateParameter(parameter.id, { unit: event.target.value })}
        />
      </div>
      {descriptionOpen && (
        <div className="mt-2 rounded-md bg-muted/55 px-2.5 py-2 text-[11px] leading-5 text-secondary-foreground">
          <div>{getParameterImpact(parameter, usage)}</div>
          <div className="mt-0.5 text-[10px] text-muted-foreground">
            引用：收入 {usage.revenue.length} / 成本 {usage.cost.length}
          </div>
        </div>
      )}
    </article>
  );
}

function ParameterEditorModal({
  initialGroupId,
  onClose,
  onCreated,
}: {
  initialGroupId: string;
  onClose: () => void;
  onCreated: (parameterId: string) => void;
}) {
  const store = useAiComputeQuoteStore();
  const index = store.blueprint.parameters.length + 1;
  const [name, setName] = useState(`新参数 ${index}`);
  const [key, setKey] = useState(`custom_parameter_${index}`);
  const [value, setValue] = useState("0");
  const [unit, setUnit] = useState("");
  const [groupId, setGroupId] = useState(initialGroupId);
  const [isKey, setIsKey] = useState(true);
  const [sensitivityEnabled, setSensitivityEnabled] = useState(true);
  const [error, setError] = useState("");

  const submit = (event: React.FormEvent) => {
    event.preventDefault();
    const normalizedName = name.trim();
    const normalizedKey = key.trim();
    if (!normalizedName || !normalizedKey) {
      setError("参数名称和字段 key 不能为空。");
      return;
    }
    if (store.blueprint.parameters.some(parameter => parameter.key === normalizedKey)) {
      setError("字段 key 已存在，请使用唯一 key。");
      return;
    }
    const id = createId("parameter");
    store.addParameter({
      id,
      name: normalizedName,
      key: normalizedKey,
      value: toNumber(value),
      unit: unit.trim(),
      category: getLegacyParameterCategory(groupId),
      groupId,
      isKey,
      sensitivityEnabled,
    });
    onCreated(id);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-foreground/25 p-4 backdrop-blur-sm">
      <form className="w-full max-w-lg rounded-xl bg-card p-5 shadow-xl" onSubmit={submit}>
        <h2 className="text-section-title">新增参数</h2>
        <div className="mt-4 grid gap-3 sm:grid-cols-2">
          <Field label="参数名称">
            <Input value={name} onChange={event => setName(event.target.value)} />
          </Field>
          <Field label="字段 key">
            <Input className="font-mono" value={key} onChange={event => setKey(event.target.value)} />
          </Field>
          <Field label="初始值">
            <Input className="numeric-value" type="number" value={value} onChange={event => setValue(event.target.value)} />
          </Field>
          <Field label="单位">
            <Input value={unit} onChange={event => setUnit(event.target.value)} placeholder="如 元/台/月" />
          </Field>
          <Field label="所属类别">
            <select
              className="h-[var(--density-input-height)] rounded-md bg-card px-3 text-sm ring-1 ring-input"
              value={groupId}
              onChange={event => setGroupId(event.target.value)}
            >
              {store.blueprint.parameterGroups.map(group => (
                <option key={group.id} value={group.id}>{group.name}</option>
              ))}
            </select>
          </Field>
          <div className="flex flex-col justify-end gap-2 pb-1">
            <label className="flex items-center gap-2 text-xs font-semibold text-secondary-foreground">
              <input type="checkbox" checked={isKey} onChange={event => setIsKey(event.target.checked)} />
              关键参数
            </label>
            <label className="flex items-center gap-2 text-xs font-semibold text-secondary-foreground">
              <input
                type="checkbox"
                checked={sensitivityEnabled}
                onChange={event => setSensitivityEnabled(event.target.checked)}
              />
              参与敏感性分析
            </label>
          </div>
        </div>
        {error && <div className="mt-3 rounded-lg bg-destructive-soft px-3 py-2 text-xs text-destructive">{error}</div>}
        <div className="mt-5 flex justify-end gap-2">
          <Button type="button" variant="outline" onClick={onClose}>取消</Button>
          <Button type="submit">新增参数</Button>
        </div>
      </form>
    </div>
  );
}

function ParameterGroupEditorModal({
  mode,
  group,
  onClose,
  onSaved,
}: {
  mode: "create" | "edit";
  group?: AiComputeQuoteParameterGroup;
  onClose: () => void;
  onSaved: () => void;
}) {
  const store = useAiComputeQuoteStore();
  const [name, setName] = useState(group?.name || "");
  const [description, setDescription] = useState(group?.description || "");
  const [error, setError] = useState("");

  const submit = (event: React.FormEvent) => {
    event.preventDefault();
    const normalizedName = name.trim();
    if (!normalizedName) {
      setError("类别名称不能为空。");
      return;
    }
    if (store.blueprint.parameterGroups.some(candidate =>
      candidate.id !== group?.id && candidate.name.trim() === normalizedName
    )) {
      setError("类别名称已存在，请使用不同名称。");
      return;
    }
    if (mode === "edit" && group) {
      store.renameParameterGroup(group.id, normalizedName, description);
    } else {
      store.addParameterGroup({
        id: createId("parameter-group"),
        name: normalizedName,
        description: description.trim() || undefined,
        builtin: false,
      });
    }
    onSaved();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-foreground/25 p-4 backdrop-blur-sm">
      <form className="w-full max-w-md rounded-xl bg-card p-5 shadow-xl" onSubmit={submit}>
        <h2 className="text-section-title">{mode === "edit" ? "编辑类别" : "新增类别"}</h2>
        <div className="mt-4 space-y-3">
          <Field label="类别名称">
            <Input value={name} onChange={event => setName(event.target.value)} autoFocus />
          </Field>
          <Field label="类别说明（可选）">
            <Input value={description} onChange={event => setDescription(event.target.value)} />
          </Field>
        </div>
        {error && <div className="mt-3 rounded-lg bg-destructive-soft px-3 py-2 text-xs text-destructive">{error}</div>}
        <div className="mt-5 flex justify-end gap-2">
          <Button type="button" variant="outline" onClick={onClose}>取消</Button>
          <Button type="submit">{mode === "edit" ? "保存" : "新增类别"}</Button>
        </div>
      </form>
    </div>
  );
}

function SensitivityPanel({
  parameterId,
  min,
  max,
  step,
  rows,
  onParameterChange,
  onMinChange,
  onMaxChange,
  onStepChange,
}: {
  parameterId: string;
  min: number;
  max: number;
  step: number;
  rows: Array<ReturnType<typeof runAiComputeQuoteSensitivity>[number] & { ictResult?: IctResult }>;
  onParameterChange: (value: string) => void;
  onMinChange: (value: number) => void;
  onMaxChange: (value: number) => void;
  onStepChange: (value: number) => void;
}) {
  const store = useAiComputeQuoteStore();
  const activeParameter = store.blueprint.parameters.find(parameter => parameter.id === parameterId);
  return (
    <section className="mx-auto w-full max-w-[1180px]">
      <div className="mb-4 flex flex-wrap items-end gap-3">
        <div className="mr-auto">
          <h2 className="text-section-title">敏感性分析工具</h2>
          <p className="text-caption text-secondary-foreground">创建临时计算快照，不修改当前金额来源参数或正式 ICT 数据。</p>
        </div>
        <Field label="分析参数">
          <select
            className="h-[var(--density-input-height)] min-w-[190px] rounded-md bg-card px-3 text-sm ring-1 ring-ring/20"
            value={parameterId}
            onChange={event => onParameterChange(event.target.value)}
          >
            {store.blueprint.parameters.filter(parameter => parameter.sensitivityEnabled !== false).map(parameter => (
              <option key={parameter.id} value={parameter.id}>{parameter.name}</option>
            ))}
          </select>
        </Field>
        <Field label="最小值"><Input className="w-28 numeric-value" type="number" value={min} onChange={event => onMinChange(toNumber(event.target.value))} /></Field>
        <Field label="最大值"><Input className="w-28 numeric-value" type="number" value={max} onChange={event => onMaxChange(toNumber(event.target.value))} /></Field>
        <Field label="步长"><Input className="w-28 numeric-value" type="number" min="0" value={step} onChange={event => onStepChange(toNumber(event.target.value))} /></Field>
      </div>
      <div className="mb-4 rounded-lg bg-primary-soft/45 px-4 py-3 text-sm text-primary">
        当前分析参数：<span className="font-bold">{activeParameter?.name || "--"}</span>；
        {activeParameter
          ? `${getParameterImpact(activeParameter, { revenue: [], cost: [] })} 建议结合 ICT NPV 与净现值率变化判断风险。`
          : "请选择一个参与敏感性分析的参数。"}
      </div>
      {rows.length > 0 ? (
        <div className="overflow-x-auto rounded-lg bg-card/70 ring-1 ring-border/60">
          <table className="w-full text-left text-sm">
            <thead className="text-caption text-secondary-foreground">
              <tr>
                <th className="px-4">参数值</th><th className="px-4">总收入（万元）</th><th className="px-4">总成本（万元）</th>
                <th className="px-4">ICT 毛利率</th><th className="px-4">ICT NPV（万元）</th><th className="px-4">ICT 净现值率</th>
              </tr>
            </thead>
            <tbody>
              {rows.map(row => (
                <tr key={row.parameterValue} className="odd:bg-muted/40">
                  <td className="px-4 numeric-value">{formatNumber(row.parameterValue, 4)}</td>
                  <td className="px-4 numeric-value">{formatWan(row.totalRevenue)}</td>
                  <td className="px-4 numeric-value">{formatWan(row.totalCost)}</td>
                  <td className="px-4 numeric-value">
                    {row.ictResult ? `${formatNumber(Number(row.ictResult.margin_rate) * 100)}%` : "--"}
                  </td>
                  <td className="px-4 numeric-value">
                    {row.ictResult ? formatWan(Number(row.ictResult.npv)) : "--"}
                  </td>
                  <td className="px-4 numeric-value">
                    {row.ictResult ? `${formatNumber(Number(row.ictResult.npv_rate) * 100)}%` : "--"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <div className="rounded-lg bg-warning-soft p-4 text-sm text-warning-foreground">请检查范围和步长，最多生成 500 组结果。</div>
      )}
    </section>
  );
}

function LineItemPanel({
  side,
  title,
  syncStatus,
}: {
  side: AiComputeQuoteSide;
  title: string;
  syncStatus: "idle" | "syncing" | "synced" | "error" | "conflict";
}) {
  const store = useAiComputeQuoteStore();
  const items = side === "revenue" ? store.blueprint.revenueItems : store.blueprint.costItems;
  const total = items.reduce((sum, item) => sum + item.amountInclTax, 0);
  const projectCycleYears = getAiComputeProjectCycleYears(store.blueprint.parameters);

  const addItem = () => {
    const firstParameter = store.blueprint.parameters[0];
    const item: AiComputeQuoteLineItem = {
      id: createId(side),
      side,
      name: side === "revenue" ? "新增收入项" : "新增成本项",
      formula: {
        version: 2,
        tokens: firstParameter
          ? [{ type: "parameter", id: firstParameter.id, name: firstParameter.name }]
          : [{ type: "constant", value: 0 }],
      },
      amountInclTax: 0,
      amountExclTax: 0,
      taxRate: 6,
      enabled: true,
      outputEnabled: true,
      fundingPlan: {
        enabled: true,
        mode: "first_year",
        yearlyAmounts: Object.fromEntries(Array.from({ length: 10 }, (_, index) => [String(index + 1), 0])),
      },
    };
    store.addLineItem(item);
  };

  return (
    <section className="mx-auto w-full max-w-[1180px]">
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3 rounded-xl bg-muted/45 p-4">
        <div>
          <h2 className="text-section-title">{title}</h2>
          <p className="text-caption text-secondary-foreground">
            默认展示金额、占比、同步和计划摘要，展开后进行详细编辑。
            <span className="ml-2 numeric-value font-bold text-foreground">含税合计 {formatWan(total)} 万元</span>
          </p>
        </div>
        <Button size="sm" onClick={addItem}>新增{side === "revenue" ? "收入" : "成本"}</Button>
      </div>
      <div className="space-y-3">
        {items.map(item => (
          <CalculationItemCard
            key={item.id}
            item={item}
            totalAmount={total}
            projectCycleYears={projectCycleYears}
            syncStatus={syncStatus}
          />
        ))}
      </div>
    </section>
  );
}

function OutputPackageModal({
  onClose,
  onRequestIctOutput,
}: {
  onClose: () => void;
  onRequestIctOutput: () => void;
}) {
  const blueprint = useAiComputeQuoteStore(state => state.blueprint);
  const output = buildAiComputeQuoteOutput(blueprint);
  const fundingPlanOutput = buildAiComputeQuoteOutputFundingPlans(blueprint);
  const lineItems = [...blueprint.revenueItems, ...blueprint.costItems];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-foreground/25 p-6 backdrop-blur-sm">
      <div className="max-h-[86vh] w-full max-w-5xl overflow-y-auto rounded-xl bg-card p-6 shadow-xl">
        <div className="mb-4 flex items-start justify-between gap-4">
          <div>
            <h2 className="text-page-title">智算输出包</h2>
            <p className="text-body text-secondary-foreground">按 ICT 科目合并金额。此处仅预览，不写入正式测算。</p>
          </div>
          <Button variant="ghost" size="icon" onClick={onClose}><AppIcon name="close" /></Button>
        </div>
        <div className="overflow-x-auto rounded-lg bg-muted/40">
          <table className="w-full text-left text-sm">
            <thead className="text-caption text-secondary-foreground">
              <tr><th className="px-4">类型</th><th className="px-4">ICT 科目</th><th className="px-4">含税（万元）</th><th className="px-4">不含税（万元）</th><th className="px-4">来源</th></tr>
            </thead>
            <tbody>
              {output.map(item => (
                <tr key={`${item.side}-${item.ictSubjectCode}`} className="odd:bg-card/70">
                  <td className="px-4">{item.side === "revenue" ? "收入" : "成本"}</td>
                  <td className="px-4 font-semibold">{item.ictSubjectName}</td>
                  <td className="px-4 numeric-value">{formatWan(item.amountInclTax)}</td>
                  <td className="px-4 numeric-value">{formatWan(item.amountExclTax)}</td>
                  <td className="px-4 text-caption text-secondary-foreground">
                    {item.sourceLineItemIds.map(id => lineItems.find(lineItem => lineItem.id === id)?.name || id).join("、")}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {output.length === 0 && <div className="rounded-lg bg-warning-soft p-4 text-warning-foreground">没有已启用且完成科目映射的输出项。</div>}
        <div className="mt-5">
          <div className="mb-2">
            <h3 className="text-section-title">年度资金计划预览</h3>
            <p className="text-caption text-secondary-foreground">
              按 ICT 科目代码合并业务项年度含税金额；关闭计划的业务项不进入本表。
            </p>
          </div>
          <div className="overflow-x-auto rounded-lg bg-muted/40">
            <table className="min-w-[1500px] w-full text-left text-sm">
              <thead className="text-caption text-secondary-foreground">
                <tr>
                  <th className="px-3">类型</th>
                  <th className="px-3">ICT 科目</th>
                  {Array.from({ length: 10 }, (_, index) => (
                    <th key={index} className="px-3">第{index + 1}年</th>
                  ))}
                  <th className="px-3">合计</th>
                  <th className="px-3">来源项</th>
                </tr>
              </thead>
              <tbody>
                {fundingPlanOutput.map(item => {
                  const plannedTotal = Object.values(item.yearlyAmounts).reduce((sum, amount) => sum + amount, 0);
                  return (
                    <tr key={`${item.side}-${item.ictSubjectCode}`} className="odd:bg-card/70">
                      <td className="px-3">{item.side === "revenue" ? "收入" : "成本"}</td>
                      <td className="px-3 font-semibold">{item.ictSubjectName}</td>
                      {Array.from({ length: 10 }, (_, index) => (
                        <td key={index} className="numeric-value px-3">
                          {formatNumber(item.yearlyAmounts[String(index + 1)] || 0)}
                        </td>
                      ))}
                      <td className="numeric-value px-3">{formatNumber(plannedTotal)}</td>
                      <td className="px-3 text-caption text-secondary-foreground">
                        {item.sourceLineItemIds.map(id => lineItems.find(lineItem => lineItem.id === id)?.name || id).join("、")}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          {fundingPlanOutput.length === 0 && (
            <div className="rounded-lg bg-warning-soft p-4 text-warning-foreground">
              没有已启用计划且完成 ICT 科目映射的输出项。
            </div>
          )}
        </div>
        <div className="mt-5 flex justify-end gap-2">
          <Button variant="outline" onClick={onClose}>关闭</Button>
          <Button onClick={onRequestIctOutput}>重新同步 ICT</Button>
        </div>
      </div>
    </div>
  );
}

function IctExportConfirmModal({
  preview,
  error,
  syncing,
  requiresReload = false,
  onConfirm,
  onCancel,
}: {
  preview: AiComputeIctExportPreview;
  error: string | null;
  syncing: boolean;
  requiresReload?: boolean;
  onConfirm: () => void;
  onCancel: () => void;
}) {
  const [showAllYears, setShowAllYears] = useState(false);
  const visibleYears = showAllYears ? 10 : 2;
  const releasedCount = preview.rows.filter(row => row.syncStatus === "released_mapping").length;
  const controlledCount = preview.rows.length - releasedCount;
  const fundingModeLabel = (mode: AiComputeLineItemFundingPlanMode) => {
    if (mode === "first_year") return "首年一次性";
    if (mode === "even") return "平均分年";
    return "自定义年度";
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-foreground/25 p-6 backdrop-blur-sm">
      <div className="flex max-h-[88vh] w-full max-w-[1480px] flex-col overflow-hidden rounded-xl bg-card shadow-xl">
        <div className="shrink-0 p-6 pb-4">
          <h2 className="text-page-title">ICT 同步明细</h2>
          <p className="mt-1 text-body text-secondary-foreground">
            用户确认同步后，系统更新 ICT 科目、年度资金计划和正式效益指标。
          </p>
          <div className="mt-3 flex flex-wrap gap-2 text-caption">
            <span className="rounded-full bg-primary-soft px-2.5 py-1 font-bold text-primary">
              当前受控科目 {controlledCount} 项
            </span>
            {releasedCount > 0 && (
              <span className="rounded-full bg-warning-soft px-2.5 py-1 font-bold text-warning-foreground">
                释放旧科目 {releasedCount} 项
              </span>
            )}
            <span className="rounded-full bg-secondary px-2.5 py-1 font-bold text-secondary-foreground">
              Scenario: {preview.scenarioId}
            </span>
            <span className="rounded-full bg-secondary px-2.5 py-1 font-bold text-secondary-foreground">
              来源集合: {preview.blueprintId}
            </span>
            <span className="rounded-full bg-secondary px-2.5 py-1 font-bold text-secondary-foreground">
              项目周期: {preview.projectYears} 年
            </span>
            <span className="rounded-full bg-primary-soft px-2.5 py-1 text-[11px] font-bold text-primary">
              项目折现率: {formatNumber(preview.discountRate * 100)}%
            </span>
          </div>
        </div>

        <div className="min-h-0 flex-1 overflow-auto px-6">
          <div
            className="overflow-hidden rounded-lg bg-muted/40"
            style={{ minWidth: showAllYears ? 2160 : 1360 }}
          >
            <table className="w-full table-fixed text-left text-sm">
              <colgroup>
                <col className="w-[52px]" />
                <col className="w-[170px]" />
                <col className="w-[120px]" />
                <col className="w-[120px]" />
                <col className="w-[128px]" />
                {Array.from({ length: visibleYears }, (_, index) => (
                  <col key={index} className="w-[96px]" />
                ))}
                {!showAllYears && <col className="w-[150px]" />}
                <col className="w-[108px]" />
                <col className="w-[230px]" />
                <col className="w-[240px]" />
              </colgroup>
              <thead className="text-caption text-secondary-foreground">
                <tr>
                  <th className="px-3 py-3">类型</th>
                  <th className="px-3 py-3">ICT 科目</th>
                  <th className="px-3 py-3">原金额</th>
                  <th className="px-3 py-3">智算金额</th>
                  <th className="px-3 py-3">写入后金额</th>
                  {Array.from({ length: visibleYears }, (_, index) => (
                    <th key={index} className="px-3 py-3">第{index + 1}年</th>
                  ))}
                  {!showAllYears && <th className="px-3 py-3">其余年度</th>}
                  <th className="px-3 py-3">计划模式</th>
                  <th className="px-3 py-3">来源项</th>
                  <th className="px-3 py-3">状态</th>
                </tr>
              </thead>
              <tbody>
                {preview.rows.map(row => (
                  <tr key={`${row.side}:${row.ictSubjectCode}`} className="odd:bg-card/70">
                    <td className="px-3 py-4 align-top">{row.side === "revenue" ? "收入" : "成本"}</td>
                    <td className="px-3 py-4 align-top font-semibold">{row.ictSubjectName}</td>
                    <td className="numeric-value px-3 py-4 align-top">{formatWan(row.originalAmount)} 万元</td>
                    <td className="numeric-value px-3 py-4 align-top text-primary">{formatWan(row.quoteAmount)} 万元</td>
                    <td className="numeric-value px-3 py-4 align-top font-bold">{formatWan(row.writtenAmount)} 万元</td>
                    {row.yearlyAmounts.slice(0, visibleYears).map((amount, index) => (
                      <td key={index} className="numeric-value px-3 py-4 align-top">{formatWan(amount)}</td>
                    ))}
                    {!showAllYears && (
                      <td className="px-3 py-4 align-top text-caption text-secondary-foreground">
                        第3-10年合计 {formatWan(row.yearlyAmounts.slice(2).reduce((sum, amount) => sum + amount, 0))} 万元
                      </td>
                    )}
                    <td className="px-3 py-4 align-top text-caption text-secondary-foreground">
                      {row.syncStatus === "released_mapping"
                        ? "释放旧计划"
                        : row.syncStatus === "zeroed_absent"
                          ? "智算未输出"
                          : row.fundingPlanModes.length > 0
                            ? row.fundingPlanModes.map(fundingModeLabel).join("、")
                            : "--"}
                    </td>
                    <td className="px-3 py-4 align-top text-caption text-secondary-foreground">
                      <div className="whitespace-normal break-words leading-5">
                        {row.sourceLineItemNames.join("、") || "--"}
                      </div>
                    </td>
                    <td className="px-3 py-4 align-top text-caption">
                      <div className={
                        row.syncStatus === "ready"
                          ? "text-success-foreground"
                          : row.syncStatus === "released_mapping"
                            ? "text-destructive"
                            : "text-warning-foreground"
                      }>
                        <span className="whitespace-nowrap">
                          {row.syncStatus === "ready"
                            ? row.originalAmount === row.writtenAmount ? "金额一致" : "将覆盖"
                            : row.syncStatus === "zeroed_error"
                              ? "异常写零"
                              : row.syncStatus === "zeroed_absent"
                                ? "智算无输出，写 0"
                              : row.syncStatus === "paused_override"
                                ? "ICT 人工覆盖"
                                : row.syncStatus === "released_mapping"
                                  ? "释放旧映射"
                                  : "需重载"}
                        </span>
                      </div>
                      {row.syncMessages?.map(message => (
                        <div key={message} className="mt-1 whitespace-normal break-words leading-5">
                          {message}
                        </div>
                      ))}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {preview.skippedUnmappedItems.length > 0 && (
            <div className="mt-4 rounded-lg bg-warning-soft p-3 text-caption text-warning-foreground">
              <div className="font-extrabold">以下项目未映射 ICT 科目，不会写入：</div>
              <div className="mt-1">{preview.skippedUnmappedItems.map(item => item.name).join("、")}</div>
            </div>
          )}
          {preview.skippedItems.length > 0 && (
            <div className="mt-3 rounded-lg bg-muted/60 p-3 text-caption text-secondary-foreground">
              <div className="font-extrabold text-foreground">其他未参与项目：</div>
              <div className="mt-1">
                {preview.skippedItems.map(item => `${item.name}（${item.reason}）`).join("、")}
              </div>
            </div>
          )}
          {preview.rows.length === 0 && (
            <div className="rounded-lg bg-warning-soft p-4 font-semibold text-warning-foreground">
              当前没有可写入的已启用、已映射且资金计划有效的智算项。
            </div>
          )}
          {error && <div className="mt-3 rounded-lg bg-destructive-soft p-3 text-caption text-destructive">{error}</div>}
        </div>

        <div className="flex shrink-0 justify-end gap-2 p-6 pt-4">
          <Button variant="outline" disabled={syncing} onClick={onCancel}>关闭</Button>
          <Button variant="secondary" onClick={() => setShowAllYears(value => !value)}>
            {showAllYears ? "收起差异" : "查看差异"}
          </Button>
          <Button disabled={syncing || preview.rows.length === 0} onClick={onConfirm}>
            {syncing ? "同步中..." : requiresReload ? "重新加载冲突并覆盖" : "智算完全覆盖 ICT"}
          </Button>
        </div>
      </div>
    </div>
  );
}
