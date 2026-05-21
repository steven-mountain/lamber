import { useEffect, useState, useCallback } from "react";
import WorkspaceHeader from "../components/WorkspaceHeader";
import AppIcon from "../components/icons/AppIcon";
import TemplateForms from "./TemplateForms";
import { useNavigationStore } from "../store/useNavigationStore";
import { validateFinancialData, type ValidationReport } from "../lib/financeValidator";
import { useAiContextStore } from "../store/useAiContextStore";
import { AI_CONTEXT_KEY, buildAiContextKey } from "../utils/aiContextKeys";
import { useIctState } from "../hooks/useIctState";
import { useIctCalculations } from "../hooks/useIctCalculations";
import { IctBasicInfo } from "../components/IctBasicInfo";
import { IctCashflowTable } from "../components/IctCashflowTable";
import { IctMetricsDashboard } from "../components/IctMetricsDashboard";
import {
  projectService,
  type Project,
  type BenefitAnalysisScheme,
  type BenefitAnalysisSnapshot
} from "../utils/projectService";

export default function IctLifecycle() {
  const { activeProjectId, activeSchemeId, entrySource, navigateTo } = useNavigationStore();
  const state = useIctState();
  const calculations = useIctCalculations(state);

  const [projects, setProjects] = useState<Project[]>([]);
  const [activeProject, setActiveProject] = useState<Project | null>(null);
  const [activeScheme, setActiveScheme] = useState<BenefitAnalysisScheme | null>(null);
  const [activeSnapshot, setActiveSnapshot] = useState<BenefitAnalysisSnapshot | null>(null);
  const [pendingNewSchemeName, setPendingNewSchemeName] = useState<string | null>(null);

  const [showSaveAsModal, setShowSaveAsModal] = useState(false);
  const [saveAsSchemeName, setSaveAsSchemeName] = useState("");

  useEffect(() => {
    projectService.getProjects().then(setProjects).catch(console.error);
  }, [activeProjectId]);


  const fillCalculatorState = useCallback((params: any) => {
    if (!params) return;

    const restoreItem = (item: any, defaultTax = 6) => ({
      incl: item ? Number(item.incl_tax) : 0,
      tax: item ? Number(item.tax_rate) : defaultTax,
      excl: item ? Number((Number(item.incl_tax) / (1 + Number(item.tax_rate) / 100)).toFixed(2)) : 0
    });

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

    const revItRestored = {
      integration: restoreItem(params.rev_it_integration, 6),
      maintenance: restoreItem(params.rev_it_maintenance, 6),
      device_sales: restoreItem(params.rev_it_device_sales, 13),
      device_lease: restoreItem(params.rev_it_device_lease, 13),
      other: restoreItem(params.rev_it_other, 6),
      cloud: restoreItem(params.rev_it_cloud, 6),
    };
    const revCtRestored = {
      line: restoreItem(params.rev_ct_line, 9),
      product: restoreItem(params.rev_ct_product, 6),
    };
    const revNonItCtRestored = restoreItem(params.rev_non_it_ct, 9);

    const costItRestored = {
      device: restoreItem(params.cost_it_device, 13),
      construction: restoreItem(params.cost_it_construction, 9),
      survey: restoreItem(params.cost_it_survey, 6),
      integration: restoreItem(params.cost_it_integration, 6),
      other: restoreItem(params.cost_it_other, 6),
      maintenance: restoreItem(params.cost_it_maintenance, 6),
      running: restoreItem(params.cost_it_running, 13),
      bidding: restoreItem(params.cost_it_bidding, 6),
      design_eval: restoreItem(params.cost_it_design_eval, 6),
      audit: restoreItem(params.cost_it_audit, 6),
    };
    const costCtRestored = {
      construction: restoreItem(params.cost_ct_construction, 9),
      maintenance: restoreItem(params.cost_ct_maintenance, 9),
      other: restoreItem(params.cost_ct_other, 6),
      bandwidth: restoreItem(params.cost_ct_bandwidth, 9),
      renewal: restoreItem(params.cost_ct_renewal, 9),
    };
    const costMixRestored = {
      non_it_ct: restoreItem(params.cost_non_it_ct, 9),
      marketing: restoreItem(params.cost_mix_marketing, 6),
      channel: restoreItem(params.cost_mix_channel, 6),
      other: restoreItem(params.cost_mix_other, 6),
    };

    state.setRevIt(revItRestored);
    state.setRevCt(revCtRestored);
    state.setRevNonItCt(revNonItCtRestored);
    state.setCostIt(costItRestored);
    state.setCostCt(costCtRestored);
    state.setCostMix(costMixRestored);

    if (params.ignore_tail_difference) {
      state.setIgnoredTailValue(params.tail_difference_value || "0");
      state.setIgnoredDataHash(JSON.stringify({
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
  }, [state]);

  const loadProjectContext = useCallback(async (pId: string | null, sId?: string | null) => {
    const targetProjectId = pId || null;
    const targetSchemeId = sId || null;

    if (!targetProjectId) {
      localStorage.removeItem("lamber_active_project_id");
      localStorage.removeItem("lamber_active_scheme_id");
      setActiveProject(null);
      setActiveScheme(null);
      setActiveSnapshot(null);
      setPendingNewSchemeName(null);
      return;
    }

    try {
      const project = await projectService.getProject(targetProjectId);
      if (!project) {
        localStorage.removeItem("lamber_active_project_id");
        localStorage.removeItem("lamber_active_scheme_id");
        setActiveProject(null);
        setActiveScheme(null);
        setActiveSnapshot(null);
        setPendingNewSchemeName(null);
        return;
      }

      localStorage.setItem("lamber_active_project_id", project.id);
      setActiveProject(project);

      const newSchemeNameLocal = localStorage.getItem("lamber_new_scheme_name");
      if (newSchemeNameLocal) {
        setPendingNewSchemeName(newSchemeNameLocal);
        localStorage.removeItem("lamber_new_scheme_name");
        setActiveScheme(null);
        setActiveSnapshot(null);

        state.setProjName(project.name);
        state.setCustomerName(project.customer_name);
        if (project.discount_rate > 0) state.setDiscountRate(project.discount_rate);
        if (project.project_years > 0) state.setProjectYears(project.project_years);
        if (project.cashflow_model) state.setCashflowModel(project.cashflow_model as any);
        return;
      }

      setPendingNewSchemeName(null);

      const projectSchemes = await projectService.getSchemes(project.id);
      let schemeToSelect: BenefitAnalysisScheme | null = null;
      if (targetSchemeId) {
        schemeToSelect = projectSchemes.find(s => s.id === targetSchemeId) || null;
      }
      if (!schemeToSelect) {
        schemeToSelect = projectSchemes.find(s => s.id === project.default_scheme_id) || projectSchemes[0] || null;
      }

      if (schemeToSelect) {
        setActiveScheme(schemeToSelect);
        localStorage.setItem("lamber_active_scheme_id", schemeToSelect.id);

        const snapshots = await projectService.getSnapshots(schemeToSelect.id);
        if (snapshots.length > 0) {
          const latestSnap = snapshots.reduce((latest, current) => current.version > latest.version ? current : latest, snapshots[0]);
          setActiveSnapshot(latestSnap);
          fillCalculatorState(latestSnap.input_params);
        } else {
          setActiveSnapshot(null);
          state.setProjName(project.name);
          state.setCustomerName(project.customer_name);
          if (project.discount_rate > 0) state.setDiscountRate(project.discount_rate);
          if (project.project_years > 0) state.setProjectYears(project.project_years);
          if (project.cashflow_model) state.setCashflowModel(project.cashflow_model as any);
        }
      } else {
        setActiveScheme(null);
        setActiveSnapshot(null);
        state.setProjName(project.name);
        state.setCustomerName(project.customer_name);
        if (project.discount_rate > 0) state.setDiscountRate(project.discount_rate);
        if (project.project_years > 0) state.setProjectYears(project.project_years);
        if (project.cashflow_model) state.setCashflowModel(project.cashflow_model as any);
      }
    } catch (err) {
      console.error("Failed to load project context:", err);
    }
  }, [state, fillCalculatorState]);

  useEffect(() => {
    loadProjectContext(activeProjectId, activeSchemeId);
  }, [activeProjectId, activeSchemeId]);

  const handleSaveToCurrent = async () => {
    if (!activeProject) return;
    const schemeId = pendingNewSchemeName ? null : (activeScheme?.id || activeProject.default_scheme_id || null);
    const schemeName = pendingNewSchemeName || activeScheme?.name || activeProject.name || "默认测算方案";

    try {
      const payload = calculations.buildInputDataPayload();

      const updatedProj = await projectService.saveBenefitScheme(
        activeProject.id,
        schemeId,
        schemeName,
        payload,
        metrics,
        pendingNewSchemeName ? true : false
      );

      setPendingNewSchemeName(null);
      const newSchemeId = updatedProj.default_scheme_id || null;
      await loadProjectContext(activeProject.id, newSchemeId);
      alert("保存测算成功！已更新项目指标并生成历史记录。");
    } catch (error) {
      console.error("保存失败:", error);
      alert("保存失败: " + error);
    }
  };

  const handleSaveAsNew = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!activeProject || !saveAsSchemeName.trim()) return;

    try {
      const payload = calculations.buildInputDataPayload();
      const updatedProj = await projectService.saveBenefitScheme(
        activeProject.id,
        null,
        saveAsSchemeName.trim(),
        payload,
        metrics,
        true
      );

      setShowSaveAsModal(false);
      setSaveAsSchemeName("");
      setPendingNewSchemeName(null);

      const newSchemeId = updatedProj.default_scheme_id || null;
      await loadProjectContext(activeProject.id, newSchemeId);
      alert("另存为新方案成功！");
    } catch (error) {
      console.error("另存为失败:", error);
      alert("另存为失败: " + error);
    }
  };

  const {
    activeTab, setActiveTab,
    projName,
    customerName,
    projectYears,
    quickRevTotal, setQuickRevTotal,
    quickRevProduct, setQuickRevProduct,
    quickCostTotal, setQuickCostTotal,
    quickCostProduct, setQuickCostProduct,
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
  } = state;

  const {
    metrics,
    selQuote,
    selMarkup,
    selActualCost,
    selFee,
    selLimit,
    revMode, setRevMode,
    revTargetType, setRevTargetType,
    revTargetValue, setRevTargetValue,
    performReverseCalculation,
    applySelectionLimit,
    handleSelFeeChange,
  } = calculations;

  const setActiveModule = useAiContextStore(storeState => storeState.setActiveModule);

  useEffect(() => {
    setActiveModule('ict');
  }, [setActiveModule]);

  const handleTabSwitch = (tab: string, templateName?: string, forceIgnore = false) => {
    const currentHash = JSON.stringify({ revIt, revCt, revNonItCt, costIt, costCt, costMix });

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

  const renderTaxGroup = (title: string, groupId: string, groupState: any, items: {key: string, label: string}[]) => (
    <div className="table-card bg-card border border-border rounded-xl p-6 shadow-sm mb-6">
      <h3 className="font-bold text-lg mb-4">{title}</h3>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {items.map(item => {
          const itemErr = reconciliationErrors.find(e => e.key === `${groupId}.${item.key}`);
          return (
            <div key={item.key} className="flex flex-col gap-1">
              <label className="text-sm font-semibold text-secondary-foreground">{item.label}</label>
              <div className="flex gap-2">
                <input type="number" placeholder="含税" className="w-full bg-card border border-input px-3 py-2 rounded-md outline-none text-sm" value={groupState[item.key].incl === 0 ? "" : groupState[item.key].incl} onChange={e => updateTaxItem(groupId, item.key, 'incl', Number(e.target.value))} />
                <input type="number" placeholder="税率" className="w-20 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm" value={groupState[item.key].tax} onChange={e => updateTaxItem(groupId, item.key, 'tax', Number(e.target.value))} />
                <input type="number" placeholder="不含税" className={`w-full bg-card border px-3 py-2 rounded-md outline-none text-sm focus:border-ring ${itemErr ? 'border-red-500 ring-1 ring-red-500' : 'border-input'}`} value={groupState[item.key].excl === 0 ? "" : groupState[item.key].excl} onChange={e => updateTaxItem(groupId, item.key, 'excl', Number(e.target.value))} />
              </div>
              {itemErr && <span className="text-[10px] text-red-500 font-bold">校验失败：偏离 {itemErr.difference} 元，要求：{itemErr.expectedExcl} 元</span>}
            </div>
          );
        })}
      </div>
    </div>
  );

  return (
    <div className="flex flex-col flex-1 animate-in fade-in duration-300 h-full overflow-hidden">
      <WorkspaceHeader
        moduleId="ict_lifecycle"
        title="ICT项目全生命周期"
        onBack={() => {
          if (entrySource === "project_board") {
            navigateTo("project_board", activeProjectId, activeSchemeId);
          } else {
            navigateTo("hub");
          }
        }}
        onPathChange={() => loadTemplates()}
      />

      <div className="flex flex-1 overflow-hidden">
        <div className="w-[260px] bg-muted p-6 overflow-y-auto flex flex-col gap-4 border-r border-border shrink-0">
          <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70 mb-1">测算流程</h3>
          <div className="flex flex-col gap-1">
            <button className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${activeTab === 'basic' ? 'bg-blue-50 text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`} onClick={() => handleTabSwitch("basic")}><AppIcon name="project" size={18} /> 项目概况与参数</button>
            <button className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${activeTab === 'revenue' ? 'bg-blue-50 text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`} onClick={() => handleTabSwitch("revenue")}><AppIcon name="revenue" size={18} /> 收入侧测算</button>
            <button className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${activeTab === 'cost' ? 'bg-blue-50 text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`} onClick={() => handleTabSwitch("cost")}><AppIcon name="cost" size={18} /> 支出侧测算</button>
            <button className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${activeTab === 'cashflow' ? 'bg-blue-50 text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`} onClick={() => handleTabSwitch("cashflow")}><AppIcon name="cashflow" size={18} /> 10年现金流推演</button>
          </div>
          <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70 mt-6 pt-4 border-t border-border mb-2">一键生成全流程文档</h3>
          <div className="flex flex-col gap-2">
            {templates.length === 0 ? <span className="text-xs text-secondary-foreground px-4">未找到模板文件</span> :
              templates.map(t => {
                const isActive = selectedTemplate === t && activeTab === 'generate';
                return (
                  <button
                    key={t}
                    className={`relative overflow-hidden px-4 py-3 rounded-lg text-sm flex items-start gap-2.5 transition-all text-left border-b border-border/60 shadow-sm ${isActive ? 'bg-blue-50 text-primary font-bold border-blue-200 shadow-sm' : 'text-secondary-foreground font-semibold hover:bg-primary/10 hover:text-primary'}`}
                    onClick={() => handleTabSwitch("generate", t)}
                  >
                    {isActive && <div className="absolute left-0 top-0 bottom-0 w-1 bg-primary" />}
                    <AppIcon name={t.endsWith('.xlsx') ? "spreadsheet" : "document"} size={18} className="mt-0.5" />
                    <span className="whitespace-normal break-words leading-relaxed flex-1">
                      {t.replace('.docx', '').replace('.xlsx', '')}
                    </span>
                  </button>
                );
              })
            }
          </div>
        </div>

        <div className="flex-1 p-6 overflow-y-auto bg-background flex flex-col">
          {activeProject ? (
            <div className="bg-card border border-border rounded-xl p-4 mb-6 flex flex-col lg:flex-row lg:items-center justify-between gap-4 shadow-sm animate-in slide-in-from-top duration-300">
              <div className="flex items-center gap-3">
                <div className="bg-primary/10 p-2.5 rounded-lg text-primary shrink-0">
                  <AppIcon name="project" size={20} />
                </div>
                <div>
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="font-extrabold text-foreground text-sm">{activeProject.name}</span>
                    <span className="text-xs text-secondary-foreground">({activeProject.customer_name})</span>
                    <span className={`text-[10px] px-2 py-0.5 rounded-full font-bold ${
                      activeProject.benefit_status === 'normal'
                        ? 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                        : activeProject.benefit_status === 'outdated'
                        ? 'bg-amber-50 text-amber-700 border border-amber-200'
                        : 'bg-slate-100 text-slate-700 border border-slate-200'
                    }`}>
                      效益状态: {activeProject.benefit_status === 'normal' ? '最新' : activeProject.benefit_status === 'outdated' ? '已失效' : '未测算'}
                    </span>
                  </div>
                  <div className="text-xs text-secondary-foreground mt-0.5">
                    {pendingNewSchemeName ? (
                      <span>拟新建方案: <span className="font-semibold text-primary">{pendingNewSchemeName}</span> (未保存)</span>
                    ) : (
                      <span>当前方案: <span className="font-semibold text-primary">{activeScheme?.name || "默认方案"}</span> {activeSnapshot ? `(v${activeSnapshot.version})` : ''}</span>
                    )}
                  </div>
                </div>
              </div>
              <div className="flex flex-wrap items-center gap-2.5 self-end lg:self-auto shrink-0">
                <select
                  onChange={(e) => {
                    const pid = e.target.value;
                    if (pid === "free") {
                      navigateTo("ict_lifecycle", null, null);
                    } else {
                      navigateTo("ict_lifecycle", pid, null);
                    }
                  }}
                  value={activeProject.id}
                  className="bg-card border border-input px-3 py-1.5 rounded-lg text-xs outline-none focus:border-ring font-semibold text-foreground cursor-pointer min-w-[160px] mr-1"
                >
                  <option value="free">断开关联 (进入自由测算)</option>
                  {projects.map(p => (
                    <option key={p.id} value={p.id}>{p.name} ({p.customer_name})</option>
                  ))}
                </select>

                <button
                  id="save_benefit_btn"
                  onClick={handleSaveToCurrent}
                  className="bg-primary text-primary-foreground font-bold px-4 py-2 rounded-lg text-xs hover:bg-primary/90 transition-all shadow-sm flex items-center gap-1.5 active:scale-[0.98]"
                >
                  <AppIcon name="save" size={14} /> 保存到当前项目
                </button>
                <button
                  id="save_as_new_benefit_btn"
                  onClick={() => {
                    setSaveAsSchemeName(activeScheme?.name ? `${activeScheme.name}_复本` : "新方案");
                    setShowSaveAsModal(true);
                  }}
                  className="bg-card border border-border text-foreground hover:bg-secondary font-bold px-4 py-2 rounded-lg text-xs transition-all shadow-sm flex items-center gap-1.5 active:scale-[0.98]"
                >
                  <AppIcon name="copy" size={14} /> 另存为新方案
                </button>
              </div>
            </div>
          ) : (
            <div className="bg-card border border-border rounded-xl p-4 mb-6 flex flex-col sm:flex-row sm:items-center justify-between gap-4 shadow-sm">
              <div className="flex items-center gap-3">
                <div className="bg-secondary p-2.5 rounded-lg text-primary shrink-0">
                  <AppIcon name="project" size={20} />
                </div>
                <div>
                  <div className="font-extrabold text-foreground text-sm flex items-center gap-2">
                    <span>自由测算模式</span>
                    <span className="text-[10px] bg-secondary text-secondary-foreground font-bold px-2 py-0.5 rounded-full">未绑定项目</span>
                  </div>
                  <div className="text-xs text-secondary-foreground mt-0.5">你可以输入参数进行效益测算。如需保存，请在右侧选择关联一个项目：</div>
                </div>
              </div>
              <div className="flex items-center gap-2 self-end sm:self-auto shrink-0">
                <select
                  onChange={(e) => {
                    const pid = e.target.value;
                    if (pid) {
                      navigateTo("ict_lifecycle", pid, null);
                    }
                  }}
                  value=""
                  className="bg-card border border-input px-3 py-1.5 rounded-lg text-xs outline-none focus:border-ring font-semibold text-foreground cursor-pointer min-w-[200px]"
                >
                  <option value="" disabled>-- 关联已有项目 --</option>
                  {projects.map(p => (
                    <option key={p.id} value={p.id}>{p.name} ({p.customer_name})</option>
                  ))}
                </select>
              </div>
            </div>
          )}

          {activeTab === "basic" && (
            <IctBasicInfo state={state} calculations={calculations} />
          )}

          {activeTab === "revenue" && (
            <div>
              <div className="mb-6 border border-border bg-card p-5 rounded-xl shadow-sm">
                <div className="flex items-center justify-between mb-2">
                  <h3 className="font-bold text-primary text-base flex items-center gap-2">
                    <AppIcon name="quickAction" size={18} /> 快捷收入拆分计算器 (融入自有产品测算)
                  </h3>
                  <span className="text-xs bg-primary/10 text-primary font-bold px-2.5 py-1 rounded-full border border-primary/20">
                    自有产品占有率要求 1%
                  </span>
                </div>
                <p className="text-xs text-secondary-foreground mb-4 leading-relaxed">
                  按项目全生命周期要求，项目需融入自有产品，且自有产品占有率要求达到收入含税总金额的 1%。
                  当您填写含税总收入（项目总金额）后，程序将默认自动计算出 1% 的自有产品含税金额。您也可按需手工调整。
                </p>
                <div className="flex gap-4 items-end">
                  <div className="flex flex-col gap-1.5 flex-1">
                    <label className="text-xs font-bold text-foreground">含税总收入 (项目总金额)</label>
                    <input
                      type="number"
                      placeholder="输入总金额"
                      value={quickRevTotal}
                      onChange={e => {
                        const val = e.target.value;
                        setQuickRevTotal(val);
                        if (val && !isNaN(Number(val))) {
                          const productVal = (Number(val) * 0.01).toFixed(2);
                          setQuickRevProduct(productVal);
                        } else {
                          setQuickRevProduct("");
                        }
                      }}
                      className="bg-card border border-input px-3 py-2 rounded-md outline-none focus:border-ring text-sm font-semibold shadow-sm"
                    />
                  </div>
                  <div className="flex flex-col gap-1.5 flex-1">
                    <label className="text-xs font-bold text-foreground flex items-center justify-between">
                      <span>含税产品收入 (自有产品)</span>
                      <span className="text-[10px] text-primary font-mono font-medium">默认1%</span>
                    </label>
                    <input
                      type="number"
                      placeholder="产品金额"
                      value={quickRevProduct}
                      onChange={e => setQuickRevProduct(e.target.value)}
                      className="bg-card border border-input px-3 py-2 rounded-md outline-none focus:border-ring text-sm font-semibold shadow-sm"
                    />
                  </div>
                  <div className="flex flex-col gap-1.5 flex-1">
                    <label className="text-xs font-bold text-primary">系统集成服务收入 (自动扣减)</label>
                    <input
                      type="number"
                      disabled
                      value={Math.max(0, (Number(quickRevTotal)||0) - (Number(quickRevProduct)||0)).toFixed(2)}
                      className="bg-slate-50 border border-input px-3 py-2 rounded-md outline-none text-sm font-bold text-foreground shadow-sm"
                    />
                  </div>
                  <button
                    onClick={() => {
                      const integration = Math.max(0, (Number(quickRevTotal)||0) - (Number(quickRevProduct)||0));
                      updateTaxItem('revIt', 'integration', 'incl', Number(integration.toFixed(2)));
                      if (quickRevProduct) updateTaxItem('revCt', 'product', 'incl', Number(Number(quickRevProduct).toFixed(2)));
                    }}
                    className="inline-flex items-center gap-2 bg-primary text-primary-foreground font-bold px-5 py-2 rounded-md text-sm hover:bg-primary/90 transition-all shadow-sm hover:shadow active:scale-[0.98]"
                  >
                    <AppIcon name="download" size={16} /> 一键填入表单
                  </button>
                </div>
              </div>
              <div className="mb-4 text-xs text-blue-700 bg-blue-50 p-3 rounded-lg border border-blue-200">
                <span className="inline-flex items-start gap-2"><AppIcon name="info" size={16} className="mt-0.5" /> <span>提示：在「CT收入」中填写的产品或专线含税收入，将会自动【1:1平过】填入对应的「CT投入」中。</span></span>
              </div>
              {renderTaxGroup("IT/移动云收入", 'revIt', revIt, [{key: 'integration', label: '系统集成服务收入'}, {key: 'maintenance', label: '维保收入'}, {key: 'device_sales', label: '设备销售收入'}, {key: 'device_lease', label: '设备租赁收入'}, {key: 'other', label: '其他收入'}, {key: 'cloud', label: '移动云-定制化收入'}])}
              {renderTaxGroup("CT收入", 'revCt', revCt, [{key: 'line', label: '专线收入'}, {key: 'product', label: '产品收入'}])}
              {renderTaxGroup("非IT/CT收入", 'revNonItCt', { item: revNonItCt }, [{key: 'item', label: '工程施工收入等'}])}
            </div>
          )}

          {activeTab === "cost" && (
            <div>
              <div className="mb-6 border border-border bg-card p-4 rounded-xl">
                <h3 className="font-bold text-primary mb-2">快捷投入拆分计算器</h3>
                <p className="text-xs text-secondary-foreground mb-4">输入含税总投入与含税产品投入，计算“系统集成服务投入”并一键填入下方表单。</p>
                <div className="flex gap-4 items-end">
                  <div className="flex flex-col gap-1 flex-1">
                    <label className="text-xs font-semibold">含税总投入</label>
                    <input type="number" value={quickCostTotal} onChange={e => setQuickCostTotal(e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md outline-none text-sm" />
                  </div>
                  <div className="flex flex-col gap-1 flex-1">
                    <label className="text-xs font-semibold">含税产品投入</label>
                    <input type="number" value={quickCostProduct} onChange={e => setQuickCostProduct(e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md outline-none text-sm" />
                  </div>
                  <div className="flex flex-col gap-1 flex-1">
                    <label className="text-xs font-semibold text-primary">系统集成服务投入 (自动)</label>
                    <input type="number" disabled value={Math.max(0, (Number(quickCostTotal)||0) - (Number(quickCostProduct)||0))} className="bg-slate-50 border border-input px-3 py-2 rounded-md outline-none text-sm font-bold text-primary" />
                  </div>
                  <button onClick={() => {
                     const integration = Math.max(0, (Number(quickCostTotal)||0) - (Number(quickCostProduct)||0));
                     updateTaxItem('costIt', 'integration', 'incl', integration);
                     if (quickCostProduct) updateTaxItem('costIt', 'device', 'incl', Number(quickCostProduct));
                  }} className="inline-flex items-center gap-2 bg-primary text-primary-foreground font-bold px-4 py-2 rounded-md text-sm hover:bg-primary/90 transition-colors"><AppIcon name="download" size={16} /> 一键填入</button>
                </div>
              </div>
              {renderTaxGroup("IT/移动云投入", 'costIt', costIt, [{key: 'device', label: '主要设备/甲供材料'}, {key: 'construction', label: '施工'}, {key: 'survey', label: '勘察设计/预备费'}, {key: 'integration', label: '集成服务'}, {key: 'other', label: '其他投入'}, {key: 'maintenance', label: '维护费用'}, {key: 'running', label: '其他运行支出（电费等）'}, {key: 'bidding', label: '中标服务费'}, {key: 'design_eval', label: '设计院成本评估费'}, {key: 'audit', label: '第三方审计评估费'}])}
              {renderTaxGroup("CT投入", 'costCt', costCt, [{key: 'construction', label: '专线建设'}, {key: 'maintenance', label: '专线维护'}, {key: 'other', label: '其他产品成本'}, {key: 'bandwidth', label: '专线带宽成本'}, {key: 'renewal', label: '专线/其他产品续签成本'}])}
              {renderTaxGroup("非IT/CT投入 & 综合类成本", 'costMix', costMix, [{key: 'non_it_ct', label: '工程施工投入等'}, {key: 'marketing', label: '融合营销成本'}, {key: 'channel', label: '渠道酬金'}, {key: 'other', label: '其他管理费用等'}])}
            </div>
          )}

          {activeTab === "cashflow" && (
            <IctCashflowTable state={state} calculations={calculations} />
          )}

          <div className={`bg-card border border-border rounded-xl p-8 shadow-sm flex-col gap-6 ${activeTab === "generate" ? "flex" : "hidden"}`}>
            <h3 className="text-lg font-bold text-foreground">即将生成：{selectedTemplate}</h3>
            <TemplateForms
              selectedTemplate={selectedTemplate}
              projectData={{ basic: {proj_name: projName, customer_name: customerName, project_years: projectYears}, cost: { it: costIt, ct: costCt, mix: costMix }, revenue: { it: revIt, ct: revCt, non_it_ct: revNonItCt } }}
              metrics={metrics}
              projectBackground={state.projectBackground}
              setProjectBackground={state.setProjectBackground}
              techItems={state.techItems}
              setTechItems={state.setTechItems}
              inqVendors={state.inqVendors}
              setInqVendors={state.setInqVendors}
              outputDir={activeProject?.folder_path || undefined}
              projectId={activeProject?.id || undefined}
            />
          </div>

          <IctMetricsDashboard metrics={metrics} />
        </div>

        {(activeTab === 'revenue' || activeTab === 'cost') && (
          <div className="w-[300px] bg-card border-l border-border p-6 flex flex-col shrink-0 overflow-y-auto animate-in slide-in-from-right duration-200">
            <h3 className="font-bold text-foreground mb-4">智能反算</h3>
            <div className="bg-card border border-border p-4 rounded-xl flex flex-col gap-4 mb-6">
              <p className="text-xs leading-relaxed text-secondary-foreground">
                反算结果为含税总额参数值，系统将根据当前资金收付模型自动分摊至各年度现金流。
              </p>
              <div className="flex gap-2 bg-background p-1 border border-border rounded-lg">
                <button className={`flex-1 py-1.5 text-sm font-semibold rounded-md ${revMode === 'cost' ? 'bg-primary text-primary-foreground shadow-sm' : 'text-secondary-foreground'}`} onClick={() => setRevMode('cost')}>反算投入</button>
                <button className={`flex-1 py-1.5 text-sm font-semibold rounded-md ${revMode === 'revenue' ? 'bg-primary text-primary-foreground shadow-sm' : 'text-secondary-foreground'}`} onClick={() => setRevMode('revenue')}>反算收入</button>
              </div>
              <div className="flex gap-2 bg-background p-1 border border-border rounded-lg">
                <button className={`flex-1 py-1.5 text-sm font-semibold rounded-md ${revTargetType === 'margin' ? 'bg-primary text-primary-foreground shadow-sm' : 'text-secondary-foreground'}`} onClick={() => setRevTargetType('margin')}>目标毛利润率</button>
                <button className={`flex-1 py-1.5 text-sm font-semibold rounded-md ${revTargetType === 'npv_rate' ? 'bg-primary text-primary-foreground shadow-sm' : 'text-secondary-foreground'}`} onClick={() => setRevTargetType('npv_rate')}>目标净现值率</button>
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-xs font-semibold text-secondary-foreground">目标值 (如0.15代表15%)</label>
                <input type="number" step="0.0001" className="bg-card border border-input px-3 py-2 rounded-md outline-none text-sm" value={revTargetValue} onChange={e => setRevTargetValue(e.target.value)} />
              </div>
              <button className="flex w-full items-center justify-center gap-2 bg-primary text-primary-foreground font-bold py-2 rounded-lg shadow-sm" onClick={performReverseCalculation}><AppIcon name="reverse" size={16} /> 智能反算</button>
            </div>
            <h3 className="font-bold text-foreground mb-4">采购甄选费测算</h3>
            <div className="bg-card border border-border p-4 rounded-xl flex flex-col gap-3">
              <div className="flex flex-col gap-1">
                 <label className="text-xs font-semibold text-secondary-foreground">供应商报价 (元)</label>
                 <input type="number" value={selQuote} onChange={e => handleSelFeeChange('quote', e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md text-sm outline-none" />
              </div>
              <div className="flex flex-col gap-1">
                 <label className="text-xs font-semibold text-secondary-foreground">代理服务费浮动 (+)</label>
                 <input type="number" value={selMarkup} onChange={e => handleSelFeeChange('markup', e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md text-sm outline-none" />
              </div>
              <div className="flex flex-col gap-1">
                 <label className="text-xs font-semibold text-secondary-foreground">测算甄选费 / 实际测算成本</label>
                 <div className="flex gap-2">
                   <input type="text" disabled value={selFee} className="bg-slate-50 border border-input px-3 py-2 rounded-md text-sm w-full text-secondary-foreground" />
                   <input type="text" disabled value={selActualCost} className="bg-slate-50 border border-input px-3 py-2 rounded-md text-sm w-full text-secondary-foreground" />
                 </div>
              </div>
              <div className="flex flex-col gap-1 mt-2 border-t border-border pt-3">
                 <label className="text-xs font-semibold text-primary">甄选最高限价 (反向测算入口)</label>
                 <input type="number" value={selLimit} onChange={e => handleSelFeeChange('limit', e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md text-sm outline-none text-foreground font-bold" />
              </div>
              <button
                onClick={applySelectionLimit}
                disabled={!selLimit}
                className="mt-2 bg-primary hover:bg-primary/95 text-primary-foreground disabled:opacity-50 disabled:cursor-not-allowed font-bold py-2.5 rounded-lg shadow-sm hover:shadow-md transition-all active:scale-[0.98] w-full text-xs flex items-center justify-center gap-1.5"
              >
                <AppIcon name="download" size={14} /> 填入集成服务
              </button>
            </div>
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
                      {err.side === 'income' ? '收入侧' : '支出侧'} - {err.taxRate}% 税率组
                    </span>
                    <span className="text-xs font-mono bg-muted px-2 py-0.5 rounded">{err.key}</span>
                  </div>
                  <div className="grid grid-cols-3 gap-4 mt-2 text-sm">
                    <div className="flex flex-col gap-1">
                      <span className="text-secondary-foreground text-xs">录入不含税</span>
                      <span className="font-bold">{err.actualExcl} 元</span>
                    </div>
                    <div className="flex flex-col gap-1">
                      <span className="text-secondary-foreground text-xs">预期绝对值</span>
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
              <h4 className="font-bold text-sm text-foreground">另存为新方案</h4>
              <button
                type="button"
                onClick={() => setShowSaveAsModal(false)}
                className="text-secondary-foreground hover:bg-secondary p-1 rounded-md"
              >
                <AppIcon name="close" size={14} />
              </button>
            </div>
            <div className="p-6">
              <label className="text-xs font-semibold text-secondary-foreground block mb-1.5">方案名称 <span className="text-red-500">*</span></label>
              <input
                id="save_as_new_scheme_name_input"
                type="text"
                required
                placeholder="例如：方案 B、第二轮测算"
                value={saveAsSchemeName}
                onChange={(e) => setSaveAsSchemeName(e.target.value)}
                className="bg-card border border-input px-3 py-2 rounded-lg text-xs outline-none focus:border-ring w-full"
              />
            </div>
            <div className="border-t border-border p-3 bg-muted/10 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setShowSaveAsModal(false)}
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
    </div>
  );
}
