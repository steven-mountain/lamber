import { useState, useEffect, useCallback, useMemo } from "react"
import { invoke } from "@tauri-apps/api/core"
import WorkspaceHeader from "../components/WorkspaceHeader"
import TemplateForms from "./TemplateForms"
import { validateFinancialData, ValidationReport } from "../lib/financeValidator"
import { useAiContextStore } from "../store/useAiContextStore"
import { useRef } from "react"
import { AI_CONTEXT_KEY, buildAiContextKey } from "../utils/aiContextKeys"
import {
  buildDistributionFromModel,
  cashflowModelLabels,
  formatDistribution,
  normalizeProjectYears,
  type CashflowModel
} from "../lib/cashflowDistribution"

interface TaxItem { incl: number; tax: number; excl: number; }
const defaultTaxItem = (tax = 6): TaxItem => ({ incl: 0, tax, excl: 0 })

export default function IctLifecycle({ onBack }: { onBack: () => void }) {
  const [activeTab, setActiveTab] = useState<"basic" | "revenue" | "cost" | "cashflow" | "generate">("basic")

  // --- Basic State ---
  const [projName, setProjName] = useState("X项目")
  const [customerName, setCustomerName] = useState("X客户")
  const [propertyRights, setPropertyRights] = useState("客户")
  const [discountRate, setDiscountRate] = useState(0.055)
  const [projectYears, setProjectYears] = useState(1)
  const [cashflowModel, setCashflowModel] = useState<CashflowModel>("model_a")
  const [distRev, setDistRev] = useState<number[]>([1, 0, 0, 0, 0, 0, 0, 0, 0, 0])
  const [distCost, setDistCost] = useState<number[]>([1, 0, 0, 0, 0, 0, 0, 0, 0, 0])
  const [projectBackground, setProjectBackground] = useState("在数字经济与制造业深度融合的国家战略推动下...")
  const [techItems, setTechItems] = useState<any[]>([
    { serviceName: '集成服务', serviceDesc: '集成实施', amount: 1, unit: '项' },
    { serviceName: '维保服务', serviceDesc: '硬件维保', amount: 1, unit: '项' }
  ])
  const [inqVendors, setInqVendors] = useState<any[]>([
    { vendorName: '厂商A', amount: 0, taxRate: 6, remark: '最低' },
    { vendorName: '厂商B', amount: 0, taxRate: 6, remark: '' },
    { vendorName: '厂商C', amount: 0, taxRate: 6, remark: '' }
  ])

  // --- Quick Calc State ---
  const [quickRevTotal, setQuickRevTotal] = useState<string>("")
  const [quickRevProduct, setQuickRevProduct] = useState<string>("")
  const [quickCostTotal, setQuickCostTotal] = useState<string>("")
  const [quickCostProduct, setQuickCostProduct] = useState<string>("")

  // --- Revenue State ---
  const [revIt, setRevIt] = useState({
    integration: defaultTaxItem(6), maintenance: defaultTaxItem(6),
    device_sales: defaultTaxItem(13), device_lease: defaultTaxItem(13),
    other: defaultTaxItem(6), cloud: defaultTaxItem(6),
  })
  const [revCt, setRevCt] = useState({ line: defaultTaxItem(9), product: defaultTaxItem(6) })
  const [revNonItCt, setRevNonItCt] = useState(defaultTaxItem(9))

  // --- Cost State ---
  const [costIt, setCostIt] = useState({
    device: defaultTaxItem(13), construction: defaultTaxItem(9),
    survey: defaultTaxItem(6), integration: defaultTaxItem(6),
    other: defaultTaxItem(6), maintenance: defaultTaxItem(6),
    running: defaultTaxItem(13), bidding: defaultTaxItem(6),
    design_eval: defaultTaxItem(6), audit: defaultTaxItem(6),
  })
  const [costCt, setCostCt] = useState({
    construction: defaultTaxItem(9), maintenance: defaultTaxItem(9),
    other: defaultTaxItem(6), bandwidth: defaultTaxItem(9), renewal: defaultTaxItem(9),
  })
  const [costMix, setCostMix] = useState({
    non_it_ct: defaultTaxItem(9), marketing: defaultTaxItem(6),
    channel: defaultTaxItem(6), other: defaultTaxItem(6),
  })

  // --- Calculation Results ---
  const [cashflowTable, setCashflowTable] = useState<any[]>([])
  const [metrics, setMetrics] = useState<any>({
    npv: 0, npv_rate: 0, margin_rate: 0, dynamic_payback: "--", irr: "--",
    it_npv: 0, it_npv_rate: 0, it_margin_rate: 0
  })

  // --- Smart Reverse ---
  const [revMode, setRevMode] = useState<"cost" | "revenue">("cost")
  const [revTargetType, setRevTargetType] = useState<"margin" | "npv_rate">("margin")
  const [revTargetValue, setRevTargetValue] = useState<string>("0.15")
  const activeDistributionYears = useMemo(() => normalizeProjectYears(projectYears), [projectYears])
  const effectiveDistRev = useMemo(
    () => cashflowModel === 'model_d'
      ? buildDistributionFromModel(cashflowModel, projectYears, distRev)
      : buildDistributionFromModel(cashflowModel, projectYears),
    [cashflowModel, distRev, projectYears]
  )
  const effectiveDistCost = useMemo(
    () => cashflowModel === 'model_d'
      ? buildDistributionFromModel(cashflowModel, projectYears, distCost)
      : buildDistributionFromModel(cashflowModel, projectYears),
    [cashflowModel, distCost, projectYears]
  )

  // --- Selection Fee Calc ---
  const [selQuote, setSelQuote] = useState<string>("")
  const [selMarkup, setSelMarkup] = useState<string>("50")
  const [selActualCost, setSelActualCost] = useState<string>("")
  const [selFee, setSelFee] = useState<string>("")
  const [selLimit, setSelLimit] = useState<string>("")

  // --- Generate Docs State ---
  const [templates, setTemplates] = useState<string[]>([])
  const [selectedTemplate, setSelectedTemplate] = useState<string>("")

  const [reconciliationErrors, setReconciliationErrors] = useState<ValidationReport[]>([])
  const [showReconciliationModal, setShowReconciliationModal] = useState(false)
  const [currentTotalDifference, setCurrentTotalDifference] = useState("0")
  const [pendingTab, setPendingTab] = useState<{tab: string, template?: string} | null>(null)
  const [showConfirmIgnore, setShowConfirmIgnore] = useState(false)
  const [ignoredTailValue, setIgnoredTailValue] = useState<string | null>(null)
  const [ignoredDataHash, setIgnoredDataHash] = useState<string | null>(null)

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
  }

  // --- AI Context Sync ---
  const updateData = useAiContextStore(state => state.updateBusinessData);
  const setActiveModule = useAiContextStore(state => state.setActiveModule);
  const syncTimerRef = useRef<NodeJS.Timeout | null>(null);

  useEffect(() => {
    setActiveModule(AI_CONTEXT_KEY.ICT_CORE);
    return () => {
      if (syncTimerRef.current) clearTimeout(syncTimerRef.current);
    };
  }, [setActiveModule]);

  useEffect(() => {
    if (cashflowModel === 'model_d') {
      setDistRev(prev => buildDistributionFromModel(cashflowModel, projectYears, prev))
      setDistCost(prev => buildDistributionFromModel(cashflowModel, projectYears, prev))
      return
    }

    const nextDist = buildDistributionFromModel(cashflowModel, projectYears)
    setDistRev(nextDist)
    setDistCost(nextDist)
  }, [cashflowModel, projectYears])

  useEffect(() => {
    // Debounced sync (500ms)
    if (syncTimerRef.current) clearTimeout(syncTimerRef.current);
    
    syncTimerRef.current = setTimeout(() => {
      const payload = {
        monetary_unit: '元',
        currency: 'CNY',
        ...getInputDataPayload(),
        project_background: projectBackground, // Explicitly include text fields
        metrics: metrics,
        cashflow: cashflowTable
      };
      updateData(AI_CONTEXT_KEY.ICT_CORE, payload);
      console.log("AI Context Synced: ICT");
    }, 500);

    return () => {
      if (syncTimerRef.current) clearTimeout(syncTimerRef.current);
    };
  }, [revIt, revCt, revNonItCt, costIt, costCt, costMix, projectBackground, metrics, cashflowTable, projName, customerName, propertyRights, discountRate, projectYears, cashflowModel, distRev, distCost, updateData]);

  const updateTaxItem = (groupId: string, key: string, field: "incl" | "tax" | "excl", val: number) => {
    // 只要有任何金额或税率变动，立即重置尾差忽略状态，确保下一次切换标签时重新校验
    if (ignoredDataHash !== null) {
      setIgnoredDataHash(null);
      setIgnoredTailValue(null);
    }

    const processItem = (groupState: any, setGroupState: any, targetKey: string) => {
      const item = { ...groupState[targetKey], [field]: isNaN(val) ? 0 : val }
      if (field === 'incl' || field === 'tax') {
        item.excl = item.incl === 0 ? 0 : Number((item.incl / (1 + item.tax / 100)).toFixed(2))
      } else if (field === 'excl') {
        item.incl = item.excl === 0 ? 0 : Number((item.excl * (1 + item.tax / 100)).toFixed(2))
      }
      setGroupState({ ...groupState, [targetKey]: item })
    }

    if (groupId === 'revIt') processItem(revIt, setRevIt, key)
    else if (groupId === 'revCt') {
      processItem(revCt, setRevCt, key)
      if (key === 'product') processItem(costCt, setCostCt, 'other')
      if (key === 'line') processItem(costCt, setCostCt, 'bandwidth')
    }
    else if (groupId === 'revNonItCt') {
      const item = { ...revNonItCt, [field]: isNaN(val) ? 0 : val }
      if (field === 'incl' || field === 'tax') {
        item.excl = item.incl === 0 ? 0 : Number((item.incl / (1 + item.tax / 100)).toFixed(2))
      } else if (field === 'excl') {
        item.incl = item.excl === 0 ? 0 : Number((item.excl * (1 + item.tax / 100)).toFixed(2))
      }
      setRevNonItCt(item)
    }
    else if (groupId === 'costIt') processItem(costIt, setCostIt, key)
    else if (groupId === 'costCt') processItem(costCt, setCostCt, key)
    else if (groupId === 'costMix') processItem(costMix, setCostMix, key)
  }

  useEffect(() => { performCalculation() }, [revIt, revCt, revNonItCt, costIt, costCt, costMix, projectYears, discountRate, cashflowModel, distRev, distCost])
  
  const loadTemplates = useCallback(async () => {
    try {
      const list: any = await invoke('get_available_templates', { moduleId: 'ict_lifecycle' })
      setTemplates(list)
    } catch (e) {
      console.error("加载 ICT 模板失败:", e)
    }
  }, [])

  useEffect(() => {
    loadTemplates()
  }, [loadTemplates])

  const getInputDataPayload = () => ({
    project_name: projName,
    customer_name: customerName,
    property_rights: propertyRights,
    discount_rate: String(discountRate),
    project_years: projectYears,
    cashflow_model: cashflowModel,
    rev_distribution: effectiveDistRev,
    cost_distribution: effectiveDistCost,
    ignore_tail_difference: ignoredTailValue !== null,
    tail_difference_value: ignoredTailValue || "0",
    rev_it_integration: { incl_tax: String(revIt.integration.incl), tax_rate: String(revIt.integration.tax) },
    rev_it_maintenance: { incl_tax: String(revIt.maintenance.incl), tax_rate: String(revIt.maintenance.tax) },
    rev_it_device_sales: { incl_tax: String(revIt.device_sales.incl), tax_rate: String(revIt.device_sales.tax) },
    rev_it_device_lease: { incl_tax: String(revIt.device_lease.incl), tax_rate: String(revIt.device_lease.tax) },
    rev_it_other: { incl_tax: String(revIt.other.incl), tax_rate: String(revIt.other.tax) },
    rev_it_cloud: { incl_tax: String(revIt.cloud.incl), tax_rate: String(revIt.cloud.tax) },
    rev_ct_line: { incl_tax: String(revCt.line.incl), tax_rate: String(revCt.line.tax) },
    rev_ct_product: { incl_tax: String(revCt.product.incl), tax_rate: String(revCt.product.tax) },
    rev_non_it_ct: { incl_tax: String(revNonItCt.incl), tax_rate: String(revNonItCt.tax) },
    cost_it_device: { incl_tax: String(costIt.device.incl), tax_rate: String(costIt.device.tax) },
    cost_it_construction: { incl_tax: String(costIt.construction.incl), tax_rate: String(costIt.construction.tax) },
    cost_it_survey: { incl_tax: String(costIt.survey.incl), tax_rate: String(costIt.survey.tax) },
    cost_it_integration: { incl_tax: String(costIt.integration.incl), tax_rate: String(costIt.integration.tax) },
    cost_it_other: { incl_tax: String(costIt.other.incl), tax_rate: String(costIt.other.tax) },
    cost_it_maintenance: { incl_tax: String(costIt.maintenance.incl), tax_rate: String(costIt.maintenance.tax) },
    cost_it_running: { incl_tax: String(costIt.running.incl), tax_rate: String(costIt.running.tax) },
    cost_it_bidding: { incl_tax: String(costIt.bidding.incl), tax_rate: String(costIt.bidding.tax) },
    cost_it_design_eval: { incl_tax: String(costIt.design_eval.incl), tax_rate: String(costIt.design_eval.tax) },
    cost_it_audit: { incl_tax: String(costIt.audit.incl), tax_rate: String(costIt.audit.tax) },
    cost_ct_construction: { incl_tax: String(costCt.construction.incl), tax_rate: String(costCt.construction.tax) },
    cost_ct_maintenance: { incl_tax: String(costCt.maintenance.incl), tax_rate: String(costCt.maintenance.tax) },
    cost_ct_other: { incl_tax: String(costCt.other.incl), tax_rate: String(costCt.other.tax) },
    cost_ct_bandwidth: { incl_tax: String(costCt.bandwidth.incl), tax_rate: String(costCt.bandwidth.tax) },
    cost_ct_renewal: { incl_tax: String(costCt.renewal.incl), tax_rate: String(costCt.renewal.tax) },
    cost_non_it_ct: { incl_tax: String(costMix.non_it_ct.incl), tax_rate: String(costMix.non_it_ct.tax) },
    cost_mix_marketing: { incl_tax: String(costMix.marketing.incl), tax_rate: String(costMix.marketing.tax) },
    cost_mix_channel: { incl_tax: String(costMix.channel.incl), tax_rate: String(costMix.channel.tax) },
    cost_mix_other: { incl_tax: String(costMix.other.incl), tax_rate: String(costMix.other.tax) },
  })

  const performCalculation = async () => {
    try {
      const res: any = await invoke('calculate_ict_benefit', { input: getInputDataPayload() })
      if (res) {
        setCashflowTable(res.cashflow)
        setMetrics(res)
      }
    } catch (e) {
      console.error(e)
    }
  }

  const handleSelFeeChange = async (type: 'quote' | 'markup' | 'limit', val: string) => {
    if (type === 'quote') setSelQuote(val)
    if (type === 'markup') setSelMarkup(val)
    if (type === 'limit') setSelLimit(val)

    const currentQuote = type === 'quote' ? val : selQuote
    const currentMarkup = type === 'markup' ? val : selMarkup
    const currentLimit = type === 'limit' ? val : selLimit

    try {
      if (type === 'quote' || type === 'markup') {
        const res: any = await invoke('calculate_selection_fee', { quote: currentQuote || "0", markup: currentMarkup || "0" })
        setSelLimit(res.final_limit)
        setSelActualCost(res.actual_cost)
        setSelFee(res.selection_fee)
      } else if (type === 'limit') {
        const res: any = await invoke('reverse_calculate_selection_fee', { limit: currentLimit || "0", markup: currentMarkup || "0" })
        setSelQuote(res.quote)
        setSelActualCost(res.actual_cost)
        setSelFee(res.selection_fee)
      }
    } catch(e) {
      console.error("甄选限价计算失败:", e)
    }
  }

  const applySelectionLimit = () => {
    if (selLimit) {
      updateTaxItem('costIt', 'integration', 'incl', Number(selLimit))
      if (selFee) {
        updateTaxItem('costIt', 'bidding', 'incl', Number(selFee))
      }
    }
  }

  const performReverseCalculation = async () => {
    if (!revTargetValue) return alert("请输入目标值！")
    try {
      const apiName = revMode === 'revenue' ? 'reverse_calc_ict_revenue_target' : 'reverse_calc_ict_target'
      const basePayload = getInputDataPayload()
      const valStr: string = await invoke(apiName, {
        input: basePayload,
        targetType: revTargetType,
        targetValue: String(revTargetValue)
      })
      
      const numVal = Number(valStr)
      const nextPayload = {
        ...basePayload,
        ...(revMode === 'revenue'
          ? { rev_it_integration: { ...basePayload.rev_it_integration, incl_tax: String(numVal) } }
          : { cost_it_integration: { ...basePayload.cost_it_integration, incl_tax: String(numVal) } })
      }

      if (revMode === 'revenue') {
        updateTaxItem('revIt', 'integration', 'incl', numVal)
        setActiveTab("revenue")
      } else {
        updateTaxItem('costIt', 'integration', 'incl', numVal)
        handleSelFeeChange('limit', String(numVal))
        setActiveTab("cost")
      }

      const refreshed: any = await invoke('calculate_ict_benefit', { input: nextPayload })
      if (refreshed) {
        setCashflowTable(refreshed.cashflow)
        setMetrics(refreshed)
      }

      const distText = revMode === 'revenue'
        ? formatDistribution(effectiveDistRev)
        : formatDistribution(effectiveDistCost)
      const targetName = revTargetType === 'margin' ? '毛利润率' : '净现值率'
      const reverseFieldName = revMode === 'revenue' ? '系统集成服务收入' : '系统集成服务成本'

      alert(
        `反算完成：${formatCurrency(numVal)}\n` +
        `目标：${targetName} ≥ ${formatPercent(Number(revTargetValue))}\n` +
        `反算字段：${reverseFieldName}\n` +
        `该结果为含税总额参数值，将按当前资金收付模型自动分摊。\n` +
        `当前资金收付模型：${cashflowModelLabels[cashflowModel]}\n` +
        `年度分布：${distText}\n` +
        `已自动刷新 10 年现金流推演。`
      )
    } catch (e) {
      alert("反推失败: " + e)
    }
  }

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
                <input type="number" placeholder="含税" className="w-full bg-muted border border-border px-3 py-2 rounded-md outline-none text-sm" value={groupState[item.key].incl === 0 ? "" : groupState[item.key].incl} onChange={e => updateTaxItem(groupId, item.key, 'incl', Number(e.target.value))} />
                <input type="number" placeholder="税率" className="w-20 bg-muted border border-border px-3 py-2 rounded-md outline-none text-sm" value={groupState[item.key].tax} onChange={e => updateTaxItem(groupId, item.key, 'tax', Number(e.target.value))} />
                <input type="number" placeholder="不含税" className={`w-full bg-background border px-3 py-2 rounded-md outline-none text-sm focus:border-primary ${itemErr ? 'border-red-500 ring-1 ring-red-500' : 'border-border'}`} value={groupState[item.key].excl === 0 ? "" : groupState[item.key].excl} onChange={e => updateTaxItem(groupId, item.key, 'excl', Number(e.target.value))} />
              </div>
              {itemErr && <span className="text-[10px] text-red-500 font-bold">校验失败：偏离 {itemErr.difference} 元，要求：{itemErr.expectedExcl} 元</span>}
            </div>
          )
        })}
      </div>
    </div>
  )

  const formatCurrency = (v: number) => new Intl.NumberFormat('zh-CN', { style: 'currency', currency: 'CNY' }).format(v)
  const formatPercent = (v: number) => (v * 100).toFixed(2) + "%"
  const distributionPreview = (
    <div className="mt-6 mb-4 border-t border-border pt-4 text-sm">
      <div className="font-bold text-foreground mb-2">资金分布预览</div>
      <div className="grid grid-cols-1 gap-1 text-secondary-foreground">
        <div>当前资金收付模型：{cashflowModelLabels[cashflowModel]}</div>
        <div>收入分布：{formatDistribution(effectiveDistRev)}</div>
        <div>成本分布：{formatDistribution(effectiveDistCost)}</div>
      </div>
    </div>
  )
  const formatYearRange = (start: number, end: number) => start + 1 === end
    ? `第 ${start + 1} 年`
    : `第 ${start + 1}-${end} 年`
  const distributionSegments = activeDistributionYears <= 5
    ? [{ label: formatYearRange(0, activeDistributionYears), start: 0, end: activeDistributionYears }]
    : [
      { label: formatYearRange(0, 5), start: 0, end: 5 },
      { label: formatYearRange(5, activeDistributionYears), start: 5, end: activeDistributionYears },
    ]

  return (
    <div className="flex flex-col flex-1 animate-in fade-in duration-300 h-full overflow-hidden">
      <WorkspaceHeader moduleId="ict_lifecycle" title="ICT项目全生命周期" onBack={onBack} onPathChange={() => loadTemplates()} />
      
      <div className="flex flex-1 overflow-hidden">
        <div className="w-[260px] bg-muted p-6 overflow-y-auto flex flex-col gap-4 border-r border-border shrink-0">
          <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70 mb-1">测算流程</h3>
          <div className="flex flex-col gap-1">
            <button className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${activeTab === 'basic' ? 'bg-primary/20 text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`} onClick={() => handleTabSwitch("basic")}><span>📋</span> 项目概况与参数</button>
            <button className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${activeTab === 'revenue' ? 'bg-primary/20 text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`} onClick={() => handleTabSwitch("revenue")}><span>💰</span> 收入侧测算</button>
            <button className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${activeTab === 'cost' ? 'bg-primary/20 text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`} onClick={() => handleTabSwitch("cost")}><span>💸</span> 支出侧测算</button>
            <button className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${activeTab === 'cashflow' ? 'bg-primary/20 text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`} onClick={() => handleTabSwitch("cashflow")}><span>📈</span> 10年现金流推演</button>
          </div>
          <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70 mt-6 pt-4 border-t border-border mb-2">一键生成全流程文档</h3>
          <div className="flex flex-col gap-2">
            {templates.length === 0 ? <span className="text-xs text-secondary-foreground px-4">未找到模板文件</span> : 
              templates.map(t => {
                const isActive = selectedTemplate === t && activeTab === 'generate';
                return (
                  <button 
                    key={t}
                    className={`relative overflow-hidden px-4 py-3 rounded-lg text-sm flex items-start gap-2.5 transition-all text-left border-b border-border/60 shadow-sm ${isActive ? 'bg-primary/20 text-primary font-bold border-transparent shadow' : 'text-secondary-foreground font-semibold hover:bg-primary/10 hover:text-primary'}`}
                    onClick={() => handleTabSwitch("generate", t)}
                  >
                    {isActive && <div className="absolute left-0 top-0 bottom-0 w-1 bg-primary" />}
                    <span className="mt-0.5 shrink-0 text-base">{t.endsWith('.xlsx') ? '📊' : '📄'}</span>
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
          {activeTab === "basic" && (
            <div className="bg-card border border-border rounded-xl p-8 shadow-sm">
              <h3 className="text-lg font-bold text-foreground mb-6">项目概况</h3>
              <div className="grid grid-cols-2 gap-6">
                <div className="flex flex-col gap-2"><label className="text-sm font-bold text-secondary-foreground">项目名称</label><input id="ict-proj-name" type="text" value={projName} onChange={e => setProjName(e.target.value)} className="bg-muted border border-border px-3.5 py-2.5 rounded-md outline-none focus:border-primary" /></div>
                <div className="flex flex-col gap-2"><label className="text-sm font-bold text-secondary-foreground">客户单位名称</label><input id="ict-customer-name" type="text" value={customerName} onChange={e => setCustomerName(e.target.value)} className="bg-muted border border-border px-3.5 py-2.5 rounded-md outline-none focus:border-primary" /></div>
                <div className="flex flex-col gap-2"><label className="text-sm font-bold text-secondary-foreground">产权归属</label><input id="ict-property-rights" type="text" value={propertyRights} onChange={e => setPropertyRights(e.target.value)} className="bg-muted border border-border px-3.5 py-2.5 rounded-md outline-none focus:border-primary" /></div>
                <div className="flex flex-col gap-2"><label className="text-sm font-bold text-secondary-foreground">项目建设/服务周期 (年)</label><input id="ict-project-years" type="number" min={1} max={10} value={projectYears} onChange={e => setProjectYears(Number(e.target.value))} className="bg-muted border border-border px-3.5 py-2.5 rounded-md outline-none focus:border-primary" /></div>
                <div className="flex flex-col gap-2"><label className="text-sm font-bold text-secondary-foreground">折现率</label><input id="ict-discount-rate" type="number" step={0.001} value={discountRate} onChange={e => setDiscountRate(Number(e.target.value))} className="bg-muted border border-border px-3.5 py-2.5 rounded-md outline-none focus:border-primary" /></div>
                <div className="flex flex-col gap-2">
                  <label className="text-sm font-bold text-secondary-foreground">资金收付模型</label>
                  <select id="ict-cashflow-model" value={cashflowModel} onChange={e => setCashflowModel(e.target.value as CashflowModel)} className="bg-muted border border-border px-3.5 py-2.5 rounded-md outline-none focus:border-primary">
                    <option value="model_a">模型 A: 100% 在第一年收付</option>
                    <option value="model_b">模型 B: 按周期等额收付 (每年 1/n)</option>
                    <option value="model_c">模型 C: 尾款质保金 (首年95%，末年5%)</option>
                    <option value="model_d">模型 D: 高级自定义分配</option>
                  </select>
                </div>
                <div className="flex flex-col gap-2 col-span-2">
                  <label className="text-sm font-bold text-secondary-foreground">项目背景</label>
                  <textarea id="ict-project-bg" rows={3} value={projectBackground} onChange={e => setProjectBackground(e.target.value)} className="bg-muted border border-border px-3.5 py-2.5 rounded-md outline-none focus:border-primary" />
                </div>
              </div>
              {cashflowModel === 'model_d' && (
                <div className="mt-6 pt-6 border-t border-border">
                  <h4 className="text-sm font-bold text-secondary-foreground mb-4">高级自定义分配 (可输入任意非负比例，系统自动归一化)</h4>
                  <p className="text-xs text-secondary-foreground mb-4">
                    仅展示项目周期内的 {activeDistributionYears} 个年份；周期外年份按 0 处理，不参与现金流分摊。
                  </p>
                  <div className="grid grid-cols-1 xl:grid-cols-2 gap-5">
                    {distributionSegments.map(segment => {
                      const segmentYears = Array.from({ length: segment.end - segment.start }, (_, idx) => segment.start + idx)

                      return (
                        <div key={segment.label} className="min-w-0 rounded-lg border border-border bg-muted/30 p-3">
                          <div className="text-xs font-bold text-secondary-foreground mb-3">{segment.label}</div>
                          <div
                            className="grid gap-2 items-center text-center text-sm"
                            style={{ gridTemplateColumns: `72px repeat(${segmentYears.length}, minmax(0, 1fr))` }}
                          >
                            <div className="font-bold text-right pr-2 text-secondary-foreground">年份</div>
                            {segmentYears.map(i => (
                              <div key={`year-${i}`} className="font-bold">{i + 1}</div>
                            ))}
                            <div className="font-bold text-right pr-2 text-secondary-foreground">收入比例</div>
                            {segmentYears.map(i => (
                              <input key={`rev-${i}`} type="number" step="0.01" value={distRev[i]} onChange={e => { const newArr = [...distRev]; newArr[i] = Number(e.target.value); setDistRev(newArr); }} className="min-w-0 w-full bg-muted border border-border rounded px-1 py-1 outline-none focus:border-primary text-center" />
                            ))}
                            <div className="font-bold text-right pr-2 text-secondary-foreground">支出比例</div>
                            {segmentYears.map(i => (
                              <input key={`cost-${i}`} type="number" step="0.01" value={distCost[i]} onChange={e => { const newArr = [...distCost]; newArr[i] = Number(e.target.value); setDistCost(newArr); }} className="min-w-0 w-full bg-muted border border-border rounded px-1 py-1 outline-none focus:border-primary text-center" />
                            ))}
                          </div>
                        </div>
                      )
                    })}
                  </div>
                </div>
              )}
              {distributionPreview}
            </div>
          )}

          {activeTab === "revenue" && (
            <div>
              <div className="mb-6 border border-primary/30 bg-primary/5 p-5 rounded-xl shadow-sm">
                <div className="flex items-center justify-between mb-2">
                  <h3 className="font-bold text-primary text-base flex items-center gap-2">
                    <span>⚡</span> 快捷收入拆分计算器 (融入自有产品测算)
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
                      className="bg-background border border-border px-3 py-2 rounded-md outline-none focus:border-primary text-sm font-semibold shadow-sm" 
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
                      className="bg-background border border-border px-3 py-2 rounded-md outline-none focus:border-primary text-sm font-semibold shadow-sm" 
                    />
                  </div>
                  <div className="flex flex-col gap-1.5 flex-1">
                    <label className="text-xs font-bold text-primary">系统集成服务收入 (自动扣减)</label>
                    <input 
                      type="number" 
                      disabled 
                      value={Math.max(0, (Number(quickRevTotal)||0) - (Number(quickRevProduct)||0)).toFixed(2)} 
                      className="bg-background/50 border border-primary/30 px-3 py-2 rounded-md outline-none text-sm font-bold text-primary shadow-sm" 
                    />
                  </div>
                  <button 
                    onClick={() => {
                      const integration = Math.max(0, (Number(quickRevTotal)||0) - (Number(quickRevProduct)||0));
                      updateTaxItem('revIt', 'integration', 'incl', Number(integration.toFixed(2)));
                      if (quickRevProduct) updateTaxItem('revCt', 'product', 'incl', Number(Number(quickRevProduct).toFixed(2)));
                    }} 
                    className="bg-primary text-primary-foreground font-bold px-5 py-2 rounded-md text-sm hover:bg-primary/90 transition-all shadow-sm hover:shadow active:scale-[0.98]"
                  >
                    ⬇️ 一键填入表单
                  </button>
                </div>
              </div>
              <div className="mb-4 text-xs text-blue-600 bg-blue-50 p-3 rounded-lg border border-blue-200">
                💡 提示：在「CT收入」中填写的产品或专线含税收入，将会自动【1:1平过】填入对应的「CT投入」中。
              </div>
              {renderTaxGroup("IT/移动云收入", 'revIt', revIt, [{key: 'integration', label: '系统集成服务收入'}, {key: 'maintenance', label: '维保收入'}, {key: 'device_sales', label: '设备销售收入'}, {key: 'device_lease', label: '设备租赁收入'}, {key: 'other', label: '其他收入'}, {key: 'cloud', label: '移动云-定制化收入'}])}
              {renderTaxGroup("CT收入", 'revCt', revCt, [{key: 'line', label: '专线收入'}, {key: 'product', label: '产品收入'}])}
              {renderTaxGroup("非IT/CT收入", 'revNonItCt', { item: revNonItCt }, [{key: 'item', label: '工程施工收入等'}])}
            </div>
          )}

          {activeTab === "cost" && (
            <div>
              <div className="mb-6 border border-primary/30 bg-primary/5 p-4 rounded-xl">
                <h3 className="font-bold text-primary mb-2">快捷投入拆分计算器</h3>
                <p className="text-xs text-secondary-foreground mb-4">输入含税总投入与含税产品投入，计算“系统集成服务投入”并一键填入下方表单。</p>
                <div className="flex gap-4 items-end">
                  <div className="flex flex-col gap-1 flex-1">
                    <label className="text-xs font-semibold">含税总投入</label>
                    <input type="number" value={quickCostTotal} onChange={e => setQuickCostTotal(e.target.value)} className="bg-background border border-border px-3 py-2 rounded-md outline-none text-sm" />
                  </div>
                  <div className="flex flex-col gap-1 flex-1">
                    <label className="text-xs font-semibold">含税产品投入</label>
                    <input type="number" value={quickCostProduct} onChange={e => setQuickCostProduct(e.target.value)} className="bg-background border border-border px-3 py-2 rounded-md outline-none text-sm" />
                  </div>
                  <div className="flex flex-col gap-1 flex-1">
                    <label className="text-xs font-semibold text-primary">系统集成服务投入 (自动)</label>
                    <input type="number" disabled value={Math.max(0, (Number(quickCostTotal)||0) - (Number(quickCostProduct)||0))} className="bg-background/50 border border-primary/30 px-3 py-2 rounded-md outline-none text-sm font-bold text-primary" />
                  </div>
                  <button onClick={() => {
                     const integration = Math.max(0, (Number(quickCostTotal)||0) - (Number(quickCostProduct)||0));
                     updateTaxItem('costIt', 'integration', 'incl', integration);
                     if (quickCostProduct) updateTaxItem('costIt', 'device', 'incl', Number(quickCostProduct));
                  }} className="bg-primary text-primary-foreground font-bold px-4 py-2 rounded-md text-sm hover:bg-primary/90 transition-colors">⬇️ 一键填入</button>
                </div>
              </div>
              {renderTaxGroup("IT/移动云投入", 'costIt', costIt, [{key: 'device', label: '主要设备/甲供材料'}, {key: 'construction', label: '施工'}, {key: 'survey', label: '勘察设计/预备费'}, {key: 'integration', label: '集成服务'}, {key: 'other', label: '其他投入'}, {key: 'maintenance', label: '维护费用'}, {key: 'running', label: '其他运行支出（电费等）'}, {key: 'bidding', label: '中标服务费'}, {key: 'design_eval', label: '设计院成本评估费'}, {key: 'audit', label: '第三方审计评估费'}])}
              {renderTaxGroup("CT投入", 'costCt', costCt, [{key: 'construction', label: '专线建设'}, {key: 'maintenance', label: '专线维护'}, {key: 'other', label: '其他产品成本'}, {key: 'bandwidth', label: '专线带宽成本'}, {key: 'renewal', label: '专线/其他产品续签成本'}])}
              {renderTaxGroup("非IT/CT投入 & 综合类成本", 'costMix', costMix, [{key: 'non_it_ct', label: '工程施工投入等'}, {key: 'marketing', label: '融合营销成本'}, {key: 'channel', label: '渠道酬金'}, {key: 'other', label: '其他管理费用等'}])}
            </div>
          )}

          {activeTab === "cashflow" && (
            <div className="bg-card border border-border rounded-xl p-8 shadow-sm">
              <h3 className="text-lg font-bold text-foreground mb-6">1-10年项目现金流推演</h3>
              {distributionPreview}
              <div className="overflow-x-auto">
                <table className="w-full text-sm text-right border-separate border-spacing-y-2">
                  <thead>
                    <tr className="text-secondary-foreground font-bold text-xs uppercase">
                      <th className="text-left pb-2">年份</th><th className="pb-2">现金流入</th><th className="pb-2">现金流出</th><th className="pb-2">净流量</th><th className="pb-2">累计净流量</th><th className="pb-2">净流量现值</th><th className="pb-2">累计现值</th>
                    </tr>
                  </thead>
                  <tbody>
                    {cashflowTable.map((row: any, i: number) => (
                      <tr key={i} className="bg-muted">
                        <td className="text-left p-3 rounded-l-md font-semibold">第{row.year}年</td>
                        <td className="p-3">{formatCurrency(row.cash_in)}</td><td className="p-3">{formatCurrency(row.cash_out)}</td>
                        <td className="p-3 font-bold text-primary">{formatCurrency(row.net_cash)}</td><td className="p-3">{formatCurrency(row.cum_net_cash)}</td>
                        <td className="p-3">{formatCurrency(row.pv)}</td><td className="p-3 rounded-r-md">{formatCurrency(row.cum_pv)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          <div className={`bg-card border border-border rounded-xl p-8 shadow-sm flex-col gap-6 ${activeTab === "generate" ? "flex" : "hidden"}`}>
            <h3 className="text-lg font-bold text-foreground">即将生成：{selectedTemplate}</h3>
            <TemplateForms 
              selectedTemplate={selectedTemplate} 
              projectData={{ basic: {proj_name: projName, customer_name: customerName, project_years: projectYears}, cost: { it: costIt, ct: costCt, mix: costMix }, revenue: { it: revIt, ct: revCt, non_it_ct: revNonItCt } }} 
              metrics={metrics}
              projectBackground={projectBackground} 
              setProjectBackground={setProjectBackground} 
              techItems={techItems}
              setTechItems={setTechItems}
              inqVendors={inqVendors}
              setInqVendors={setInqVendors}
            />
          </div>

          <div className="mt-auto pt-6 border-t border-border mt-8">
            <h3 className="text-sm font-bold text-secondary-foreground mb-4">实时效益评估结果</h3>
            <div className="grid grid-cols-5 gap-4">
              <div className="bg-muted p-4 rounded-lg flex flex-col gap-1 border border-border">
                <span className="text-xs font-semibold text-secondary-foreground">项目净现值 (NPV)</span>
                <span id="ict-metric-npv" className="text-lg font-bold">{formatCurrency(metrics.npv)}</span>
              </div>
              <div className="bg-muted p-4 rounded-lg flex flex-col gap-1 border border-border">
                <span className="text-xs font-semibold text-secondary-foreground">净现值率</span>
                <span id="ict-metric-npv-rate" className="text-lg font-bold text-green-600">{formatPercent(metrics.npv_rate)}</span>
              </div>
              <div className="bg-muted p-4 rounded-lg flex flex-col gap-1 border border-border">
                <span className="text-xs font-semibold text-secondary-foreground">毛利润率</span>
                <span id="ict-metric-margin" className="text-lg font-bold text-green-600">{formatPercent(metrics.margin_rate)}</span>
              </div>
              <div className="bg-muted p-4 rounded-lg flex flex-col gap-1 border border-border">
                <span className="text-xs font-semibold text-secondary-foreground">动态回收期 (年)</span>
                <span id="ict-metric-payback" className="text-lg font-bold">{metrics.dynamic_payback}</span>
              </div>
              <div className="bg-muted p-4 rounded-lg flex flex-col gap-1 border border-border">
                <span className="text-xs font-semibold text-secondary-foreground">内部收益率 (IRR)</span>
                <span id="ict-metric-irr" className="text-lg font-bold">{metrics.irr}</span>
              </div>
            </div>
          </div>
        </div>
        
        {(activeTab === 'revenue' || activeTab === 'cost') && (
          <div className="w-[300px] bg-card border-l border-border p-6 flex flex-col shrink-0 overflow-y-auto animate-in slide-in-from-right duration-200">
            <h3 className="font-bold text-foreground mb-4">智能反算</h3>
            <div className="bg-muted border border-border p-4 rounded-xl flex flex-col gap-4 mb-6">
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
                <input type="number" step="0.0001" className="bg-background border border-border px-3 py-2 rounded-md outline-none text-sm" value={revTargetValue} onChange={e => setRevTargetValue(e.target.value)} />
              </div>
              <button className="bg-primary text-primary-foreground font-bold py-2 rounded-lg shadow-sm w-full" onClick={performReverseCalculation}>⚡ 智能反算</button>
            </div>
            <h3 className="font-bold text-foreground mb-4">采购甄选费测算</h3>
            <div className="bg-muted border border-border p-4 rounded-xl flex flex-col gap-3">
              <div className="flex flex-col gap-1">
                 <label className="text-xs font-semibold text-secondary-foreground">供应商报价 (元)</label>
                 <input type="number" value={selQuote} onChange={e => handleSelFeeChange('quote', e.target.value)} className="bg-background border border-border px-3 py-2 rounded-md text-sm outline-none" />
              </div>
              <div className="flex flex-col gap-1">
                 <label className="text-xs font-semibold text-secondary-foreground">代理服务费浮动 (+)</label>
                 <input type="number" value={selMarkup} onChange={e => handleSelFeeChange('markup', e.target.value)} className="bg-background border border-border px-3 py-2 rounded-md text-sm outline-none" />
              </div>
              <div className="flex flex-col gap-1">
                 <label className="text-xs font-semibold text-secondary-foreground">测算甄选费 / 实际测算成本</label>
                 <div className="flex gap-2">
                   <input type="text" disabled value={selFee} className="bg-background/50 border border-border px-3 py-2 rounded-md text-sm w-full text-secondary-foreground" />
                   <input type="text" disabled value={selActualCost} className="bg-background/50 border border-border px-3 py-2 rounded-md text-sm w-full text-secondary-foreground" />
                 </div>
              </div>
              <div className="flex flex-col gap-1 mt-2 border-t border-border pt-3">
                 <label className="text-xs font-semibold text-primary">甄选最高限价 (反向测算入口)</label>
                 <input type="number" value={selLimit} onChange={e => handleSelFeeChange('limit', e.target.value)} className="bg-primary/5 border border-primary px-3 py-2 rounded-md text-sm outline-none text-primary font-bold" />
              </div>
              <button 
                onClick={applySelectionLimit} 
                disabled={!selLimit}
                className="mt-2 bg-primary hover:bg-primary/95 text-primary-foreground disabled:opacity-50 disabled:cursor-not-allowed font-bold py-2.5 rounded-lg shadow-sm hover:shadow-md transition-all active:scale-[0.98] w-full text-xs flex items-center justify-center gap-1.5"
              >
                <span>⬇️</span> 填入集成服务
              </button>
            </div>
          </div>
        )}
      </div>

      {showReconciliationModal && (
        <div className="fixed inset-0 z-50 bg-background/80 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in">
          <div className="bg-card border border-red-500/30 rounded-xl shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col">
            <div className="bg-red-500/10 border-b border-red-500/20 px-6 py-4 flex items-center gap-3">
              <span className="text-red-600 text-2xl">⚠️</span>
              <div>
                <h2 className="font-bold text-red-600 text-lg">0 容差财务核算拦截</h2>
                <p className="text-xs text-red-600/80 mt-0.5">检测到税前/税后金额转换存在微小尾差，系统已拦截保存操作。</p>
              </div>
            </div>
            
            <div className="p-6 overflow-y-auto max-h-[60vh] flex flex-col gap-4 bg-muted/30">
              {reconciliationErrors.map((err, i) => (
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
          <div className="bg-card border border-border rounded-xl shadow-2xl w-full max-w-md overflow-hidden flex flex-col">
            <div className="px-6 py-4 border-b border-border bg-yellow-500/10 flex items-center gap-2">
              <span className="text-yellow-600 font-bold text-lg">⚠️</span>
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
    </div>
  )
}
