import { useEffect, useState } from "react"
import { invoke } from "@tauri-apps/api/core"
import WorkspaceHeader from "../components/WorkspaceHeader"
import { useAiContextStore } from "../store/useAiContextStore"
import { AI_CONTEXT_KEY } from "../utils/aiContextKeys"

export default function BenefitTool({ onBack }: { onBack: () => void }) {
  const [subView, setSubView] = useState<"single" | "batch">("single")
  const [input, setInput] = useState({
    total_income_incl: 1000000,
    tax_rate_it: 0.06,
    tax_rate_ct: 0.06,
    ct_income_incl_opt: "",
    calc_mode: "margin",
    target_value: 0.15
  })
  
  const [result, setResult] = useState<any>(null)
  const [warning, setWarning] = useState("")
  const setActiveModule = useAiContextStore(state => state.setActiveModule)
  const updateBusinessData = useAiContextStore(state => state.updateBusinessData)

  useEffect(() => {
    setActiveModule(AI_CONTEXT_KEY.BENEFIT_CORE)
  }, [setActiveModule])

  useEffect(() => {
    updateBusinessData(AI_CONTEXT_KEY.BENEFIT_CORE, {
      monetary_unit: '元',
      currency: 'CNY',
      view: subView,
      input,
      result,
      warning,
    })
  }, [input, result, subView, updateBusinessData, warning])

  const handleCalc = async () => {
    try {
      setWarning("")
      const payload = {
        tax_rate_it: String(input.tax_rate_it),
        tax_rate_ct: String(input.tax_rate_ct),
        total_income_incl: String(input.total_income_incl),
        calc_mode: input.calc_mode,
        target_value: String(input.target_value),
        ct_income_incl_opt: input.ct_income_incl_opt === "" ? null : String(input.ct_income_incl_opt)
      }
      const res: any = await invoke('calculate_benefit', { input: payload })
      setResult(res)
      if (res.warning_message) {
        setWarning(res.warning_message)
      }
    } catch (e: any) {
      alert('发生计算错误: ' + e)
    }
  }

  const handleBatch = async () => {
    try {
      const selected = await invoke('plugin:dialog|open', {
        multiple: false,
        filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }]
      }) as string
      if (!selected) return
      
      await invoke('process_excel_batch', { 
        moduleId: 'benefit_tool',
        filePath: selected 
      })
      
      if (confirm("✅ 批量处理完成！文件已保存至工作空间 output 目录。\n是否立即打开输出目录？")) {
        const modulePath: string = await invoke('get_module_path', { moduleId: 'benefit_tool' })
        invoke('open_file', { path: `${modulePath}/output` })
      }
    } catch (e: any) {
      alert("批量处理失败: " + e)
    }
  }

  const formatCurrency = (val: number) => new Intl.NumberFormat('zh-CN', { style: 'currency', currency: 'CNY' }).format(val || 0)
  const formatPercent = (val: number) => ((val || 0) * 100).toFixed(2) + '%'

  return (
    <div className="flex flex-col flex-1 animate-in fade-in duration-300 h-full">
      <WorkspaceHeader moduleId="benefit_tool" title="项目效益分析" onBack={onBack} />
      <div className="flex flex-1 overflow-hidden">
        <div className="w-[340px] bg-muted p-6 overflow-y-auto flex flex-col gap-4 border-r border-border shrink-0">
          <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70 mb-1">效益分析导航</h3>
          <div className="flex flex-col gap-1 mb-6">
            <button 
              className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${subView === 'single' ? 'bg-primary/20 text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`}
              onClick={() => setSubView("single")}
            >
              <span>📊</span> 单项效益测算
            </button>
            <button 
              className={`px-4 py-3 rounded-lg font-semibold text-sm flex items-center gap-2.5 transition-colors ${subView === 'batch' ? 'bg-primary/20 text-primary' : 'text-secondary-foreground hover:bg-secondary hover:text-primary'}`}
              onClick={() => setSubView("batch")}
            >
              <span>📁</span> 批量效益处理
            </button>
          </div>

          {subView === "single" ? (
            <div className="flex flex-col gap-4">
              <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70 mt-2">计算配置</h3>
              <div className="flex flex-col gap-2">
                <label className="text-xs font-bold text-secondary-foreground">含税总收入 (元)</label>
                <input type="number" value={input.total_income_incl} onChange={e => setInput({...input, total_income_incl: Number(e.target.value)})} className="bg-card border border-border text-foreground px-3.5 py-2.5 rounded-md outline-none text-sm font-semibold focus:border-primary focus:ring-2 focus:ring-primary/20 transition-all" />
              </div>
              <div className="flex flex-col gap-2">
                <label className="text-xs font-bold text-secondary-foreground">IT税率</label>
                <input type="number" step={0.01} value={input.tax_rate_it} onChange={e => setInput({...input, tax_rate_it: Number(e.target.value)})} className="bg-card border border-border text-foreground px-3.5 py-2.5 rounded-md outline-none text-sm font-semibold focus:border-primary focus:ring-2 focus:ring-primary/20 transition-all" />
              </div>
              <div className="flex flex-col gap-2">
                <label className="text-xs font-bold text-secondary-foreground">CT税率</label>
                <input type="number" step={0.01} value={input.tax_rate_ct} onChange={e => setInput({...input, tax_rate_ct: Number(e.target.value)})} className="bg-card border border-border text-foreground px-3.5 py-2.5 rounded-md outline-none text-sm font-semibold focus:border-primary focus:ring-2 focus:ring-primary/20 transition-all" />
              </div>
              <div className="flex flex-col gap-2">
                <label className="text-xs font-bold text-secondary-foreground">CT产品含税总额 (选填)</label>
                <input type="number" step={0.01} placeholder="未填由系统智能推算" value={input.ct_income_incl_opt} onChange={e => setInput({...input, ct_income_incl_opt: e.target.value})} className="bg-card border border-border text-foreground px-3.5 py-2.5 rounded-md outline-none text-sm font-semibold focus:border-primary focus:ring-2 focus:ring-primary/20 transition-all" />
              </div>

              <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70 mt-2">模式选择</h3>
              <div className="flex flex-col gap-3 bg-card border border-border p-3 rounded-lg shadow-sm">
                <label className="flex items-center gap-2 text-sm font-semibold cursor-pointer">
                  <input type="radio" name="mode" checked={input.calc_mode === 'margin'} onChange={() => setInput({...input, calc_mode: 'margin'})} className="accent-primary" />
                  [模式1] 指定毛利润反推投入
                </label>
                <label className="flex items-center gap-2 text-sm font-semibold cursor-pointer">
                  <input type="radio" name="mode" checked={input.calc_mode === 'npv'} onChange={() => setInput({...input, calc_mode: 'npv'})} className="accent-primary" />
                  [模式2] 指定净现值率反推投入
                </label>
                <label className="flex items-center gap-2 text-sm font-semibold cursor-pointer">
                  <input type="radio" name="mode" checked={input.calc_mode === 'total_cost'} onChange={() => setInput({...input, calc_mode: 'total_cost'})} className="accent-primary" />
                  [模式3] 已知总投入正算效益
                </label>
              </div>

              <div className="flex flex-col gap-2 mt-2">
                <label className="text-xs font-bold text-secondary-foreground">目标值 (小数或金额)</label>
                <input type="number" step={0.0001} value={input.target_value} onChange={e => setInput({...input, target_value: Number(e.target.value)})} className="bg-card border border-border text-foreground px-3.5 py-2.5 rounded-md outline-none text-sm font-semibold focus:border-primary focus:ring-2 focus:ring-primary/20 transition-all" />
              </div>

              <button onClick={handleCalc} className="bg-gradient-to-b from-primary to-primary/90 text-primary-foreground font-semibold py-3 px-6 rounded-md shadow-sm hover:shadow-md hover:-translate-y-[1px] transition-all mt-4 w-full">
                开始测算
              </button>
              {warning && (
                <div className="bg-destructive/20 text-destructive-foreground px-4 py-3 rounded-md mt-2 text-sm font-semibold">
                  ⚠️ {warning}
                </div>
              )}
            </div>
          ) : (
             <div className="flex flex-col gap-4">
              <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70 mt-2">批量处理说明</h3>
              <p className="text-sm text-secondary-foreground leading-relaxed">
                批量模式下，系统将读取 Excel 中的「项目名称」、「项目总收入」、「项目总投入」等列进行自动测算并生成报告。
              </p>
             </div>
          )}
        </div>

        <div className="flex-1 p-8 overflow-y-auto bg-background">
          {subView === "single" ? (
            <div>
              <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70 mb-4">测算结果账单</h3>
              <div className="bg-card border border-border rounded-xl p-6 shadow-sm overflow-x-auto">
                <table className="w-full text-sm text-right border-separate border-spacing-y-2">
                  <thead>
                    <tr className="text-secondary-foreground font-bold uppercase text-xs">
                      <th className="text-left pb-2">资金流向明细</th>
                      <th className="pb-2">含税总收入</th>
                      <th className="pb-2">不含税收入</th>
                      <th className="pb-2">含税总投入</th>
                      <th className="pb-2">不含税投入</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr className="bg-muted">
                      <td className="text-left p-3 rounded-l-md font-semibold">IT集成部分</td>
                      <td className="p-3">{formatCurrency(result?.it_income_incl)}</td>
                      <td className="p-3">{formatCurrency(result?.it_income_excl)}</td>
                      <td className="p-3">{formatCurrency(result?.it_cost_incl)}</td>
                      <td className="p-3 rounded-r-md">{formatCurrency(result?.it_cost_excl)}</td>
                    </tr>
                    <tr className="bg-muted">
                      <td className="text-left p-3 rounded-l-md font-semibold">CT产品部分</td>
                      <td className="p-3">{formatCurrency(result?.ct_income_incl)}</td>
                      <td className="p-3">{formatCurrency(result?.ct_income_excl)}</td>
                      <td className="p-3">{formatCurrency(result?.ct_cost_incl)}</td>
                      <td className="p-3 rounded-r-md">{formatCurrency(result?.ct_cost_excl)}</td>
                    </tr>
                    <tr className="bg-primary/10 text-primary font-bold">
                      <td className="text-left p-3 rounded-l-md">项目合并总计</td>
                      <td className="p-3">{formatCurrency(result?.total_income_incl)}</td>
                      <td className="p-3">{formatCurrency(result?.total_income_excl)}</td>
                      <td className="p-3">{formatCurrency(result?.total_cost_incl)}</td>
                      <td className="p-3 rounded-r-md">{formatCurrency(result?.total_cost_excl)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div className="grid grid-cols-3 gap-4 mt-6">
                <div className="bg-muted rounded-xl p-6">
                  <div className="text-xs uppercase font-bold text-secondary-foreground">项目毛利润率</div>
                  <div className="text-3xl font-extrabold text-primary mt-2">{formatPercent(result?.margin_rate)}</div>
                </div>
                <div className="bg-muted rounded-xl p-6">
                  <div className="text-xs uppercase font-bold text-secondary-foreground">项目整体 NPV 回报率</div>
                  <div className="text-3xl font-extrabold text-primary mt-2">{formatPercent(result?.npv_rate)}</div>
                </div>
                <div className="bg-muted rounded-xl p-6">
                  <div className="text-xs uppercase font-bold text-secondary-foreground">仅 IT 侧 NPV 回报率</div>
                  <div className="text-3xl font-extrabold text-primary mt-2">{formatPercent(result?.it_npv_rate)}</div>
                </div>
              </div>
            </div>
          ) : (
            <div>
              <div className="flex justify-between items-center mb-6">
                <h3 className="text-xl font-bold text-foreground m-0">批量导入处理</h3>
              </div>
              <div 
                className="border-2 border-dashed border-border bg-muted/50 hover:bg-muted hover:border-primary transition-all rounded-xl p-12 text-center cursor-pointer flex flex-col items-center gap-4"
                onClick={handleBatch}
              >
                <div className="text-5xl">📁</div>
                <div className="font-semibold text-lg">点击选择本地 Excel (.xlsx)</div>
                <div className="text-sm text-secondary-foreground max-w-md">支持一键处理多项目清单，系统将自动识别运算模式并生成批处理结果。</div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
