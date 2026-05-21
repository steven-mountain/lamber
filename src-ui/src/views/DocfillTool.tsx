import { useState, useEffect, useCallback } from "react"
import { invoke } from "@tauri-apps/api/core"
import WorkspaceHeader from "../components/WorkspaceHeader"
import AppIcon from "../components/icons/AppIcon"
import { useAiContextStore } from "../store/useAiContextStore"
import { useRef } from "react"
import { AI_CONTEXT_KEY, buildAiContextKey } from "../utils/aiContextKeys"

export default function DocfillTool({ onBack }: { onBack: () => void }) {
  const [templateName, setTemplateName] = useState("未选择模板")
  const [templatePath, setTemplatePath] = useState("")
  const [variables, setVariables] = useState<string[]>([])
  const [formData, setFormData] = useState<Record<string, string>>({})
  const [availableTemplates, setAvailableTemplates] = useState<string[]>([])

  const loadTemplates = useCallback(async () => {
    try {
      const list: string[] = await invoke('get_available_templates', { moduleId: 'docfill_tool' })
      setAvailableTemplates(list)
    } catch (e) {
      console.error("加载模板失败:", e)
    }
  }, [])

  useEffect(() => {
    loadTemplates()
  }, [loadTemplates])

  const handleSelectTemplate = async (name: string) => {
    try {
      // Get the absolute path for extraction
      const modulePath: string = await invoke('get_module_path', { moduleId: 'docfill_tool' })
      const fullPath = `${modulePath}/templates/${name}`
      
      setTemplatePath(fullPath)
      setTemplateName(name)
      const vars: string[] = await invoke('extract_docx_variables', { path: fullPath })
      setVariables(vars)
      const initialForm: Record<string, string> = {}
      vars.forEach(v => initialForm[v] = "")
      setFormData(initialForm)
    } catch (e) {
      console.error(e)
      alert("加载模板变量失败: " + e)
    }
  }

  // --- AI Context Sync for Docfill ---
  const updateData = useAiContextStore(state => state.updateBusinessData);
  const setActiveModule = useAiContextStore(state => state.setActiveModule);
  const syncTimerRef = useRef<NodeJS.Timeout | null>(null);
  const buildDocfillPayload = () => ({
    module: 'docfill',
    template_name: templateName,
    template_path: templatePath,
    available_templates: availableTemplates,
    variables,
    form_data: formData,
    has_template_selected: Boolean(templatePath),
  });

  useEffect(() => {
    const nextModule = templatePath
      ? buildAiContextKey('docfill', 'template', templateName)
      : AI_CONTEXT_KEY.DOCFILL_CORE;
    setActiveModule(nextModule);
  }, [templateName, templatePath, setActiveModule]);

  useEffect(() => {
    updateData(AI_CONTEXT_KEY.DOCFILL_CORE, buildDocfillPayload());
  }, [availableTemplates, formData, templateName, templatePath, updateData, variables]);

  useEffect(() => {
    // Debounced sync (500ms)
    if (syncTimerRef.current) clearTimeout(syncTimerRef.current);
    
    syncTimerRef.current = setTimeout(() => {
      const payload = buildDocfillPayload();
      updateData(AI_CONTEXT_KEY.DOCFILL_CORE, payload);
      if (!templatePath) return;
      const templateId = buildAiContextKey('docfill', 'template', templateName);
      updateData(templateId, payload);
      console.log(`AI Context Synced: ${templateId}`);
    }, 500);

    return () => {
      if (syncTimerRef.current) clearTimeout(syncTimerRef.current);
    };
  }, [availableTemplates, templateName, templatePath, variables, formData, updateData]);

  const handleGenerate = async () => {
    if (!templatePath) return alert("请先选择模板")
    try {
      const outPath: string = await invoke('plugin:dialog|save', {
        filters: [{ name: 'Word Document', extensions: ['docx'] }]
      })
      
      if (outPath) {
        await invoke('generate_docx', {
          template_path: templatePath,
          variables: formData,
          output_path: outPath
        })
        alert("生成成功：" + outPath)
      }
    } catch (e) {
      alert("生成失败: " + e)
    }
  }

  return (
    <div className="flex flex-col flex-1 animate-in fade-in duration-300 h-full">
      <WorkspaceHeader moduleId="docfill_tool" title="申报材料制作" onBack={onBack} onPathChange={() => loadTemplates()} />
      
      <div className="flex flex-1 overflow-hidden">
        <div className="w-[340px] bg-muted p-6 overflow-y-auto flex flex-col gap-4 border-r border-border shrink-0">
          <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70">1. 选择目录下的模板</h3>
          
          <div className="flex flex-col gap-2">
            {availableTemplates.length === 0 ? (
              <div className="text-xs text-secondary-foreground p-4 bg-card rounded-lg border border-dashed border-border">
                templates 目录下未找到模板文件
              </div>
            ) : (
              availableTemplates.map(t => (
                <button 
                  key={t}
                  onClick={() => handleSelectTemplate(t)}
                  className={`text-left px-4 py-3 rounded-lg text-sm font-semibold transition-all border ${templateName === t ? 'bg-blue-50 border-blue-200 text-primary shadow-sm' : 'bg-card border-border hover:bg-secondary text-foreground'}`}
                >
                  <span className="inline-flex items-center gap-2">
                    <AppIcon name="document" size={16} />
                    {t}
                  </span>
                </button>
              ))
            )}
          </div>
          
          {templatePath && (
            <div className="mt-4 pt-4 border-t border-border flex flex-col gap-4">
               <div className="text-[10px] uppercase font-extrabold text-secondary-foreground opacity-60">当前已选</div>
               <div className="text-xs text-primary font-bold break-all bg-primary/5 p-2 rounded border border-primary/20">{templateName}</div>
               <button 
                onClick={handleGenerate}
                className="flex w-full items-center justify-center gap-2 bg-primary hover:bg-primary/90 text-primary-foreground font-semibold py-3 px-6 rounded-md shadow-sm hover:shadow-md hover:-translate-y-[1px] transition-all"
              >
                <AppIcon name="generate" size={18} /> 执行生成
              </button>
            </div>
          )}
        </div>
        
        <div className="flex-1 p-8 overflow-y-auto bg-background">
          <div className="bg-card border border-border rounded-xl p-8 shadow-sm h-full">
            <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70 mb-2">2. 填写材料信息</h3>
            <p className="text-sm text-secondary-foreground mb-6">
              请在左侧选择模板后，系统会自动提取变量（形如 {'{变量名}'} ）。
            </p>
            
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              {variables.map(variable => (
                <div key={variable} className="flex flex-col gap-2">
                  <label className="text-sm font-bold text-secondary-foreground">{variable}</label>
                  <input 
                    type="text" 
                    value={formData[variable]} 
                    onChange={e => setFormData({...formData, [variable]: e.target.value})}
                    className="bg-card border border-input text-foreground px-3.5 py-2.5 rounded-md outline-none text-sm font-semibold focus:border-ring focus:ring-2 focus:ring-ring/20 transition-all"
                  />
                </div>
              ))}
            </div>
            {variables.length === 0 && (
              <div className="text-center text-secondary-foreground py-12 bg-muted rounded-lg border border-dashed border-border">
                等待提取模板变量...
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
