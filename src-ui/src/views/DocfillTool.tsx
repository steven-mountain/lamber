import { useState } from "react"
import { invoke } from "@tauri-apps/api/core"

export default function DocfillTool({ onBack }: { onBack: () => void }) {
  const [templateName, setTemplateName] = useState("未选择模板")
  const [templatePath, setTemplatePath] = useState("")
  const [variables, setVariables] = useState<string[]>([])
  const [formData, setFormData] = useState<Record<string, string>>({})

  const handleSelectTemplate = async () => {
    try {
      const selected = await invoke('plugin:dialog|open', {
        multiple: false,
        filters: [{ name: 'Word Documents', extensions: ['docx'] }]
      }) as string
      
      if (selected) {
        setTemplatePath(selected)
        setTemplateName(selected)
        const vars: string[] = await invoke('extract_docx_variables', { path: selected })
        setVariables(vars)
        const initialForm: Record<string, string> = {}
        vars.forEach(v => initialForm[v] = "")
        setFormData(initialForm)
      }
    } catch (e) {
      console.error(e)
    }
  }

  const handleGenerate = async () => {
    if (!templatePath) return alert("请先选择模板")
    try {
      const outPath = await invoke('plugin:dialog|save', {
        filters: [{ name: 'Word Document', extensions: ['docx'] }]
      }) as string
      
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
      <div className="h-16 bg-card border-b border-border flex items-center px-6 gap-4 shrink-0">
        <button onClick={onBack} className="text-secondary-foreground hover:text-primary hover:bg-secondary font-semibold flex items-center gap-1.5 px-3 py-2 rounded-lg transition-colors">
          <span>←</span> 返回集市
        </button>
        <h2 className="m-0 text-lg font-bold text-foreground border-l-2 border-border pl-4">申报材料制作</h2>
      </div>
      
      <div className="flex flex-1 overflow-hidden">
        <div className="w-[340px] bg-muted p-6 overflow-y-auto flex flex-col gap-4 border-r border-border shrink-0">
          <h3 className="text-xs uppercase tracking-wide font-extrabold text-secondary-foreground opacity-70">1. 上传模板文件</h3>
          
          <div 
            onClick={handleSelectTemplate}
            className="border-2 border-dashed border-border rounded-xl p-8 text-center bg-card hover:border-primary hover:bg-muted cursor-pointer transition-all"
          >
            <div className="text-3xl mb-2">📄</div>
            <div className="text-sm font-semibold">选择本地 Docx 模板</div>
          </div>
          
          <div className="text-xs text-primary font-medium break-all">{templateName}</div>
          
          <button 
            onClick={handleGenerate}
            className="bg-gradient-to-b from-primary to-primary/90 text-primary-foreground font-semibold py-3 px-6 rounded-md shadow-sm hover:shadow-md hover:-translate-y-[1px] transition-all mt-4 w-full"
          >
            执行生成
          </button>
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
                    className="bg-muted border border-border text-foreground px-3.5 py-2.5 rounded-md outline-none text-sm font-semibold focus:border-primary focus:ring-2 focus:ring-primary/20 transition-all"
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
