import { useEffect, useState } from "react"
import BenefitTool from "./views/BenefitTool"
import DocfillTool from "./views/DocfillTool"
import IctLifecycle from "./views/IctLifecycle"
import AiFloatingLauncher from "./components/ai/AiFloatingLauncher"
import AiFloatingWindow from "./components/ai/AiFloatingWindow"
import AppIcon from "./components/icons/AppIcon"
import { useAiContextStore } from "./store/useAiContextStore"
import { emitTo } from "@tauri-apps/api/event"

const AI_ASSISTANT_LABEL = "ai-assistant"
const AI_CURRENT_VIEW_KEY = "lamber_ai_current_view"

function isTauriRuntime() {
  return typeof window !== "undefined" && Boolean((window as Window & { __TAURI_INTERNALS__?: unknown }).__TAURI_INTERNALS__)
}

function getAiAssistantView() {
  const hash = window.location.hash
  if (!hash.startsWith("#/ai-assistant")) return null

  const query = hash.includes("?") ? hash.slice(hash.indexOf("?") + 1) : ""
  const view = new URLSearchParams(query).get("view")
  return view || "hub"
}

export default function App() {
  const [currentView, setCurrentView] = useState("hub")
  const setActiveModule = useAiContextStore(state => state.setActiveModule)
  const aiAssistantView = getAiAssistantView()

  useEffect(() => {
    if (aiAssistantView) return

    localStorage.setItem(AI_CURRENT_VIEW_KEY, currentView)
    if (currentView === "hub") {
      setActiveModule("hub")
    }

    if (isTauriRuntime()) {
      emitTo(AI_ASSISTANT_LABEL, "lamber-ai-view-changed", { view: currentView })
        .catch((error) => console.warn("Failed to sync AI assistant view:", error))
    }
  }, [aiAssistantView, currentView, setActiveModule])

  if (aiAssistantView) {
    return <AiFloatingWindow currentView={aiAssistantView} />
  }

  return (
    <div className="flex flex-col h-screen overflow-hidden bg-background text-foreground">
      {currentView === "hub" ? (
        <HubView onOpenTool={setCurrentView} />
      ) : currentView === "benefit" ? (
        <BenefitTool onBack={() => setCurrentView("hub")} />
      ) : currentView === "docfill" ? (
        <DocfillTool onBack={() => setCurrentView("hub")} />
      ) : currentView === "ict" ? (
        <IctLifecycle onBack={() => setCurrentView("hub")} />
      ) : (
        <div className="p-8">
          <button onClick={() => setCurrentView("hub")} className="mb-4 text-primary font-bold">← 返回</button>
          <p>模块正在开发中...</p>
        </div>
      )}
      
      <AiFloatingLauncher currentView={currentView} />
    </div>
  )
}

function HubView({ onOpenTool }: { onOpenTool: (view: string) => void }) {
  return (
    <div className="relative flex min-h-0 flex-1 flex-col items-center overflow-y-auto px-6 py-8 animate-in fade-in duration-500 md:px-10">
      <div className="absolute left-6 top-5 flex items-center gap-2 font-bold text-foreground before:h-4 before:w-1 before:rounded-sm before:bg-primary before:content-[''] md:left-10">
        云数中心工具集
      </div>
      <div className="mb-10 mt-20 text-center md:mb-12">
        <h1 className="text-4xl font-extrabold mb-2 text-foreground tracking-tight">云数中心工具集</h1>
        <p className="text-secondary-foreground font-medium">请选择需要使用的工具模块</p>
      </div>
      <div className="grid w-full max-w-5xl grid-cols-[repeat(auto-fit,minmax(240px,1fr))] gap-6 pb-8">
        <div 
          className="flex min-h-[180px] cursor-pointer flex-col items-center justify-center rounded-2xl border border-border bg-card p-6 text-center shadow-sm transition-all hover:-translate-y-1 hover:border-primary/50 hover:shadow-lg"
          onClick={() => onOpenTool("benefit")}
        >
          <div className="mb-5 flex h-16 w-16 items-center justify-center rounded-2xl bg-secondary text-primary transition-colors">
            <AppIcon name="calculator" size={30} />
          </div>
          <div className="font-bold text-lg mb-1">项目效益分析</div>
          <div className="text-sm text-secondary-foreground">测算项目经济效益</div>
        </div>
        <div 
          className="flex min-h-[180px] cursor-pointer flex-col items-center justify-center rounded-2xl border border-border bg-card p-6 text-center shadow-sm transition-all hover:-translate-y-1 hover:border-primary/50 hover:shadow-lg"
          onClick={() => onOpenTool("docfill")}
        >
          <div className="mb-5 flex h-16 w-16 items-center justify-center rounded-2xl bg-secondary text-primary transition-colors">
            <AppIcon name="document" size={30} />
          </div>
          <div className="font-bold text-lg mb-1">申报材料制作</div>
        </div>
        <div 
          className="flex min-h-[180px] cursor-pointer flex-col items-center justify-center rounded-2xl border border-border bg-card p-6 text-center shadow-sm transition-all hover:-translate-y-1 hover:border-primary/50 hover:shadow-lg"
          onClick={() => onOpenTool("ict")}
        >
          <div className="mb-5 flex h-16 w-16 items-center justify-center rounded-2xl bg-secondary text-primary transition-colors">
            <AppIcon name="cashflow" size={30} />
          </div>
          <div className="font-bold text-lg mb-1">ICT项目全生命周期</div>
          <div className="text-sm text-secondary-foreground mt-1">经济效益与过程评估</div>
        </div>
      </div>
    </div>
  )
}


