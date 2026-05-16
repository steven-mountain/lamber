import { useState } from "react"
import BenefitTool from "./views/BenefitTool"
import DocfillTool from "./views/DocfillTool"
import IctLifecycle from "./views/IctLifecycle"
import AiConsultantDrawer from "./components/AiConsultantDrawer"

export default function App() {
  const [currentView, setCurrentView] = useState("hub")

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
      
      {/* Global AI Consultant */}
      <AiConsultantDrawer />
    </div>
  )
}

function HubView({ onOpenTool }: { onOpenTool: (view: string) => void }) {
  return (
    <div className="p-10 flex flex-col items-center relative flex-1 animate-in fade-in duration-500">
      <div className="absolute top-5 left-6 font-bold text-foreground flex items-center gap-2 before:content-[''] before:w-1 before:h-4 before:bg-primary before:rounded-sm">
        云数中心工具集
      </div>
      <div className="mt-24 text-center mb-16">
        <h1 className="text-4xl font-extrabold mb-2 text-foreground tracking-tight">云数中心工具集</h1>
        <p className="text-secondary-foreground font-medium">请选择需要使用的工具模块</p>
      </div>
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8 w-full max-w-4xl">
        <div 
          className="bg-card border border-border shadow-sm hover:shadow-lg rounded-2xl p-8 flex flex-col items-center text-center cursor-pointer transition-all hover:-translate-y-1 hover:border-primary/50"
          onClick={() => onOpenTool("benefit")}
        >
          <div className="w-16 h-16 bg-secondary text-primary rounded-2xl flex items-center justify-center text-3xl mb-5 transition-colors">
            📊
          </div>
          <div className="font-bold text-lg mb-1">项目效益分析</div>
          <div className="text-sm text-secondary-foreground">测算项目经济效益</div>
        </div>
        <div 
          className="bg-card border border-border shadow-sm hover:shadow-lg rounded-2xl p-8 flex flex-col items-center text-center cursor-pointer transition-all hover:-translate-y-1 hover:border-primary/50"
          onClick={() => onOpenTool("docfill")}
        >
          <div className="w-16 h-16 bg-green-50 text-green-600 rounded-2xl flex items-center justify-center text-3xl mb-5 transition-colors">
            📄
          </div>
          <div className="font-bold text-lg mb-1">申报材料制作</div>
        </div>
        <div 
          className="bg-card border border-border shadow-sm hover:shadow-lg rounded-2xl p-8 flex flex-col items-center text-center cursor-pointer transition-all hover:-translate-y-1 hover:border-primary/50"
          onClick={() => onOpenTool("ict")}
        >
          <div className="w-16 h-16 bg-amber-50 text-amber-600 rounded-2xl flex items-center justify-center text-3xl mb-5 transition-colors">
            🔄
          </div>
          <div className="font-bold text-lg mb-1">ICT项目全生命周期</div>
          <div className="text-sm text-secondary-foreground mt-1">经济效益与过程评估</div>
        </div>
      </div>
    </div>
  )
}


