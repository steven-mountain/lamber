import { useEffect, useState } from "react"
import { invoke } from "@tauri-apps/api/core"
import BenefitTool from "./views/BenefitTool"
import DocfillTool from "./views/DocfillTool"
import IctLifecycle from "./views/IctLifecycle"
import ProjectBoard from "./views/ProjectBoard"
import DataManagement from "./views/DataManagement"
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

import { useNavigationStore } from "./store/useNavigationStore"

export default function App() {
  const { currentView, navigateTo } = useNavigationStore()
  const setActiveModule = useAiContextStore(state => state.setActiveModule)
  const aiAssistantView = getAiAssistantView()

  const [showMigrationModal, setShowMigrationModal] = useState(false)
  const [isMigrating, setIsMigrating] = useState(false)
  const [migrationReport, setMigrationReport] = useState<any | null>(null)

  useEffect(() => {
    if (aiAssistantView) return

    // Run startup check for SQLite database migration
    if (isTauriRuntime()) {
      invoke<boolean>("check_db_migration")
        .then((needed) => {
          if (needed) {
            setShowMigrationModal(true)
          }
        })
        .catch((err) => console.error("Failed to check db migration:", err))
    }
  }, [aiAssistantView])

  useEffect(() => {
    if (aiAssistantView) return

    localStorage.setItem(AI_CURRENT_VIEW_KEY, currentView)
    if (currentView === "hub") {
      setActiveModule("hub")
    } else if (currentView === "project_board" || currentView === "ict_lifecycle") {
      setActiveModule("ict")
    } else {
      setActiveModule(currentView)
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
      {/* SQLite Migration Modal */}
      {showMigrationModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm animate-in fade-in duration-300">
          <div className="w-full max-w-lg rounded-2xl bg-card p-8 shadow-2xl animate-in zoom-in-95 slide-in-from-bottom-4 duration-300">
            <h2 className="text-xl font-bold mb-3 flex items-center gap-2">
              <span className="inline-block w-2.5 h-6 rounded bg-primary"></span>
              数据库升级与迁移
            </h2>
            
            {!migrationReport ? (
              <>
                <p className="text-secondary-foreground text-sm leading-relaxed mb-6">
                  系统已引入 <strong>SQLite</strong> 数据库作为核心持久化存储，提升数据读写速度与一致性。
                  检测到您存在旧版 <code>projects_store.json</code>。是否立即迁移数据？
                </p>
                <div className="rounded-xl bg-muted/60 p-4 mb-6 text-xs text-secondary-foreground space-y-2">
                  <div className="flex items-center gap-2 font-semibold">
                    <span className="w-1.5 h-1.5 rounded-full bg-primary"></span>
                    安全自动备份：系统会在同目录下创建原数据的备份文件。
                  </div>
                  <div className="flex items-center gap-2 font-semibold">
                    <span className="w-1.5 h-1.5 rounded-full bg-primary"></span>
                    零损事务保障：迁移使用数据库事务，出错将自动完整回滚。
                  </div>
                </div>
                
                <div className="flex justify-end gap-3">
                  <button
                    disabled={isMigrating}
                    onClick={async () => {
                      try {
                        await invoke("skip_db_migration")
                        setShowMigrationModal(false)
                      } catch (err) {
                        console.error("Failed to skip migration:", err)
                      }
                    }}
                    className="px-4 py-2 rounded-xl text-sm font-semibold bg-muted hover:bg-muted/80 active:scale-[0.98] transition-all disabled:opacity-50"
                  >
                    暂不迁移
                  </button>
                  <button
                    disabled={isMigrating}
                    onClick={async () => {
                      setIsMigrating(true)
                      try {
                        const report = await invoke("run_db_migration")
                        setMigrationReport(report)
                      } catch (err: any) {
                        alert("迁移失败: " + err)
                        setIsMigrating(false)
                      }
                    }}
                    className="px-5 py-2 rounded-xl text-sm font-semibold bg-primary text-primary-foreground hover:bg-primary/90 active:scale-[0.98] transition-all flex items-center gap-2 disabled:opacity-50"
                  >
                    {isMigrating ? (
                      <>
                        <span className="animate-spin rounded-full h-4 w-4 border-2 border-primary-foreground border-t-transparent"></span>
                        正在迁移...
                      </>
                    ) : (
                      "立即迁移"
                    )}
                  </button>
                </div>
              </>
            ) : (
              <>
                <div className="mb-6 space-y-4">
                  <div className="flex items-center gap-2 text-green-500 font-semibold">
                    <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
                      <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                    <span>迁移成功完成</span>
                  </div>
                  
                  <div className="rounded-xl bg-muted/60 p-4 text-xs font-mono space-y-1.5 text-secondary-foreground">
                    <div>项目统计: <span className="font-bold tabular-nums">{migrationReport.projectsCount}</span> 个</div>
                    <div>方案统计: <span className="font-bold tabular-nums">{migrationReport.schemesCount}</span> 个</div>
                    <div>历史快照: <span className="font-bold tabular-nums">{migrationReport.snapshotsCount}</span> 个</div>
                    <div>文件关联: <span className="font-bold tabular-nums">{migrationReport.filesCount}</span> 个</div>
                    <div className="pt-2 text-[10px] break-all border-t border-border mt-2">
                      备份路径: <br />
                      <span className="text-secondary-foreground/75 select-all">{migrationReport.backupPath}</span>
                    </div>
                  </div>
                </div>

                <div className="flex justify-end">
                  <button
                    onClick={() => {
                      setShowMigrationModal(false)
                      // Force reload to reload lists
                      window.location.reload()
                    }}
                    className="px-5 py-2 rounded-xl text-sm font-semibold bg-primary text-primary-foreground hover:bg-primary/90 active:scale-[0.98] transition-all"
                  >
                    开始体验
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}
      {currentView === "hub" ? (
        <HubView onOpenTool={(view) => navigateTo(view as any)} />
      ) : currentView === "benefit" ? (
        <BenefitTool onBack={() => navigateTo("hub")} />
      ) : currentView === "docfill" ? (
        <DocfillTool onBack={() => navigateTo("hub")} />
      ) : currentView === "project_board" ? (
        <ProjectBoard
          onBack={() => navigateTo("hub")}
          onOpenCalc={(projectId, schemeId) => navigateTo("ict_lifecycle", projectId, schemeId)}
        />
      ) : currentView === "ict_lifecycle" ? (
        <IctLifecycle />
      ) : currentView === "data_management" ? (
        <DataManagement onBack={() => navigateTo("hub")} />
      ) : (
        <div className="p-8">
          <button onClick={() => navigateTo("hub")} className="mb-4 text-primary font-bold">← 返回</button>
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
      <div className="grid w-full max-w-5xl grid-cols-[repeat(auto-fit,minmax(220px,1fr))] gap-6 pb-8">
        <div
          className="flex min-h-[180px] cursor-pointer flex-col items-center justify-center rounded-2xl border border-border bg-card p-6 text-center shadow-sm transition-all hover:-translate-y-1 hover:border-primary/50 hover:shadow-lg animate-in slide-in-from-bottom duration-300"
          onClick={() => onOpenTool("benefit")}
        >
          <div className="mb-5 flex h-16 w-16 items-center justify-center rounded-2xl bg-secondary text-primary transition-colors">
            <AppIcon name="calculator" size={30} />
          </div>
          <div className="font-bold text-lg mb-1">投资效益测算</div>
          <div className="text-sm text-secondary-foreground">测算项目经济效益</div>
        </div>
        <div
          className="flex min-h-[180px] cursor-pointer flex-col items-center justify-center rounded-2xl border border-border bg-card p-6 text-center shadow-sm transition-all hover:-translate-y-1 hover:border-primary/50 hover:shadow-lg animate-in slide-in-from-bottom duration-300 delay-75"
          onClick={() => onOpenTool("docfill")}
        >
          <div className="mb-5 flex h-16 w-16 items-center justify-center rounded-2xl bg-secondary text-primary transition-colors">
            <AppIcon name="document" size={30} />
          </div>
          <div className="font-bold text-lg mb-1">申报材料制作</div>
          <div className="text-sm text-secondary-foreground">快速生成项目申报方案</div>
        </div>
        <div
          className="flex min-h-[180px] cursor-pointer flex-col items-center justify-center rounded-2xl border border-border bg-card p-6 text-center shadow-sm transition-all hover:-translate-y-1 hover:border-primary/50 hover:shadow-lg animate-in slide-in-from-bottom duration-300 delay-150"
          onClick={() => onOpenTool("project_board")}
        >
          <div className="mb-5 flex h-16 w-16 items-center justify-center rounded-2xl bg-secondary text-primary transition-colors">
            <AppIcon name="project" size={30} />
          </div>
          <div className="font-bold text-lg mb-1">项目管理看板</div>
          <div className="text-sm text-secondary-foreground mt-1">管理项目生命周期、文件及方案</div>
        </div>
        <div
          className="flex min-h-[180px] cursor-pointer flex-col items-center justify-center rounded-2xl border border-border bg-card p-6 text-center shadow-sm transition-all hover:-translate-y-1 hover:border-primary/50 hover:shadow-lg animate-in slide-in-from-bottom duration-300 delay-200"
          onClick={() => onOpenTool("ict_lifecycle")}
        >
          <div className="mb-5 flex h-16 w-16 items-center justify-center rounded-2xl bg-secondary text-primary transition-colors">
            <AppIcon name="cashflow" size={30} />
          </div>
          <div className="font-bold text-lg mb-1">ICT项目全生命周期</div>
          <div className="text-sm text-secondary-foreground mt-1">测算、现金流推演与智能反算</div>
        </div>
        <div
          className="flex min-h-[180px] cursor-pointer flex-col items-center justify-center rounded-2xl border border-border bg-card p-6 text-center shadow-sm transition-all hover:-translate-y-1 hover:border-primary/50 hover:shadow-lg animate-in slide-in-from-bottom duration-300 delay-300"
          onClick={() => onOpenTool("data_management")}
        >
          <div className="mb-5 flex h-16 w-16 items-center justify-center rounded-2xl bg-secondary text-primary transition-colors">
            <AppIcon name="settings" size={30} />
          </div>
          <div className="font-bold text-lg mb-1">数据管理中心</div>
          <div className="text-sm text-secondary-foreground mt-1 font-medium">配置根目录、重定位与健康自愈</div>
        </div>
      </div>
    </div>
  )
}


