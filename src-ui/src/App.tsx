import { useEffect } from "react";
import { emitTo } from "@tauri-apps/api/event";
import IctLifecycle from "./views/IctLifecycle";
import ProjectBoard from "./views/ProjectBoard";
import DataManagement from "./views/DataManagement";
import AiFloatingLauncher from "./components/ai/AiFloatingLauncher";
import AiFloatingWindow from "./components/ai/AiFloatingWindow";
import AppIcon, { type AppIconName } from "./components/icons/AppIcon";
import WorkspaceGate from "./components/workspace/WorkspaceGate";
import { useAiContextStore } from "./store/useAiContextStore";
import { useNavigationStore } from "./store/useNavigationStore";
import { useWorkspaceStore } from "./store/useWorkspaceStore";
import { useGlobalSaveShortcut } from "./hooks/useGlobalSaveShortcut";
import { useUnsavedChangesGuard } from "./hooks/useUnsavedChangesGuard";

const AI_ASSISTANT_LABEL = "ai-assistant";
const AI_CURRENT_VIEW_KEY = "lamber_ai_current_view";

function isTauriRuntime() {
  return typeof window !== "undefined" && Boolean((window as Window & { __TAURI_INTERNALS__?: unknown }).__TAURI_INTERNALS__);
}

function getAiAssistantView() {
  const hash = window.location.hash;
  if (!hash.startsWith("#/ai-assistant")) return null;

  const query = hash.includes("?") ? hash.slice(hash.indexOf("?") + 1) : "";
  const view = new URLSearchParams(query).get("view");
  return view || "hub";
}

export default function App() {
  const { currentView, navigateTo } = useNavigationStore();
  const setActiveModule = useAiContextStore(state => state.setActiveModule);
  const { isWorkspaceReady, refreshWorkspaceState } = useWorkspaceStore();
  const aiAssistantView = getAiAssistantView();
  useGlobalSaveShortcut();
  useUnsavedChangesGuard();

  useEffect(() => {
    if (aiAssistantView) return;
    if (isTauriRuntime()) refreshWorkspaceState();
  }, [aiAssistantView, refreshWorkspaceState]);

  useEffect(() => {
    if (aiAssistantView) return;

    localStorage.setItem(AI_CURRENT_VIEW_KEY, currentView);
    if (currentView === "hub") {
      setActiveModule("hub");
    } else if (currentView === "project_board") {
      setActiveModule("project_board.core");
    } else if (currentView === "ict_lifecycle") {
      setActiveModule("ict");
    } else {
      setActiveModule(currentView);
    }

    if (isTauriRuntime()) {
      emitTo(AI_ASSISTANT_LABEL, "lamber-ai-view-changed", { view: currentView })
        .catch(error => console.warn("Failed to sync AI assistant view:", error));
    }
  }, [aiAssistantView, currentView, setActiveModule]);

  if (aiAssistantView) {
    return <AiFloatingWindow currentView={aiAssistantView} />;
  }

  return (
    <div className="flex h-screen flex-col overflow-hidden bg-background text-foreground">
      {currentView === "hub" ? (
        <HubView onOpenTool={(view) => navigateTo(view as any)} />
      ) : currentView === "project_board" ? (
        <ProjectBoard
          onBack={() => navigateTo("hub")}
          onOpenCalc={(projectId, schemeId) => navigateTo("ict_lifecycle", projectId, schemeId)}
        />
      ) : currentView === "ict_lifecycle" ? (
        <IctLifecycle />
      ) : currentView === "data_management" ? (
        isWorkspaceReady ? (
          <DataManagement onBack={() => navigateTo("hub")} />
        ) : (
          <div className="flex flex-col flex-1 h-full overflow-hidden bg-background text-foreground animate-in fade-in duration-300">
            <header className="flex items-center justify-between px-6 py-4 shrink-0 bg-card shadow-sm">
              <div className="flex items-center gap-3">
                <button
                  onClick={() => navigateTo("hub")}
                  className="text-secondary-foreground hover:text-primary hover:bg-secondary font-semibold flex items-center gap-1.5 px-3 py-2 rounded-lg transition-colors"
                >
                  <span>←</span> 返回集市
                </button>
                <div>
                  <h1 className="text-xl font-bold tracking-tight">数据管理中心</h1>
                  <p className="text-xs text-secondary-foreground mt-0.5">项目根目录管理、文件健康度检测与路径批量重定位</p>
                </div>
              </div>
            </header>
            <WorkspaceGate onBack={() => navigateTo("hub")} backLabel="返回集市" />
          </div>
        )
      ) : (
        <div className="p-8">
          <button onClick={() => navigateTo("hub")} className="mb-4 font-bold text-primary">返回</button>
          <p>模块正在开发中...</p>
        </div>
      )}

      <AiFloatingLauncher currentView={currentView} />
    </div>
  );
}

function HubView({ onOpenTool }: { onOpenTool: (view: string) => void }) {
  return (
    <div className="relative flex min-h-0 flex-1 flex-col items-center overflow-y-auto px-6 py-8 animate-in fade-in duration-500 md:px-10">
      <div className="absolute left-6 top-5 flex items-center gap-2 font-bold text-foreground before:h-4 before:w-1 before:rounded-sm before:bg-primary before:content-[''] md:left-10">
        云数中心工具集
      </div>
      <div className="mb-10 mt-20 text-center md:mb-12">
        <h1 className="mb-2 text-4xl font-extrabold tracking-tight text-foreground">云数中心工具集</h1>
        <p className="font-medium text-secondary-foreground">请选择需要使用的工具模块</p>
      </div>
      <div className="grid w-full max-w-5xl grid-cols-[repeat(auto-fit,minmax(220px,1fr))] gap-6 pb-8">
        <HubCard
          icon="project"
          title="项目看板"
          description="先选择项目工作区，再管理项目生命周期"
          delay=""
          onClick={() => onOpenTool("project_board")}
        />
        <HubCard
          icon="cashflow"
          title="ICT项目全生命周期"
          description="测算、现金流推演与智能反算"
          delay="delay-75"
          onClick={() => onOpenTool("ict_lifecycle")}
        />
        <HubCard
          icon="settings"
          title="数据管理中心"
          description="配置根目录、重定位与健康自愈"
          delay="delay-150"
          onClick={() => onOpenTool("data_management")}
        />
      </div>
    </div>
  );
}

function HubCard({
  icon,
  title,
  description,
  delay,
  onClick,
}: {
  icon: AppIconName;
  title: string;
  description: string;
  delay: string;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      className={`flex min-h-[180px] cursor-pointer flex-col items-center justify-center rounded-2xl border border-border bg-card p-6 text-center shadow-sm transition-all hover:-translate-y-1 hover:border-primary/50 hover:shadow-lg animate-in slide-in-from-bottom duration-300 ${delay}`}
      onClick={onClick}
    >
      <div className="mb-5 flex h-16 w-16 items-center justify-center rounded-2xl bg-secondary text-primary transition-colors">
        <AppIcon name={icon} size={30} />
      </div>
      <div className="mb-1 text-lg font-bold">{title}</div>
      <div className="text-sm text-secondary-foreground">{description}</div>
    </button>
  );
}
