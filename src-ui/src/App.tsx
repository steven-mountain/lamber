import { listenBenefitSimulations } from "./services/benefitSimulation";
import { assertStructureBinding, finishStructureRequest, listenStructureRequests, useStructureRequest, type ReverseRequest } from './services/chatStructureReverse';
import { getCurrentWindow } from '@tauri-apps/api/window';
import { assertDocumentBinding, finishDocumentRequest, useDocumentRequest, listenDocumentRequests } from './services/chatDocumentGeneration';
import { useLatestCallback } from './hooks/useLatestCallback';
import { useEffect } from "react";
import { emitTo, listen } from "@tauri-apps/api/event";
import IctLifecycle from "./views/IctLifecycle";
import ProjectBoard from "./views/ProjectBoard";
import DataManagement from "./views/DataManagement";
import PresetCenterView from "./views/PresetCenterView";
import AiComputeQuoteView from "./features/ai-compute-quote/AiComputeQuoteView";
import SettingsView from "./components/settings/SettingsView";
import AgentApprovalDialog from "./components/ai/AgentApprovalDialog";
import AgentLabView from "./components/ai/AgentLabView";
import AiFloatingLauncher from "./components/ai/AiFloatingLauncher";
import AiFloatingWindow from "./components/ai/AiFloatingWindow";
import AppIcon, { type AppIconName } from "./components/icons/AppIcon";
import WorkspaceGate from "./components/workspace/WorkspaceGate";
import { useAppearanceStore } from "./store/useAppearanceStore";
import { useAiContextStore } from "./store/useAiContextStore";
import { useNavigationStore } from "./store/useNavigationStore";
import { useWorkspaceStore } from "./store/useWorkspaceStore";
import { useGlobalSaveShortcut } from "./hooks/useGlobalSaveShortcut";
import { useUnsavedChangesGuard } from "./hooks/useUnsavedChangesGuard";

const AI_ASSISTANT_LABEL = "ai-assistant";
const AI_CURRENT_VIEW_KEY = "lamber_ai_current_view";
const WORKSPACE_STATE_CHANGED_EVENT = "lamber-workspace-state-changed";

function isTauriRuntime() {
  return typeof window !== "undefined" && Boolean((window as Window & { __TAURI_INTERNALS__?: unknown }).__TAURI_INTERNALS__);
}

/** `#/agent-lab` opens the dsh agent bench; see `AgentLabView`. */
function isAgentLabRoute() {
  return window.location.hash.startsWith("#/agent-lab");
}

function getAiAssistantView() {
  const hash = window.location.hash;
  if (!hash.startsWith("#/ai-assistant")) return null;

  const query = hash.includes("?") ? hash.slice(hash.indexOf("?") + 1) : "";
  const view = new URLSearchParams(query).get("view");
  return view || "hub";
}

export default function App() {
  useEffect(() => {
    if (!isTauriRuntime() || getCurrentWindow().label !== "main") return;
    let disposed = false;
    let unlisten: (() => void) | undefined;
    void listenBenefitSimulations().then(stop => { if (disposed) stop(); else unlisten = stop; });
    return () => { disposed = true; unlisten?.(); };
  }, []);
  const { currentView, settingsReturnView, navigateTo } = useNavigationStore();
  const aiLauncherVisible = useAppearanceStore(state => state.settings.aiLauncherVisible);
  const setActiveModule = useAiContextStore(state => state.setActiveModule);
  const { isWorkspaceReady, refreshWorkspaceState } = useWorkspaceStore();
  const aiAssistantView = getAiAssistantView();
  const pendingDocument = useDocumentRequest(value => value.request);
  const pendingDocumentPhase = useDocumentRequest(value => value.phase);
  const pendingStructure = useStructureRequest(value => value.request);
  const pendingStructurePhase = useStructureRequest(value => value.phase);
  const activeDocumentProjectId = useNavigationStore(value => value.activeProjectId);
  const documentWorkspaceId = useWorkspaceStore(value => value.workspaceId);
  const agentLab = isAgentLabRoute();
  useGlobalSaveShortcut();
  const { confirmOrSave } = useUnsavedChangesGuard();
  const openChatDocument = useLatestCallback(async (request: import('./services/chatDocumentGeneration').DocumentRequest) => {
    await getCurrentWindow().unminimize();
    await getCurrentWindow().show();
    await getCurrentWindow().setFocus();
    const navigation = useNavigationStore.getState();
    if (navigation.currentView !== 'ict_lifecycle' || navigation.activeProjectId !== request.projectId) {
      if (!await confirmOrSave()) throw new Error('已取消切换项目，未生成文档。');
      await assertDocumentBinding(request);
      navigateTo('ict_lifecycle', request.projectId);
    }
  });
  const openStructureEditor = useLatestCallback(async (request: ReverseRequest) => {
    if (useDocumentRequest.getState().request) throw new Error('主窗口正在处理文档操作，请完成后再反算。');
    const navigation = useNavigationStore.getState();
    const same = navigation.currentView === 'ict_lifecycle' && navigation.activeProjectId === request.projectId && navigation.activeSchemeId === request.schemeId;
    if (request.action === 'apply' && !same) throw new Error('项目、方案或页面已变化，请重新选择科目并读取范围。');
    if (!same) {
      await getCurrentWindow().unminimize();
      await getCurrentWindow().show();
      if (!await confirmOrSave()) throw new Error('已取消切换项目，未执行结构反算。');
      await assertStructureBinding(request);
      navigateTo('ict_lifecycle', request.projectId, request.schemeId);
    }
  });
  useEffect(() => {
    if (aiAssistantView || agentLab || !isTauriRuntime()) return;
    let disposed = false;
    let stop: (() => void) | undefined;
    void listenStructureRequests(openStructureEditor).then(unlisten => { if (disposed) unlisten(); else stop = unlisten; }).catch(console.error);
    return () => { disposed = true; stop?.(); };
  }, [aiAssistantView, agentLab, openStructureEditor]);
  useEffect(() => {
    if (aiAssistantView || agentLab || !isTauriRuntime()) return;
    let disposed = false;
    let stop: (() => void) | undefined;
    void listenDocumentRequests(openChatDocument).then(unlisten => { if (disposed) unlisten(); else stop = unlisten; }).catch(console.error);
    return () => { disposed = true; stop?.(); };
  }, [aiAssistantView, agentLab, openChatDocument]);

  useEffect(() => {
    if (aiAssistantView || agentLab || !pendingDocument || (pendingDocumentPhase === 'loading' || pendingDocumentPhase === 'running')) return;
    if (currentView !== 'ict_lifecycle' || activeDocumentProjectId !== pendingDocument.projectId || documentWorkspaceId !== pendingDocument.workspaceId) {
      void finishDocumentRequest(pendingDocument, { status: 'cancelled', message: '项目、工作区或页面已切换；已停止尚未执行的生成。' }).catch(console.error);
    }
  }, [aiAssistantView, agentLab, pendingDocument, pendingDocumentPhase, currentView, activeDocumentProjectId, documentWorkspaceId]);
  useEffect(() => {
    if (aiAssistantView || agentLab || !pendingStructure || pendingStructurePhase !== 'opening') return;
    if (currentView !== 'ict_lifecycle' || activeDocumentProjectId !== pendingStructure.projectId || documentWorkspaceId !== pendingStructure.workspaceId) {
      void finishStructureRequest(pendingStructure, { status: 'error', message: '项目、工作区或页面已切换，已停止尚未执行的结构反算。' }).catch(console.error);
    }
  }, [aiAssistantView, agentLab, pendingStructure, pendingStructurePhase, currentView, activeDocumentProjectId, documentWorkspaceId]);

  useEffect(() => {
    if (aiAssistantView || agentLab) return;
    if (!isTauriRuntime()) return;

    let cancelled = false;
    let unlisten: (() => void) | null = null;

    listen(WORKSPACE_STATE_CHANGED_EVENT, () => {
      if (!cancelled) {
        refreshWorkspaceState();
      }
    }).then((handler) => {
      if (cancelled) {
        handler();
      } else {
        unlisten = handler;
      }
    }).catch(error => console.warn("Failed to listen for workspace state changes:", error));

    refreshWorkspaceState();

    return () => {
      cancelled = true;
      unlisten?.();
    };
  }, [agentLab, aiAssistantView, refreshWorkspaceState]);

  useEffect(() => {
    if (aiAssistantView || agentLab) return;

    localStorage.setItem(AI_CURRENT_VIEW_KEY, currentView);
    if (currentView === "hub") {
      setActiveModule("hub");
    } else if (currentView === "project_board") {
      setActiveModule("project_board.core");
    } else if (currentView === "ict_lifecycle") {
      setActiveModule("ict");
    } else if (currentView === "ai_compute_quote") {
      setActiveModule("ai_compute_quote");
    } else {
      setActiveModule(currentView);
    }

    if (isTauriRuntime()) {
      emitTo(AI_ASSISTANT_LABEL, "lamber-ai-view-changed", { view: currentView })
         .catch(error => console.warn("Failed to sync AI assistant view:", error));
    }
  }, [agentLab, aiAssistantView, currentView, setActiveModule]);

  // The bench needs the approval dialog alongside it: the backend parks a tool
  // call waiting for that answer.
  if (agentLab) {
    return (
      <>
        <AgentLabView />
        <AgentApprovalDialog />
      </>
    );
  }

  if (aiAssistantView) {
    return (
      <>
        <AiFloatingWindow currentView={aiAssistantView} />
        <AgentApprovalDialog />
      </>
    );
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
      ) : currentView === "ai_compute_quote" ? (
        isWorkspaceReady ? (
          <AiComputeQuoteView />
        ) : (
          <div className="flex flex-col flex-1 h-full overflow-hidden bg-background text-foreground animate-in fade-in duration-300">
            <header className="flex items-center justify-between px-6 py-4 shrink-0 bg-card shadow-sm">
              <div className="flex items-center gap-3">
                <button onClick={() => navigateTo("hub")} className="text-secondary-foreground hover:text-primary hover:bg-secondary font-semibold flex items-center gap-1.5 px-3 py-2 rounded-lg transition-colors text-body">
                  <span>←</span> 返回集市
                </button>
                <div>
                  <h1 className="text-page-title font-bold tracking-tight">智算测算</h1>
                  <p className="text-caption text-secondary-foreground mt-0.5">请先打开工作区并从项目看板进入智算项目</p>
                </div>
              </div>
            </header>
            <WorkspaceGate onBack={() => navigateTo("hub")} backLabel="返回集市" />
          </div>
        )
      ) : currentView === "settings" ? (
        <SettingsView onBack={() => navigateTo(settingsReturnView || "hub")} />
      ) : currentView === "preset_center" ? (
        isWorkspaceReady ? (
          <PresetCenterView onBack={() => navigateTo("hub")} />
        ) : (
          <div className="flex flex-col flex-1 h-full overflow-hidden bg-background text-foreground animate-in fade-in duration-300">
            <header className="flex items-center justify-between px-6 py-4 shrink-0 bg-card shadow-sm">
              <div className="flex items-center gap-3">
                <button
                  onClick={() => navigateTo("hub")}
                  className="text-secondary-foreground hover:text-primary hover:bg-secondary font-semibold flex items-center gap-1.5 px-3 py-2 rounded-lg transition-colors text-body"
                >
                  <span>←</span> 返回集市
                </button>
                <div>
                  <h1 className="text-page-title font-bold tracking-tight">常用资料与项目预设</h1>
                  <p className="text-caption text-secondary-foreground mt-0.5">请先打开工作区后管理可复用资料</p>
                </div>
              </div>
            </header>
            <WorkspaceGate onBack={() => navigateTo("hub")} backLabel="返回集市" />
          </div>
        )
      ) : currentView === "data_management" ? (
        isWorkspaceReady ? (
          <DataManagement onBack={() => navigateTo("hub")} />
        ) : (
          <div className="flex flex-col flex-1 h-full overflow-hidden bg-background text-foreground animate-in fade-in duration-300">
            <header className="flex items-center justify-between px-6 py-4 shrink-0 bg-card shadow-sm">
              <div className="flex items-center gap-3">
                <button
                  onClick={() => navigateTo("hub")}
                  className="text-secondary-foreground hover:text-primary hover:bg-secondary font-semibold flex items-center gap-1.5 px-3 py-2 rounded-lg transition-colors text-body"
                >
                  <span>←</span> 返回集市
                </button>
                <div>
                  <h1 className="text-page-title font-bold tracking-tight">数据管理中心</h1>
                  <p className="text-caption text-secondary-foreground mt-0.5">项目根目录管理、文件健康度检测与路径批量重定位</p>
                </div>
              </div>
            </header>
            <WorkspaceGate onBack={() => navigateTo("hub")} backLabel="返回集市" />
          </div>
        )
      ) : (
        <div className="p-8 text-body">
          <button onClick={() => navigateTo("hub")} className="mb-4 font-bold text-primary">返回</button>
          <p>模块正在开发中...</p>
        </div>
      )}

      {aiLauncherVisible && <AiFloatingLauncher currentView={currentView} />}
      {/* Mounted at the root: the backend parks a tool call awaiting this
          answer, so the listener must outlive any individual panel. */}
      <AgentApprovalDialog />
    </div>
  );
}

function HubView({ onOpenTool }: { onOpenTool: (view: string) => void }) {
  return (
    <div className="relative flex min-h-0 flex-1 flex-col items-center overflow-y-auto px-6 py-8 animate-in fade-in duration-500 md:px-10">
      <div className="absolute left-6 top-5 flex items-center gap-2 text-body-strong text-foreground before:h-4 before:w-1 before:rounded-sm before:bg-primary before:content-[''] md:left-10">
        云数中心工具集
      </div>
      <div className="absolute right-6 top-4 md:right-10 flex items-center gap-3">
        <button
          onClick={() => onOpenTool("settings")}
          className="flex h-9 items-center gap-2 rounded-lg border border-border bg-card px-3 text-sm font-semibold text-secondary-foreground hover:bg-secondary hover:text-foreground transition-all shadow-sm"
        >
          <AppIcon name="settings" size={16} />
          系统设置
        </button>
      </div>
      <div className="mb-10 mt-20 text-center md:mb-12">
        <h1 className="mb-2 text-display font-extrabold tracking-tight text-foreground">云数中心工具集</h1>
        <p className="text-body font-medium text-secondary-foreground">请选择需要使用的工具模块</p>
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
          icon="calculator"
          title="智算测算"
          description="从项目看板维护智算金额来源并同步到 ICT"
          delay="delay-100"
          onClick={() => onOpenTool("project_board")}
        />
        <HubCard
          icon="settings"
          title="数据管理中心"
          description="配置根目录、重定位与健康自愈"
          delay="delay-150"
          onClick={() => onOpenTool("data_management")}
        />
        <HubCard
          icon="presets"
          title="常用资料与项目预设"
          description="管理可复用字段、文本片段与表单快捷填充"
          delay="delay-200"
          onClick={() => onOpenTool("preset_center")}
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
      <div className="mb-1 text-section-title font-bold">{title}</div>
      <div className="text-body text-secondary-foreground">{description}</div>
    </button>
  );
}
