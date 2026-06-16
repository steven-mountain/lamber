import { useState, type ReactNode } from "react";
import { useModulePath } from "../hooks/useModulePath";
import AppIcon from "./icons/AppIcon";
import GlobalSaveButton from "./GlobalSaveButton";
import { useNavigationStore } from "../store/useNavigationStore";
import { useWorkspaceStore } from "../store/useWorkspaceStore";

interface WorkspaceHeaderProps {
  moduleId: string;
  title: string;
  onBack: () => void;
  backLabel?: string;
  onPathChange?: (newPath: string) => void;
  contextContent?: ReactNode;
  inlineWorkspaceControl?: boolean;
}

export default function WorkspaceHeader({
  moduleId,
  title,
  onBack,
  backLabel = "返回集市",
  onPathChange,
  contextContent,
  inlineWorkspaceControl = false,
}: WorkspaceHeaderProps) {
  const { path, isLoading, updatePath } = useModulePath(moduleId);
  const workspaceName = useWorkspaceStore(state => state.workspaceName);
  const [isWorkspaceMenuOpen, setIsWorkspaceMenuOpen] = useState(false);

  const handleUpdate = async () => {
    const newPath = await updatePath();
    setIsWorkspaceMenuOpen(false);
    if (newPath && onPathChange) {
      onPathChange(newPath);
    }
  };

  const handleOpenPath = () => {
    if (!path) return;
    window.open(`file://${path}`);
    setIsWorkspaceMenuOpen(false);
  };

  return (
    <div className="flex flex-col w-full shrink-0">
      <div className="h-16 bg-card border-b border-border flex items-center px-6 gap-4">
        <button 
          onClick={onBack} 
          className="text-secondary-foreground hover:text-primary hover:bg-secondary font-semibold flex items-center gap-1.5 px-3 py-2 rounded-lg transition-colors"
        >
          <span>←</span> {backLabel}
        </button>
        <h2 className="m-0 text-lg font-bold text-foreground border-l-2 border-border pl-4">{title}</h2>
        {contextContent && (
          <div className="min-w-0 flex-1 px-4">
            {contextContent}
          </div>
        )}
        <div className="ml-auto flex items-center gap-2">
          {inlineWorkspaceControl && (
            <div className="relative">
              <button
                type="button"
                onClick={() => setIsWorkspaceMenuOpen(current => !current)}
                className={`flex h-9 max-w-[280px] items-center gap-2 rounded-lg px-3 text-xs font-bold shadow-sm transition-all ${
                  path
                    ? "bg-secondary/70 text-foreground hover:bg-secondary"
                    : "bg-destructive-soft text-destructive hover:bg-destructive-soft/80"
                }`}
                title={path || "未设置工作目录"}
              >
                <AppIcon name={path ? "folder" : "warning"} size={15} />
                <span className="min-w-0 truncate">
                  当前工作区：{isLoading ? "加载中..." : workspaceName || "未设置"}
                </span>
                <AppIcon name={isWorkspaceMenuOpen ? "chevronUp" : "chevronDown"} size={14} />
              </button>

              {isWorkspaceMenuOpen && (
                <div className="absolute right-0 top-11 z-40 w-[min(24rem,calc(100vw-2rem))] rounded-xl bg-popover p-3 text-popover-foreground shadow-lg ring-1 ring-border/60">
                  <div className="rounded-lg bg-muted/50 px-3 py-2">
                    <div className="text-[10px] font-extrabold uppercase tracking-wider text-secondary-foreground opacity-70">
                      当前模块工作空间
                    </div>
                    <div className={`mt-1 break-all text-xs font-semibold leading-5 ${path ? "text-foreground" : "text-destructive"}`}>
                      {isLoading ? "加载中..." : path || "未设置工作目录"}
                    </div>
                  </div>
                  <div className="mt-3 flex justify-end gap-2">
                    {path && (
                      <button
                        type="button"
                        onClick={handleOpenPath}
                        className="rounded-lg bg-card px-3 py-2 text-xs font-bold text-foreground shadow-sm hover:bg-secondary transition-colors"
                      >
                        打开目录
                      </button>
                    )}
                    <button
                      type="button"
                      onClick={handleUpdate}
                      className="rounded-lg bg-primary px-3 py-2 text-xs font-bold text-primary-foreground shadow-sm hover:bg-primary/90 transition-colors"
                    >
                      {path ? "修改目录" : "立即设置"}
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}
          <GlobalSaveButton />
          <button
            onClick={() => useNavigationStore.getState().navigateTo("settings")}
            className="flex h-9 w-9 items-center justify-center rounded-lg border border-border bg-card text-secondary-foreground hover:bg-secondary hover:text-foreground transition-all shadow-sm"
            title="系统设置"
          >
            <AppIcon name="settings" size={18} />
          </button>
        </div>
      </div>

      {!inlineWorkspaceControl && (
        <div className="px-8 py-4 bg-secondary/30 border-b border-border flex items-center justify-between animate-in slide-in-from-top duration-300">
          <div className="flex flex-col gap-1">
            <div className="text-[10px] uppercase tracking-wider font-extrabold text-secondary-foreground opacity-60">
              当前模块工作空间
            </div>
            <div className={`text-sm font-mono flex items-center gap-2 ${path ? 'text-primary' : 'text-destructive font-bold'}`}>
              <AppIcon name={path ? "folder" : "warning"} size={16} />
              {isLoading ? '加载中...' : path || '未设置工作目录 (点击右侧按钮进行配置)'}
            </div>
          </div>

          <div className="flex gap-3">
            {path && (
              <button
                onClick={() => window.open(`file://${path}`)}
                className="px-4 py-2 bg-card border border-border text-xs font-bold rounded-lg hover:bg-secondary transition-all shadow-sm"
              >
                打开目录
              </button>
            )}
            <button
              onClick={handleUpdate}
              className="px-4 py-2 bg-primary text-primary-foreground text-xs font-bold rounded-lg hover:shadow-lg hover:-translate-y-0.5 transition-all active:translate-y-0"
            >
              {path ? '修改目录' : '立即设置'}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
