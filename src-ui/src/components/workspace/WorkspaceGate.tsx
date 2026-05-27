import { Clock3, Database, FolderOpen, Plus } from "lucide-react";
import { useWorkspaceStore } from "../../store/useWorkspaceStore";

interface WorkspaceGateProps {
  compact?: boolean;
  onBack?: () => void;
  backLabel?: string;
  onCurrentWorkspaceSelected?: () => void;
  onWorkspaceChanged?: () => void;
}

export default function WorkspaceGate({
  compact = false,
  onBack,
  backLabel = "返回项目列表",
  onCurrentWorkspaceSelected,
  onWorkspaceChanged,
}: WorkspaceGateProps) {
  const {
    currentWorkspace,
    recentWorkspaces,
    error,
    isLoading,
    selectAndOpenWorkspace,
    selectAndCreateWorkspace,
    openRecentWorkspace,
  } = useWorkspaceStore();

  const handleWorkspaceClick = async (path: string, isCurrent: boolean) => {
    if (isCurrent) {
      onCurrentWorkspaceSelected?.();
      return;
    }
    await openRecentWorkspace(path);
    onWorkspaceChanged?.();
  };

  return (
    <div className={`flex flex-1 flex-col bg-background px-6 ${compact ? "py-6" : "py-8"}`}>
      <div className="mx-auto flex w-full max-w-6xl flex-1 flex-col gap-6">
        <div className="flex flex-col gap-4 rounded-lg bg-card p-5 shadow-sm md:flex-row md:items-center md:justify-between">
          <div className="min-w-0">
            <div className="text-[10px] font-extrabold uppercase tracking-wide text-secondary-foreground">Project Workspace</div>
            <h2 className="mt-1 text-2xl font-extrabold text-foreground">项目工作区</h2>
            <p className="mt-1 text-sm text-secondary-foreground">
              选择一个工作区后进入对应的项目看板。工作区名称默认使用根目录名称。
            </p>
            {error && <div className="mt-3 rounded-md bg-destructive/10 px-3 py-2 text-sm font-semibold text-destructive">{error}</div>}
          </div>

          <div className="flex flex-wrap gap-2">
            {onBack && currentWorkspace && (
              <button
                type="button"
                onClick={onBack}
                className="rounded-md bg-muted px-4 py-2 text-xs font-bold text-foreground"
              >
                {backLabel}
              </button>
            )}
            <button
              type="button"
              disabled={isLoading}
              onClick={selectAndOpenWorkspace}
              className="inline-flex items-center gap-2 rounded-md bg-muted px-4 py-2 text-xs font-bold text-foreground disabled:opacity-50"
            >
              <FolderOpen className="h-4 w-4" />
              打开其他目录
            </button>
            <button
              type="button"
              disabled={isLoading}
              onClick={selectAndCreateWorkspace}
              className="inline-flex items-center gap-2 rounded-md bg-primary px-4 py-2 text-xs font-bold text-primary-foreground disabled:opacity-50"
            >
              <Plus className="h-4 w-4" />
              新建工作区
            </button>
          </div>
        </div>

        {recentWorkspaces.length === 0 ? (
          <div className="flex flex-1 items-center justify-center rounded-lg bg-muted/40 p-10 text-center">
            <div>
              <Database className="mx-auto mb-3 h-9 w-9 text-secondary-foreground" />
              <div className="text-base font-bold text-foreground">暂无已记录的工作区</div>
              <div className="mt-1 text-sm text-secondary-foreground">请新建工作区，或打开一个已有 Lamber 工作区。</div>
            </div>
          </div>
        ) : (
          <div className="grid grid-cols-1 gap-4 md:grid-cols-2 xl:grid-cols-3">
            {recentWorkspaces.map(item => {
              const isCurrent = currentWorkspace?.workspaceRoot === item.path;
              return (
                <button
                  key={item.path}
                  type="button"
                  disabled={isLoading && !isCurrent}
                  onClick={(event) => {
                    event.preventDefault();
                    void handleWorkspaceClick(item.path, isCurrent);
                  }}
                  className={`min-h-[170px] rounded-lg bg-card p-5 text-left shadow-sm transition-all hover:-translate-y-0.5 hover:shadow-md disabled:opacity-60 ${
                    isCurrent ? "ring-2 ring-primary/30" : ""
                  }`}
                >
                  <div className="flex items-start justify-between gap-3">
                    <div className="flex min-w-0 items-center gap-3">
                      <span className={`flex h-11 w-11 shrink-0 items-center justify-center rounded-lg ${isCurrent ? "bg-primary text-primary-foreground" : "bg-muted text-secondary-foreground"}`}>
                        <Database className="h-5 w-5" />
                      </span>
                      <div className="min-w-0">
                        <div className="truncate text-base font-extrabold text-foreground">{item.name}</div>
                        <div className="mt-1 truncate font-mono text-[11px] text-secondary-foreground">{item.path}</div>
                      </div>
                    </div>
                    {isCurrent && (
                      <span className="shrink-0 rounded-md bg-primary/10 px-2 py-1 text-[10px] font-extrabold text-primary">当前</span>
                    )}
                  </div>
                  <div className="mt-6 flex items-center gap-2 text-xs font-semibold text-secondary-foreground">
                    <Clock3 className="h-3.5 w-3.5" />
                    最近打开 {new Date(item.lastOpenedAt).toLocaleString()}
                  </div>
                </button>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}
