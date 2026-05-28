import { useState } from "react";
import { Clock3, Database, FolderOpen, X } from "lucide-react";
import { useWorkspaceStore } from "../../store/useWorkspaceStore";
import { useProjectStore } from "../../store/useProjectStore";
import { workspaceService, parseWorkspaceError } from "../../utils/workspaceService";

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
    error: storeError,
    isLoading: storeLoading,
    openRecentWorkspace,
    initializeWorkspaceFromExisting,
  } = useWorkspaceStore();

  // Local state for candidate project imports
  const [importModalOpen, setImportModalOpen] = useState(false);
  const [selectedPath, setSelectedPath] = useState("");
  const [workspaceName, setWorkspaceName] = useState("");
  const [candidates, setCandidates] = useState<string[]>([]);
  const [selectedCandidates, setSelectedCandidates] = useState<string[]>([]);
  const [createProjectJson, setCreateProjectJson] = useState(true);
  const [createSubDirs, setCreateSubDirs] = useState(true);

  const [localError, setLocalError] = useState<string | null>(null);
  const [localLoading, setLocalLoading] = useState(false);

  const error = localError || storeError;
  const isLoading = localLoading || storeLoading;

  const handleWorkspaceClick = async (path: string, isCurrent: boolean) => {
    if (isCurrent) {
      onCurrentWorkspaceSelected?.();
      return;
    }
    setLocalError(null);
    try {
      await openRecentWorkspace(path);
      onWorkspaceChanged?.();
    } catch (err) {
      setLocalError(parseWorkspaceError(err).message);
    }
  };

  const handleSelectFolder = async () => {
    setLocalError(null);
    try {
      const path = await workspaceService.selectFolder();
      if (!path) return;

      setLocalLoading(true);
      const inspect = await workspaceService.inspectPath(path);

      if (inspect.status === "workspace") {
        useProjectStore.getState().clearCurrentProject();
        await openRecentWorkspace(path);
        onWorkspaceChanged?.();
        setLocalLoading(false);
      } else if (inspect.status === "legacySuspected") {
        setLocalError(inspect.message || "疑似旧版数据目录，本阶段不会覆盖。");
        setLocalLoading(false);
      } else if (inspect.status === "importablePlainDirectory") {
        // Resolve a friendly workspace name
        const folderName = path.split(/[/\\]/).pop() || "Lamber Workspace";
        const parsedCandidates: string[] = JSON.parse(inspect.message || "[]");

        setSelectedPath(path);
        setWorkspaceName(folderName);
        setCandidates(parsedCandidates);
        setSelectedCandidates(parsedCandidates); // Default all selected
        setImportModalOpen(true);
        setLocalLoading(false);
      } else if (inspect.status === "nonEmptyNonWorkspace") {
        const confirmInit = confirm("该目录非空且不是 Lamber 工作区。确认要在此目录初始化工作区吗？");
        if (confirmInit) {
          useProjectStore.getState().clearCurrentProject();
          await workspaceService.create(path, undefined, true);
          await useWorkspaceStore.getState().refreshWorkspaceState();
          onWorkspaceChanged?.();
        }
        setLocalLoading(false);
      } else {
        // Empty or initializable
        useProjectStore.getState().clearCurrentProject();
        await workspaceService.create(path, undefined, false);
        await useWorkspaceStore.getState().refreshWorkspaceState();
        onWorkspaceChanged?.();
        setLocalLoading(false);
      }
    } catch (err) {
      setLocalError(parseWorkspaceError(err).message);
      setLocalLoading(false);
    }
  };

  const handleToggleCandidate = (candidate: string) => {
    setSelectedCandidates(prev =>
      prev.includes(candidate)
        ? prev.filter(c => c !== candidate)
        : [...prev, candidate]
    );
  };

  const handleToggleAllCandidates = () => {
    if (selectedCandidates.length === candidates.length) {
      setSelectedCandidates([]);
    } else {
      setSelectedCandidates([...candidates]);
    }
  };

  const handleConfirmImport = async (importProjects: boolean) => {
    setLocalLoading(true);
    setLocalError(null);
    try {
      const options = {
        workspaceName: workspaceName.trim() || undefined,
        selectedDirectories: importProjects ? selectedCandidates : [],
        createProjectJson,
        createSubDirs: createSubDirs,
      };

      await initializeWorkspaceFromExisting(selectedPath, options);
      setImportModalOpen(false);
      onWorkspaceChanged?.();
    } catch (err) {
      setLocalError(parseWorkspaceError(err).message);
    } finally {
      setLocalLoading(false);
    }
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
            {onBack && (
              <button
                type="button"
                onClick={onBack}
                className="rounded-md bg-muted px-4 py-2 text-xs font-bold text-foreground hover:bg-muted/80 transition-colors"
              >
                {backLabel}
              </button>
            )}
            <button
              type="button"
              disabled={isLoading}
              onClick={handleSelectFolder}
              className="inline-flex items-center gap-2 rounded-md bg-muted px-4 py-2 text-xs font-bold text-foreground hover:bg-muted/80 disabled:opacity-50 transition-colors"
            >
              <FolderOpen className="h-4 w-4" />
              打开目录 / 初始化工作区
            </button>
          </div>
        </div>

        {recentWorkspaces.length === 0 ? (
          <div className="flex flex-1 items-center justify-center rounded-lg bg-muted/40 p-10 text-center">
            <div>
              <Database className="mx-auto mb-3 h-9 w-9 text-secondary-foreground" />
              <div className="text-base font-bold text-foreground">暂无已记录的工作区</div>
              <div className="mt-1 text-sm text-secondary-foreground">请选择目录以打开或新建 Lamber 工作区。</div>
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

      {/* Candidate Subdirectories Import Modal Dialog */}
      {importModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-background/80 backdrop-blur-sm">
          <div className="relative w-full max-w-lg rounded-lg bg-card p-6 shadow-xl flex flex-col gap-5 border border-muted">
            <button
              type="button"
              onClick={() => setImportModalOpen(false)}
              className="absolute right-4 top-4 text-secondary-foreground hover:text-foreground transition-colors"
            >
              <X className="h-4 w-4" />
            </button>

            <div>
              <h3 className="text-lg font-extrabold text-foreground">初始化并导入工作区</h3>
              <p className="mt-1 text-xs text-secondary-foreground">
                检测到该目录包含以下子文件夹。你可以选择将它们导入为项目，或仅初始化为空工作区。
              </p>
            </div>

            <div className="flex flex-col gap-1.5">
              <label htmlFor="modal-workspace-name-input" className="text-xs font-extrabold text-secondary-foreground">
                工作区名称
              </label>
              <input
                id="modal-workspace-name-input"
                type="text"
                value={workspaceName}
                onChange={e => setWorkspaceName(e.target.value)}
                className="w-full rounded-md bg-muted px-3 py-2 text-xs font-semibold text-foreground focus:outline-none focus:ring-1 focus:ring-primary"
                placeholder="工作区名称"
              />
            </div>

            <div className="flex flex-col gap-2">
              <div className="flex items-center justify-between">
                <span className="text-xs font-extrabold text-secondary-foreground">可选项目目录 ({selectedCandidates.length}/{candidates.length})</span>
                <button
                  type="button"
                  onClick={handleToggleAllCandidates}
                  className="text-xs text-primary font-bold hover:underline"
                >
                  {selectedCandidates.length === candidates.length ? "取消全选" : "全选"}
                </button>
              </div>
              <div className="max-h-48 overflow-y-auto rounded-md bg-muted/60 p-2 flex flex-col gap-1">
                {candidates.map(c => (
                  <label
                    key={c}
                    className="flex items-center gap-2 rounded px-2 py-1.5 hover:bg-card transition-colors cursor-pointer select-none"
                  >
                    <input
                      type="checkbox"
                      checked={selectedCandidates.includes(c)}
                      onChange={() => handleToggleCandidate(c)}
                      className="rounded border-muted text-primary focus:ring-primary h-3.5 w-3.5 bg-muted"
                    />
                    <span className="text-xs font-semibold text-foreground">{c}</span>
                  </label>
                ))}
              </div>
            </div>

            <div className="flex flex-col gap-2 rounded bg-muted/30 p-3">
              <label className="flex items-center gap-2 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={createProjectJson}
                  onChange={e => setCreateProjectJson(e.target.checked)}
                  className="rounded border-muted text-primary focus:ring-primary h-3.5 w-3.5 bg-muted"
                />
                <span className="text-xs font-bold text-foreground">自动创建或补全 project.json 配置文件</span>
              </label>
              <label className="flex items-center gap-2 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={createSubDirs}
                  onChange={e => setCreateSubDirs(e.target.checked)}
                  className="rounded border-muted text-primary focus:ring-primary h-3.5 w-3.5 bg-muted"
                />
                <span className="text-xs font-bold text-foreground">自动创建 assets/documents/analyses 目录</span>
              </label>
            </div>

            <div className="flex justify-end gap-2 mt-2">
              <button
                type="button"
                disabled={isLoading}
                onClick={() => setImportModalOpen(false)}
                className="rounded-md bg-muted px-4 py-2 text-xs font-bold text-foreground hover:bg-muted/80 transition-colors disabled:opacity-50"
              >
                取消
              </button>
              <button
                type="button"
                disabled={isLoading}
                onClick={() => handleConfirmImport(false)}
                className="rounded-md bg-muted px-4 py-2 text-xs font-bold text-foreground hover:bg-muted/80 transition-colors disabled:opacity-50"
              >
                仅初始化为空工作区
              </button>
              <button
                type="button"
                disabled={isLoading || selectedCandidates.length === 0}
                onClick={() => handleConfirmImport(true)}
                className="rounded-md bg-primary px-4 py-2 text-xs font-bold text-primary-foreground hover:bg-primary/90 transition-colors disabled:opacity-50"
              >
                初始化并导入项目
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
