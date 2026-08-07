import { useCallback, useState, useEffect, type KeyboardEvent } from "react"
import { invoke } from "@tauri-apps/api/core"
import { Clock3, Database } from "lucide-react"
import AppIcon from "../components/icons/AppIcon"
import { useWorkspaceStore } from "../store/useWorkspaceStore"
import { useSaveStore } from "../store/useSaveStore"
import { useNavigationStore } from "../store/useNavigationStore"
import { useUnsavedChangesGuard } from "../hooks/useUnsavedChangesGuard"
import { workspaceMaintenanceService, type ExternalPathInfo, type HealthCheckResult as WorkspaceHealthCheckResult, type WorkspaceBackupInfo } from "../services/workspaceMaintenanceService"

interface ProjectRoot {
  id: string
  name: string
  rootPath: string
  rootAlias?: string
  isDefault: boolean
  createdAt: string
  updatedAt: string
}

interface HealthReport {
  totalFiles: number
  healthyFiles: number
  missingFiles: number
  recoverableFiles: number
  details: {
    fileId: string
    projectId: string
    fileName: string
    currentPath: string
    status: "healthy" | "recoverable" | "missing"
    recoveredPath?: string
  }[]
}

interface RelocationPreview {
  totalItems: number
  matchedItems: number
  missingItems: number
  details: {
    itemId: string
    itemType: "file" | "directory"
    name: string
    oldPath: string
    newPath: string
    exists: boolean
  }[]
}

export default function DataManagement({ onBack }: { onBack: () => void }) {
  const [activeTab, setActiveTab] = useState<"workspaces" | "workspace" | "roots" | "health">("workspaces")
  const currentWorkspace = useWorkspaceStore(state => state.currentWorkspace)
  const recentWorkspaces = useWorkspaceStore(state => state.recentWorkspaces)
  const refreshWorkspaceState = useWorkspaceStore(state => state.refreshWorkspaceState)
  const openRecentWorkspace = useWorkspaceStore(state => state.openRecentWorkspace)
  const forgetWorkspace = useWorkspaceStore(state => state.forgetWorkspace)
  const clearAllDirty = useSaveStore(state => state.clearAllDirty)
  const navigateTo = useNavigationStore(state => state.navigateTo)
  const { confirmOrSave } = useUnsavedChangesGuard()
  
  // Roots state
  const [roots, setRoots] = useState<ProjectRoot[]>([])
  const [isAddingRoot, setIsAddingRoot] = useState(false)
  const [newRootName, setNewRootName] = useState("")
  const [newRootPath, setNewRootPath] = useState("")
  const [newRootAlias, setNewRootAlias] = useState("")
  const [newRootIsDefault, setNewRootIsDefault] = useState(false)
  const [editingRootId, setEditingRootId] = useState<string | null>(null)
  const [editingName, setEditingName] = useState("")
  const [editingAlias, setEditingAlias] = useState("")

  // Health state
  const [healthReport, setHealthReport] = useState<HealthReport | null>(null)
  const [isCheckingHealth, setIsCheckingHealth] = useState(false)

  // Relocation state
  const [relocateRootId, setRelocateRootId] = useState("")
  const [relocateNewPath, setRelocateNewPath] = useState("")
  const [relocationPreview, setRelocationPreview] = useState<RelocationPreview | null>(null)
  const [isPreviewingRelocation, setIsPreviewingRelocation] = useState(false)
  const [isRelocating, setIsRelocating] = useState(false)

  const [notification, setNotification] = useState<{ message: string; type: "success" | "error" } | null>(null)
  const [workspaceBackups, setWorkspaceBackups] = useState<WorkspaceBackupInfo[]>([])
  const [workspaceHealth, setWorkspaceHealth] = useState<WorkspaceHealthCheckResult | null>(null)
  const [externalPaths, setExternalPaths] = useState<ExternalPathInfo[]>([])
  const [workspaceBusy, setWorkspaceBusy] = useState(false)
  const [selectedManagedWorkspacePath, setSelectedManagedWorkspacePath] = useState<string | null>(null)

  const showNotification = useCallback((message: string, type: "success" | "error" = "success") => {
    setNotification({ message, type })
    setTimeout(() => setNotification(null), 4000)
  }, [])

  const getParentDirectory = (filePath: string) => {
    const trimmed = filePath.replace(/[\\/]+$/, "")
    const index = Math.max(trimmed.lastIndexOf("\\"), trimmed.lastIndexOf("/"))
    if (index <= 0) return filePath
    return trimmed.slice(0, index)
  }

  const getWorkspacePathKey = (path?: string | null) => {
    return (path || "").trim().replace(/[\\/]+$/, "").replace(/\\/g, "/").toLowerCase()
  }

  const isSameWorkspacePath = (left?: string | null, right?: string | null) => {
    return Boolean(left && right && getWorkspacePathKey(left) === getWorkspacePathKey(right))
  }

  const formatWorkspaceTime = (value?: string | null) => {
    if (!value) return "未知"
    const date = new Date(value)
    if (Number.isNaN(date.getTime())) return value
    return date.toLocaleString()
  }

  const ensureSavedForWorkspaceOperation = async (label: string) => {
    const saveStore = useSaveStore.getState()
    if (!saveStore.hasUnsavedChanges()) return true
    const shouldSave = confirm(`${label} 前需要先保存当前项目修改。是否保存并继续？`)
    if (!shouldSave) return false
    try {
      await saveStore.saveCurrentProject()
      return true
    } catch (err: any) {
      showNotification(`保存失败，已取消操作：${err?.message || err}`, "error")
      return false
    }
  }

  const refreshWorkspaceMaintenance = useCallback(async () => {
    if (!useWorkspaceStore.getState().currentWorkspace) {
      setWorkspaceBackups([])
      setExternalPaths([])
      setWorkspaceHealth(null)
      return
    }
    try {
      const [backups, external] = await Promise.all([
        workspaceMaintenanceService.listBackups(),
        workspaceMaintenanceService.listExternalPaths(),
      ])
      setWorkspaceBackups(backups)
      setExternalPaths(external)
    } catch (err: any) {
      showNotification(`加载工作区维护信息失败: ${err?.message || err}`, "error")
    }
  }, [showNotification])

  const handleCreateWorkspaceBackup = async () => {
    if (!(await ensureSavedForWorkspaceOperation("手动备份"))) return
    setWorkspaceBusy(true)
    try {
      const backup = await workspaceMaintenanceService.createBackup()
      showNotification(`备份完成: ${backup.path}`)
      await refreshWorkspaceMaintenance()
    } catch (err: any) {
      showNotification(`备份失败: ${err?.message || err}`, "error")
    } finally {
      setWorkspaceBusy(false)
    }
  }

  const handleExportWorkspace = async () => {
    if (!(await ensureSavedForWorkspaceOperation("导出 Workspace"))) return
    setWorkspaceBusy(true)
    try {
      const result = await workspaceMaintenanceService.exportWorkspace(null, {
        includeBackups: false,
        includeExports: false,
        allowWarnings: false,
      })
      showNotification(`导出完成: ${result.archivePath}`)
      await refreshWorkspaceMaintenance()
      if (confirm("导出完成，是否打开所在文件夹？")) {
        await workspaceMaintenanceService.revealInFileManager(getParentDirectory(result.archivePath))
      }
    } catch (err: any) {
      showNotification(`导出失败: ${err?.message || err}`, "error")
    } finally {
      setWorkspaceBusy(false)
    }
  }

  const handleImportWorkspaceArchive = async () => {
    setWorkspaceBusy(true)
    try {
      const zipPath = await invoke<string | null>("select_local_file", {
        title: "选择 .lamber.zip 工作区压缩包",
        extensions: ["zip", "lamber.zip"],
      })
      if (!zipPath) return
      const targetDir = await invoke<string | null>("select_local_folder")
      if (!targetDir) return
      const openAfterImport = confirm("导入成功后是否立即打开该 Workspace？")
      if (openAfterImport) {
        const canProceed = await confirmOrSave()
        if (!canProceed) return
      }
      const result = await workspaceMaintenanceService.importWorkspace(zipPath, targetDir, openAfterImport, "rename")
      if (result.opened) {
        clearAllDirty()
      }
      await refreshWorkspaceState()
      showNotification(`导入完成: ${result.workspaceRoot}`)
      await refreshWorkspaceMaintenance()
    } catch (err: any) {
      showNotification(`导入失败: ${err?.message || err}`, "error")
    } finally {
      setWorkspaceBusy(false)
    }
  }

  const handleRunWorkspaceHealthCheck = async () => {
    setWorkspaceBusy(true)
    try {
      const result = await workspaceMaintenanceService.runHealthCheck()
      setWorkspaceHealth(result)
      showNotification(`工作区健康检查完成: ${result.status}`)
      await refreshWorkspaceMaintenance()
    } catch (err: any) {
      showNotification(`工作区健康检查失败: ${err?.message || err}`, "error")
    } finally {
      setWorkspaceBusy(false)
    }
  }

  const handleRepairWorkspaceIssue = async (issueId: string) => {
    if (!confirm("修复前会先备份当前数据库。确认执行该修复项？")) return
    setWorkspaceBusy(true)
    try {
      const result = await workspaceMaintenanceService.repairIssues([issueId])
      setWorkspaceHealth(result.health)
      showNotification(`修复完成，处理 ${result.repaired} 项`)
      await refreshWorkspaceMaintenance()
    } catch (err: any) {
      showNotification(`修复失败: ${err?.message || err}`, "error")
    } finally {
      setWorkspaceBusy(false)
    }
  }

  const handleConvertInternalPaths = async () => {
    setWorkspaceBusy(true)
    try {
      const preview = await workspaceMaintenanceService.convertInternalAbsolutePathsToRelative(true)
      if (preview.candidates.length === 0) {
        showNotification("未发现需要转换的内部绝对路径")
        return
      }
      const ok = confirm(`发现 ${preview.candidates.length} 个内部绝对路径。执行前会备份数据库，是否转换为相对路径？`)
      if (!ok) return
      const result = await workspaceMaintenanceService.convertInternalAbsolutePathsToRelative(false)
      showNotification(`路径转换完成: ${result.applied} 项`)
      await handleRunWorkspaceHealthCheck()
    } catch (err: any) {
      showNotification(`路径转换失败: ${err?.message || err}`, "error")
    } finally {
      setWorkspaceBusy(false)
    }
  }

  const handleRestoreWorkspaceBackup = async (backupId: string) => {
    if (!confirm("恢复备份会替换当前 Workspace 数据库，执行前会自动备份当前数据库。确认继续？")) return
    const canProceed = await confirmOrSave()
    if (!canProceed) return
    setWorkspaceBusy(true)
    try {
      await workspaceMaintenanceService.restoreBackup(backupId)
      clearAllDirty()
      await refreshWorkspaceState()
      await refreshWorkspaceMaintenance()
      setWorkspaceHealth(null)
      showNotification("备份恢复完成")
    } catch (err: any) {
      showNotification(`恢复失败: ${err?.message || err}`, "error")
    } finally {
      setWorkspaceBusy(false)
    }
  }

  const handleDeleteWorkspaceBackup = async (backup: WorkspaceBackupInfo) => {
    if (!confirm(`确定删除该备份文件？\n\n${backup.fileName}\n\n删除后无法从这个备份恢复。`)) return
    setWorkspaceBusy(true)
    try {
      await workspaceMaintenanceService.deleteBackup(backup.id)
      setWorkspaceBackups(prev => prev.filter(item => item.id !== backup.id))
      showNotification("备份已删除")
    } catch (err: any) {
      showNotification(`删除备份失败: ${err?.message || err}`, "error")
    } finally {
      setWorkspaceBusy(false)
    }
  }

  const handleClearWorkspaceBackups = async () => {
    if (workspaceBackups.length === 0) return
    if (!confirm(`确定清空当前列表中的 ${workspaceBackups.length} 个备份文件？\n\n该操作不会影响当前数据库，但删除后无法从这些备份恢复。`)) return
    setWorkspaceBusy(true)
    try {
      const failed: string[] = []
      for (const backup of workspaceBackups) {
        try {
          await workspaceMaintenanceService.deleteBackup(backup.id)
        } catch {
          failed.push(backup.fileName)
        }
      }
      await refreshWorkspaceMaintenance()
      if (failed.length > 0) {
        showNotification(`已清理部分备份，${failed.length} 个删除失败`, "error")
      } else {
        showNotification("备份列表已清空")
      }
    } finally {
      setWorkspaceBusy(false)
    }
  }

  const handleOpenManagedWorkspace = async (path: string) => {
    if (isSameWorkspacePath(currentWorkspace?.workspaceRoot, path)) {
      setSelectedManagedWorkspacePath(path)
      navigateTo("project_board")
      return
    }
    const canProceed = await confirmOrSave()
    if (!canProceed) {
      showNotification("已取消打开工作区")
      return
    }
    setWorkspaceBusy(true)
    try {
      await openRecentWorkspace(path)
      const openedWorkspace = useWorkspaceStore.getState().currentWorkspace
      if (!isSameWorkspacePath(openedWorkspace?.workspaceRoot, path)) {
        const error = useWorkspaceStore.getState().error
        showNotification(`打开工作区失败: ${error || "未能切换到目标工作区"}`, "error")
        return
      }
      setSelectedManagedWorkspacePath(path)
      setWorkspaceHealth(null)
      setWorkspaceBackups([])
      setExternalPaths([])
      showNotification("工作区已打开")
      navigateTo("project_board")
    } catch (err: any) {
      showNotification(`打开工作区失败: ${err?.message || err}`, "error")
    } finally {
      setWorkspaceBusy(false)
    }
  }

  const handleRevealManagedWorkspace = async (path: string) => {
    if (!path) {
      showNotification("工作区路径为空，无法定位", "error")
      return
    }
    setWorkspaceBusy(true)
    try {
      await workspaceMaintenanceService.revealInFileManager(path)
    } catch (err: any) {
      showNotification(`定位失败: ${err?.message || err}`, "error")
    } finally {
      setWorkspaceBusy(false)
    }
  }

  const handleManagedWorkspaceCardKeyDown = (event: KeyboardEvent<HTMLDivElement>, path: string) => {
    if (event.key !== "Enter" && event.key !== " ") return
    event.preventDefault()
    setSelectedManagedWorkspacePath(path)
  }

  const handleForgetManagedWorkspace = async (path: string, name: string) => {
    const isCurrent = isSameWorkspacePath(currentWorkspace?.workspaceRoot, path)
    const confirmed = confirm(
      isCurrent
        ? `确定取消关联当前工作区「${name}」吗？\n\n这会关闭当前工作区并从本机列表移除，但不会删除磁盘上的工作区文件夹。`
        : `确定取消关联工作区「${name}」吗？\n\n这只会从本机列表移除记录，不会删除磁盘上的工作区文件夹。`
    )
    if (!confirmed) return
    if (isCurrent) {
      const canProceed = await confirmOrSave()
      if (!canProceed) return
    }
    setWorkspaceBusy(true)
    try {
      await forgetWorkspace(path)
      if (isSameWorkspacePath(selectedManagedWorkspacePath, path)) {
        setSelectedManagedWorkspacePath(null)
      }
      if (isCurrent) {
        clearAllDirty()
        setWorkspaceBackups([])
        setWorkspaceHealth(null)
        setExternalPaths([])
      }
      showNotification("已取消关联工作区")
    } catch (err: any) {
      showNotification(`取消关联失败: ${err?.message || err}`, "error")
    } finally {
      setWorkspaceBusy(false)
    }
  }

  const fetchRoots = useCallback(async () => {
    try {
      const data = await invoke<ProjectRoot[]>("get_project_roots")
      setRoots(data)
      if (data.length > 0) setRelocateRootId(current => current || data[0].id)
    } catch (err: any) {
      showNotification("加载根目录失败: " + err, "error")
    }
  }, [showNotification])

  useEffect(() => {
    void fetchRoots()
  }, [fetchRoots])

  useEffect(() => {
    if (activeTab === "workspace") {
      void refreshWorkspaceMaintenance()
    }
  }, [activeTab, refreshWorkspaceMaintenance])

  const handleSelectFolder = async (isNew: boolean) => {
    try {
      const path = await invoke<string | null>("select_local_folder")
      if (path) {
        if (isNew) {
          setNewRootPath(path)
          // Pre-fill name from the folder name
          const parts = path.replace(/\\/g, "/").split("/")
          const lastPart = parts[parts.length - 1] || parts[parts.length - 2] || "根目录"
          setNewRootName(lastPart)
        } else {
          setRelocateNewPath(path)
        }
      }
    } catch (err) {
      console.error(err)
    }
  }

  const handleAddRoot = async () => {
    if (!newRootName.trim() || !newRootPath.trim()) {
      showNotification("根目录名称和路径不能为空", "error")
      return
    }
    try {
      await invoke("create_project_root", {
        name: newRootName,
        rootPath: newRootPath,
        rootAlias: newRootAlias || null,
        isDefault: newRootIsDefault,
      })
      showNotification("创建项目根目录成功")
      setIsAddingRoot(false)
      setNewRootName("")
      setNewRootPath("")
      setNewRootAlias("")
      setNewRootIsDefault(false)
      fetchRoots()
    } catch (err: any) {
      showNotification("创建失败: " + err, "error")
    }
  }

  const handleSetDefault = async (id: string) => {
    try {
      await invoke("set_default_project_root", { id })
      showNotification("设置默认根目录成功")
      fetchRoots()
    } catch (err: any) {
      showNotification("设置失败: " + err, "error")
    }
  }

  const handleDeleteRoot = async (id: string) => {
    if (!confirm("确定要删除该项目根目录吗？如果项目目录或文件正在使用该根目录，删除将被阻止。")) return
    try {
      await invoke("delete_project_root", { id })
      showNotification("删除成功")
      fetchRoots()
    } catch (err: any) {
      showNotification(err, "error")
    }
  }

  const handleStartEdit = (root: ProjectRoot) => {
    setEditingRootId(root.id)
    setEditingName(root.name)
    setEditingAlias(root.rootAlias || "")
  }

  const handleSaveEdit = async (root: ProjectRoot) => {
    try {
      await invoke("update_project_root", {
        root: {
          ...root,
          name: editingName,
          rootAlias: editingAlias || null,
        },
      })
      showNotification("更新成功")
      setEditingRootId(null)
      fetchRoots()
    } catch (err: any) {
      showNotification("更新失败: " + err, "error")
    }
  }

  const handleRunHealthCheck = async () => {
    setIsCheckingHealth(true)
    try {
      const report = await invoke<HealthReport>("run_file_health_check")
      setHealthReport(report)
      showNotification(`检查完成。健康数: ${report.healthyFiles}, 自动修复数: ${report.recoverableFiles}, 断链数: ${report.missingFiles}`)
    } catch (err: any) {
      showNotification("检查失败: " + err, "error")
    } finally {
      setIsCheckingHealth(false)
    }
  }

  const handlePreviewRelocation = async () => {
    if (!relocateRootId || !relocateNewPath) {
      showNotification("请选择根目录并指定新物理路径", "error")
      return
    }
    setIsPreviewingRelocation(true)
    try {
      const preview = await invoke<RelocationPreview>("get_relocation_preview", {
        oldRootId: relocateRootId,
        newRootPath: relocateNewPath,
      })
      setRelocationPreview(preview)
    } catch (err: any) {
      showNotification("预览失败: " + err, "error")
    } finally {
      setIsPreviewingRelocation(false)
    }
  }

  const handleExecuteRelocation = async () => {
    if (!relocateRootId || !relocateNewPath) return
    if (!confirm("确定要对该根目录执行一键重定位吗？所有引用该根目录的数据库条目和项目文件都将被重定位。")) return
    setIsRelocating(true)
    try {
      await invoke("execute_bulk_relocation", {
        oldRootId: relocateRootId,
        newRootPath: relocateNewPath,
      })
      showNotification("一键批量重定位成功！")
      setRelocationPreview(null)
      setRelocateNewPath("")
      fetchRoots()
      if (healthReport) {
        handleRunHealthCheck()
      }
    } catch (err: any) {
      showNotification("重定位失败: " + err, "error")
    } finally {
      setIsRelocating(false)
    }
  }

  return (
    <div className="flex flex-col h-full bg-background animate-in fade-in duration-300">
      {/* Top Header */}
      <div className="flex items-center justify-between px-6 py-4 bg-card shadow-sm select-none">
        <div className="flex items-center gap-3">
          <button
            onClick={onBack}
            className="text-secondary-foreground hover:text-primary hover:bg-secondary font-semibold flex items-center gap-1.5 px-3 py-2 rounded-lg transition-colors"
          >
            <span>←</span> 返回集市
          </button>
          <div>
            <h1 className="text-xl font-bold tracking-tight">数据管理中心</h1>
            <p className="text-xs text-secondary-foreground">项目根目录管理、文件健康度检测与路径批量重定位</p>
          </div>
        </div>

        {/* Tab Selection */}
        <div className="flex rounded-lg bg-muted p-1 text-sm font-semibold select-none">
          <button
            onClick={() => setActiveTab("workspaces")}
            className={`px-4 py-1.5 rounded-md transition-all ${
              activeTab === "workspaces"
                ? "bg-card text-foreground shadow-sm"
                : "text-secondary-foreground hover:text-foreground"
            }`}
          >
            工作区管理
          </button>
          <button
            onClick={() => setActiveTab("workspace")}
            className={`px-4 py-1.5 rounded-md transition-all ${
              activeTab === "workspace"
                ? "bg-card text-foreground shadow-sm"
                : "text-secondary-foreground hover:text-foreground"
            }`}
          >
            工作区维护
          </button>
          <button
            onClick={() => setActiveTab("roots")}
            className={`px-4 py-1.5 rounded-md transition-all ${
              activeTab === "roots"
                ? "bg-card text-foreground shadow-sm"
                : "text-secondary-foreground hover:text-foreground"
            }`}
          >
            根目录配置
          </button>
          <button
            onClick={() => setActiveTab("health")}
            className={`px-4 py-1.5 rounded-md transition-all ${
              activeTab === "health"
                ? "bg-card text-foreground shadow-sm"
                : "text-secondary-foreground hover:text-foreground"
            }`}
          >
            健康检查与重定位
          </button>
        </div>
      </div>

      {/* Notifications */}
      {notification && (
        <div className={`fixed top-4 left-1/2 -translate-x-1/2 z-50 px-5 py-3 rounded-xl shadow-lg border text-sm font-medium transition-all animate-in fade-in slide-in-from-top-4 duration-300 ${
          notification.type === "success"
            ? "bg-success-soft border-success/20 text-success-foreground"
            : "bg-destructive-soft border-destructive/20 text-destructive"
        }`}>
          {notification.message}
        </div>
      )}

      {/* Main Content Pane */}
      <div className="flex-1 overflow-y-auto p-6 md:p-8 space-y-6">
        {activeTab === "workspaces" ? (
          <div className="mx-auto flex w-full max-w-6xl flex-col gap-6">
            <div className="flex flex-col gap-4 rounded-xl bg-card p-6 shadow-sm md:flex-row md:items-center md:justify-between">
              <div className="min-w-0">
                <div className="text-[10px] font-extrabold uppercase tracking-wide text-secondary-foreground">Workspace Registry</div>
                <h2 className="mt-1 text-2xl font-extrabold text-foreground">工作区管理</h2>
                <p className="mt-1 text-sm text-secondary-foreground">
                  管理本机记录的 Workspace。取消关联只移除本机记录，不会删除磁盘上的工作区文件夹。
                </p>
              </div>
              <div className="flex flex-wrap gap-2">
                <button disabled={workspaceBusy} onClick={refreshWorkspaceState} className="rounded-md bg-muted px-4 py-2 text-xs font-bold text-foreground hover:bg-muted/80 disabled:opacity-50">
                  刷新
                </button>
                <button disabled={workspaceBusy} onClick={handleImportWorkspaceArchive} className="inline-flex items-center gap-2 rounded-md bg-primary px-4 py-2 text-xs font-bold text-primary-foreground hover:bg-primary/90 disabled:opacity-50">
                  <Database className="h-4 w-4" />
                  导入 .lamber.zip
                </button>
              </div>
            </div>

            {recentWorkspaces.length === 0 ? (
              <div className="flex min-h-[260px] items-center justify-center rounded-xl bg-muted/40 p-10 text-center">
                <div>
                  <Database className="mx-auto mb-3 h-9 w-9 text-secondary-foreground" />
                  <div className="text-base font-bold text-foreground">暂无已关联工作区</div>
                  <div className="mt-1 text-sm text-secondary-foreground">打开、初始化或导入 Workspace 后会显示在这里。</div>
                </div>
              </div>
            ) : (
              <div className="grid grid-cols-1 gap-4 md:grid-cols-2 xl:grid-cols-3">
                {recentWorkspaces.map(item => {
                  const isCurrent = isSameWorkspacePath(currentWorkspace?.workspaceRoot, item.path)
                  const isSelected = selectedManagedWorkspacePath
                    ? isSameWorkspacePath(selectedManagedWorkspacePath, item.path)
                    : isCurrent
                  return (
                    <div
                      key={item.path}
                      role="button"
                      tabIndex={0}
                      onClick={() => setSelectedManagedWorkspacePath(item.path)}
                      onKeyDown={(event) => handleManagedWorkspaceCardKeyDown(event, item.path)}
                      className={`min-h-[190px] cursor-pointer rounded-xl bg-card p-5 text-left shadow-sm transition-all hover:-translate-y-0.5 hover:shadow-md focus:outline-none focus:ring-2 focus:ring-primary/25 ${
                        isSelected ? "ring-2 ring-primary/30" : ""
                      }`}
                    >
                      <div className="flex items-start justify-between gap-3">
                        <div className="flex min-w-0 items-center gap-3">
                          <span className={`flex h-11 w-11 shrink-0 items-center justify-center rounded-lg ${isSelected ? "bg-primary text-primary-foreground" : "bg-muted text-secondary-foreground"}`}>
                            <Database className="h-5 w-5" />
                          </span>
                          <div className="min-w-0">
                            <div className="truncate text-base font-extrabold text-foreground">{item.name || "未命名工作区"}</div>
                            <div className="mt-1 truncate font-mono text-[11px] text-secondary-foreground">{item.path}</div>
                          </div>
                        </div>
                        {isCurrent && (
                          <span className="shrink-0 rounded-md bg-primary/10 px-2 py-1 text-[10px] font-extrabold text-primary">当前</span>
                        )}
                      </div>
                      <div className="mt-6 flex items-center gap-2 text-xs font-semibold text-secondary-foreground">
                        <Clock3 className="h-3.5 w-3.5" />
                        最近打开 {formatWorkspaceTime(item.lastOpenedAt)}
                      </div>
                      <div className="mt-1 truncate font-mono text-[10px] text-secondary-foreground">
                        workspaceId: {item.workspaceId}
                      </div>
                      <div className="mt-5 flex flex-wrap gap-2">
                        <button
                          type="button"
                          disabled={workspaceBusy}
                          onClick={(event) => {
                            event.stopPropagation()
                            void handleOpenManagedWorkspace(item.path)
                          }}
                          className="rounded-md bg-primary/10 px-3 py-1.5 text-xs font-bold text-primary hover:bg-primary/15 disabled:opacity-50"
                        >
                          打开
                        </button>
                        <button
                          type="button"
                          disabled={workspaceBusy}
                          onClick={(event) => {
                            event.stopPropagation()
                            void handleRevealManagedWorkspace(item.path)
                          }}
                          className="rounded-md bg-muted px-3 py-1.5 text-xs font-bold text-foreground hover:bg-muted/80 disabled:opacity-50"
                        >
                          定位
                        </button>
                        <button
                          type="button"
                          disabled={workspaceBusy}
                          onClick={(event) => {
                            event.stopPropagation()
                            void handleForgetManagedWorkspace(item.path, item.name || "未命名工作区")
                          }}
                          className="rounded-md bg-muted px-3 py-1.5 text-xs font-bold text-secondary-foreground hover:bg-muted/80 disabled:opacity-50"
                        >
                          取消关联
                        </button>
                      </div>
                    </div>
                  )
                })}
              </div>
            )}
          </div>
        ) : activeTab === "workspace" ? (
          <div className="max-w-6xl mx-auto space-y-6">
            <div className="rounded-xl bg-card p-6 shadow-sm space-y-5">
              <div className="flex flex-col gap-4 lg:flex-row lg:items-start lg:justify-between">
                <div className="min-w-0">
                  <h3 className="font-bold text-base flex items-center gap-2">
                    <span className="w-2 h-5 rounded bg-primary"></span>
                    Workspace 迁移、备份与恢复
                  </h3>
                  <div className="mt-3 rounded-lg bg-muted/45 p-4 space-y-1">
                    <div className="text-sm font-extrabold text-foreground">{currentWorkspace?.workspaceName || "未打开工作区"}</div>
                    <div className="font-mono text-[11px] text-secondary-foreground break-all">{currentWorkspace?.workspaceRoot || "可在工作区管理中打开一个 Workspace"}</div>
                    <div className="text-[11px] text-secondary-foreground">workspaceId: {currentWorkspace?.workspaceId || "-"}</div>
                  </div>
                </div>
                <div className="flex flex-wrap gap-2">
                  <button disabled={workspaceBusy || !currentWorkspace} onClick={handleCreateWorkspaceBackup} className="rounded-lg bg-muted px-4 py-2 text-xs font-bold text-foreground hover:bg-muted/80 disabled:opacity-50">
                    立即备份
                  </button>
                  <button disabled={workspaceBusy || !currentWorkspace} onClick={handleExportWorkspace} className="rounded-lg bg-primary px-4 py-2 text-xs font-bold text-primary-foreground hover:bg-primary/90 disabled:opacity-50">
                    导出 .lamber.zip
                  </button>
                  <button disabled={workspaceBusy} onClick={handleImportWorkspaceArchive} className="rounded-lg bg-muted px-4 py-2 text-xs font-bold text-foreground hover:bg-muted/80 disabled:opacity-50">
                    导入 .lamber.zip
                  </button>
                  <button disabled={workspaceBusy || !currentWorkspace?.workspaceRoot} onClick={() => currentWorkspace?.workspaceRoot && workspaceMaintenanceService.revealInFileManager(currentWorkspace.workspaceRoot)} className="rounded-lg bg-card px-4 py-2 text-xs font-bold text-foreground shadow-sm hover:bg-muted disabled:opacity-50">
                    打开所在文件夹
                  </button>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
              <div className="rounded-xl bg-card p-6 shadow-sm space-y-4">
                <div className="flex items-center justify-between gap-3">
                  <h3 className="font-bold text-base flex items-center gap-2">
                    <span className="w-2 h-5 rounded bg-primary"></span>
                    备份列表 ({workspaceBackups.length})
                  </h3>
                  <div className="flex shrink-0 items-center gap-2">
                    <button disabled={workspaceBusy || workspaceBackups.length === 0} onClick={handleClearWorkspaceBackups} className="rounded-lg bg-muted px-3 py-1.5 text-xs font-bold text-secondary-foreground hover:bg-muted/80 disabled:opacity-50">
                      清空
                    </button>
                    <button disabled={workspaceBusy} onClick={refreshWorkspaceMaintenance} className="rounded-lg bg-muted px-3 py-1.5 text-xs font-bold hover:bg-muted/80 disabled:opacity-50">
                      刷新
                    </button>
                  </div>
                </div>
                {workspaceBackups.length === 0 ? (
                  <div className="rounded-lg bg-muted/35 p-6 text-center text-sm text-secondary-foreground">暂无备份。打开工作区时会自动生成每日数据库备份。</div>
                ) : (
                  <div className="space-y-2 max-h-80 overflow-y-auto">
                    {workspaceBackups.map(backup => (
                      <div key={backup.id} className="rounded-lg bg-muted/35 p-3 flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
                        <div className="min-w-0">
                          <div className="font-mono text-xs font-bold text-foreground truncate">{backup.fileName}</div>
                          <div className="text-[11px] text-secondary-foreground">
                            {new Date(backup.createdAt).toLocaleString()} · {(backup.sizeBytes / 1024 / 1024).toFixed(2)} MB
                          </div>
                        </div>
                        <div className="flex shrink-0 gap-2">
                          <button disabled={workspaceBusy} onClick={() => handleRestoreWorkspaceBackup(backup.id)} className="rounded-md bg-primary/10 px-3 py-1.5 text-xs font-bold text-primary hover:bg-primary/15 disabled:opacity-50">
                            恢复
                          </button>
                          <button disabled={workspaceBusy} onClick={() => workspaceMaintenanceService.revealInFileManager(backup.path)} className="rounded-md bg-card px-3 py-1.5 text-xs font-bold text-foreground shadow-sm hover:bg-muted disabled:opacity-50">
                            定位
                          </button>
                          <button disabled={workspaceBusy} onClick={() => handleDeleteWorkspaceBackup(backup)} className="rounded-md bg-card px-3 py-1.5 text-xs font-bold text-secondary-foreground shadow-sm hover:bg-muted disabled:opacity-50">
                            删除
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div className="rounded-xl bg-card p-6 shadow-sm space-y-4">
                <div className="flex items-center justify-between gap-3">
                  <h3 className="font-bold text-base flex items-center gap-2">
                    <span className="w-2 h-5 rounded bg-primary"></span>
                    Workspace 健康检查
                  </h3>
                  <div className="flex gap-2">
                    <button disabled={workspaceBusy || !currentWorkspace} onClick={handleConvertInternalPaths} className="rounded-lg bg-muted px-3 py-1.5 text-xs font-bold hover:bg-muted/80 disabled:opacity-50">
                      转换内部绝对路径
                    </button>
                    <button disabled={workspaceBusy || !currentWorkspace} onClick={handleRunWorkspaceHealthCheck} className="rounded-lg bg-primary px-3 py-1.5 text-xs font-bold text-primary-foreground hover:bg-primary/90 disabled:opacity-50">
                      运行检查
                    </button>
                  </div>
                </div>
                {!workspaceHealth ? (
                  <div className="rounded-lg bg-muted/35 p-6 text-center text-sm text-secondary-foreground">运行健康检查后，会显示结构、路径、模板资产和第三阶段状态表问题。</div>
                ) : (
                  <div className="space-y-3">
                    <div className={`inline-flex rounded-md px-2.5 py-1 text-xs font-extrabold ${
                      workspaceHealth.status === "normal" ? "bg-success-soft text-success" :
                      workspaceHealth.status === "warning" ? "bg-warning-soft text-warning-foreground" :
                      "bg-destructive-soft text-destructive"
                    }`}>
                      {workspaceHealth.status}
                    </div>
                    <div className="space-y-2 max-h-80 overflow-y-auto">
                      {workspaceHealth.items.length === 0 ? (
                        <div className="rounded-lg bg-muted/35 p-4 text-sm text-secondary-foreground">未发现问题。</div>
                      ) : workspaceHealth.items.map(item => (
                        <div key={item.id} className="rounded-lg bg-muted/35 p-3 space-y-2">
                          <div className="flex flex-wrap items-center gap-2">
                            <span className={`rounded-md px-2 py-0.5 text-[10px] font-extrabold ${
                              item.severity === "error" ? "bg-destructive-soft text-destructive" :
                              item.severity === "warning" ? "bg-warning-soft text-warning-foreground" :
                              "bg-muted text-secondary-foreground"
                            }`}>
                              {item.severity}
                            </span>
                            <span className="rounded-md bg-card px-2 py-0.5 text-[10px] font-bold text-secondary-foreground">{item.category}</span>
                            <span className="text-xs font-bold text-foreground">{item.message}</span>
                          </div>
                          {item.detail && <div className="font-mono text-[10px] text-secondary-foreground break-all">{item.detail}</div>}
                          {item.repairable && (
                            <button disabled={workspaceBusy} onClick={() => handleRepairWorkspaceIssue(item.id)} className="rounded-md bg-primary/10 px-3 py-1.5 text-xs font-bold text-primary hover:bg-primary/15 disabled:opacity-50">
                              修复
                            </button>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            </div>

            <div className="rounded-xl bg-card p-6 shadow-sm space-y-4">
              <h3 className="font-bold text-base flex items-center gap-2">
                <span className="w-2 h-5 rounded bg-primary"></span>
                外部路径提示 ({externalPaths.length})
              </h3>
              {externalPaths.length === 0 ? (
                <div className="rounded-lg bg-muted/35 p-4 text-sm text-secondary-foreground">当前未发现外部路径引用。</div>
              ) : (
                <div className="space-y-2">
                  {externalPaths.map((item, idx) => (
                    <div key={`${item.path}-${idx}`} className="rounded-lg bg-muted/35 p-3">
                      <div className="flex flex-wrap items-center gap-2">
                        <span className="rounded-md bg-card px-2 py-0.5 text-[10px] font-bold text-secondary-foreground">{item.pathType}</span>
                        <span className={item.exists ? "text-[10px] font-bold text-success" : "text-[10px] font-bold text-destructive"}>
                          {item.exists ? "存在" : "缺失"}
                        </span>
                        {item.projectName && <span className="text-[11px] text-secondary-foreground">{item.projectName}</span>}
                      </div>
                      <div className="mt-2 font-mono text-[10px] text-secondary-foreground break-all">{item.path}</div>
                      <div className="mt-1 text-[11px] text-secondary-foreground">{item.impact}</div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        ) : activeTab === "roots" ? (
          <div className="max-w-5xl mx-auto space-y-6">
            
            {/* New Root Registration */}
            <div className="rounded-xl bg-card p-6 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <h3 className="font-bold text-base flex items-center gap-2">
                  <span className="w-2 h-5 rounded bg-primary"></span>
                  登记项目根目录
                </h3>
                {!isAddingRoot && (
                  <button
                    onClick={() => setIsAddingRoot(true)}
                    className="px-4 py-2 rounded-xl text-xs font-bold bg-primary text-primary-foreground hover:bg-primary/90 transition-all flex items-center gap-1.5"
                  >
                    <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
                      <path strokeLinecap="round" strokeLinejoin="round" d="M12 4v16m8-8H4" />
                    </svg>
                    添加新根目录
                  </button>
                )}
              </div>

              {isAddingRoot && (
                <div className="bg-muted/40 rounded-xl p-5 space-y-4 animate-in slide-in-from-top-2 duration-300">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="space-y-1.5">
                      <label className="text-xs font-semibold text-secondary-foreground">根目录名称 <span className="text-red-500">*</span></label>
                      <input
                        type="text"
                        placeholder="例如: 售前方案库"
                        value={newRootName}
                        onChange={(e) => setNewRootName(e.target.value)}
                        className="w-full px-3.5 py-2 rounded-lg bg-card border-none text-sm placeholder:text-muted-foreground focus:ring-1 focus:ring-primary outline-none"
                      />
                    </div>
                    <div className="space-y-1.5">
                      <label className="text-xs font-semibold text-secondary-foreground">别名 (可选)</label>
                      <input
                        type="text"
                        placeholder="例如: 财务盘"
                        value={newRootAlias}
                        onChange={(e) => setNewRootAlias(e.target.value)}
                        className="w-full px-3.5 py-2 rounded-lg bg-card border-none text-sm placeholder:text-muted-foreground focus:ring-1 focus:ring-primary outline-none"
                      />
                    </div>
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-semibold text-secondary-foreground">物理磁盘路径 <span className="text-red-500">*</span></label>
                    <div className="flex gap-2">
                      <input
                        type="text"
                        placeholder="例如: D:\ProjectData"
                        value={newRootPath}
                        onChange={(e) => setNewRootPath(e.target.value)}
                        className="flex-1 px-3.5 py-2 rounded-lg bg-card border-none text-sm placeholder:text-muted-foreground focus:ring-1 focus:ring-primary outline-none"
                      />
                      <button
                        onClick={() => handleSelectFolder(true)}
                        className="px-4 py-2 rounded-lg bg-muted text-xs font-bold hover:bg-muted/80 transition-all flex items-center gap-1 shrink-0"
                      >
                        <AppIcon name="folder" size={14} />
                        选择文件夹
                      </button>
                    </div>
                  </div>

                  <div className="flex items-center gap-2 select-none">
                    <input
                      type="checkbox"
                      id="default-root-chk"
                      checked={newRootIsDefault}
                      onChange={(e) => setNewRootIsDefault(e.target.checked)}
                      className="rounded text-primary focus:ring-primary"
                    />
                    <label htmlFor="default-root-chk" className="text-xs font-semibold text-secondary-foreground cursor-pointer">
                      设为系统默认项目根目录
                    </label>
                  </div>

                  <div className="flex justify-end gap-2.5 pt-2">
                    <button
                      onClick={() => setIsAddingRoot(false)}
                      className="px-4 py-2 rounded-lg text-xs font-bold bg-muted hover:bg-muted/80 active:scale-[0.98] transition-all"
                    >
                      取消
                    </button>
                    <button
                      onClick={handleAddRoot}
                      className="px-5 py-2 rounded-lg text-xs font-bold bg-primary text-primary-foreground hover:bg-primary/90 active:scale-[0.98] transition-all"
                    >
                      提交登记
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Roots List */}
            <div className="rounded-xl bg-card p-6 shadow-sm space-y-4">
              <h3 className="font-bold text-base flex items-center gap-2">
                <span className="w-2 h-5 rounded bg-primary"></span>
                已登记根目录列表 ({roots.length})
              </h3>
              
              {roots.length === 0 ? (
                <div className="py-12 text-center text-secondary-foreground/70 text-sm">
                  暂未登记任何项目根目录，系统当前仅通过物理绝对路径直接关联文件。
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm border-collapse">
                    <thead>
                      <tr className="bg-muted/50 text-secondary-foreground text-xs font-semibold select-none">
                        <th className="p-3.5 first:rounded-l-lg">状态</th>
                        <th className="p-3.5">名称</th>
                        <th className="p-3.5">物理路径</th>
                        <th className="p-3.5">别名</th>
                        <th className="p-3.5 last:rounded-r-lg text-right">操作</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-muted/30">
                      {roots.map((root) => {
                        const isEditing = editingRootId === root.id
                        return (
                          <tr key={root.id} className="hover:bg-muted/20 transition-colors">
                            <td className="p-3.5">
                              {root.isDefault ? (
                                <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-primary/10 text-primary">
                                  <span className="w-1.5 h-1.5 rounded-full bg-primary animate-pulse"></span>
                                  默认
                                </span>
                              ) : (
                                <button
                                  onClick={() => handleSetDefault(root.id)}
                                  className="text-[11px] font-bold text-secondary-foreground hover:text-primary transition-colors"
                                >
                                  设为默认
                                </button>
                              )}
                            </td>
                            <td className="p-3.5 font-semibold text-foreground">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editingName}
                                  onChange={(e) => setEditingName(e.target.value)}
                                  className="px-2 py-1 rounded bg-muted text-sm outline-none w-36"
                                />
                              ) : (
                                root.name
                              )}
                            </td>
                            <td className="p-3.5 font-mono text-xs text-secondary-foreground/90 max-w-xs truncate" title={root.rootPath}>
                              {root.rootPath}
                            </td>
                            <td className="p-3.5 text-xs text-secondary-foreground">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editingAlias}
                                  onChange={(e) => setEditingAlias(e.target.value)}
                                  placeholder="别名"
                                  className="px-2 py-1 rounded bg-muted text-sm outline-none w-28"
                                />
                              ) : (
                                root.rootAlias || "-"
                              )}
                            </td>
                            <td className="p-3.5 text-right space-x-3">
                              {isEditing ? (
                                <>
                                  <button
                                    onClick={() => handleSaveEdit(root)}
                                    className="text-xs font-bold text-primary hover:underline"
                                  >
                                    保存
                                  </button>
                                  <button
                                    onClick={() => setEditingRootId(null)}
                                    className="text-xs font-bold text-secondary-foreground hover:underline"
                                  >
                                    取消
                                  </button>
                                </>
                              ) : (
                                <>
                                  <button
                                    onClick={() => handleStartEdit(root)}
                                    className="text-xs font-bold text-secondary-foreground hover:text-foreground hover:underline"
                                  >
                                    编辑
                                  </button>
                                  <button
                                    onClick={() => handleDeleteRoot(root.id)}
                                    className="text-xs font-bold text-rose-500 hover:text-rose-600 hover:underline"
                                  >
                                    删除
                                  </button>
                                </>
                              )}
                            </td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

          </div>
        ) : (
          <div className="max-w-5xl mx-auto space-y-6">
            
            {/* Health Check Dashboard */}
            <div className="rounded-xl bg-card p-6 shadow-sm space-y-5">
              <div className="flex items-center justify-between">
                <div>
                  <h3 className="font-bold text-base flex items-center gap-2">
                    <span className="w-2 h-5 rounded bg-primary"></span>
                    文件链接与路径健康度检查
                  </h3>
                  <p className="text-xs text-secondary-foreground mt-0.5">扫描全局绑定的项目文件，对于被移位的文件，自动利用文件指纹算法尝试弹性修复。</p>
                </div>
                <button
                  disabled={isCheckingHealth}
                  onClick={handleRunHealthCheck}
                  className="px-5 py-2.5 rounded-xl text-xs font-bold bg-primary text-primary-foreground hover:bg-primary/90 active:scale-[0.98] transition-all flex items-center gap-1.5 disabled:opacity-50 shrink-0"
                >
                  {isCheckingHealth ? (
                    <>
                      <span className="animate-spin rounded-full h-3.5 w-3.5 border-2 border-primary-foreground border-t-transparent"></span>
                      正在排查...
                    </>
                  ) : (
                    <>
                      <AppIcon name="quickAction" size={14} />
                      一键执行健康检查与自愈
                    </>
                  )}
                </button>
              </div>

              {healthReport && (
                <div className="space-y-5">
                  {/* Summary metrics cards */}
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-4 select-none">
                    <div className="bg-muted/40 rounded-xl p-4 text-center">
                      <div className="text-secondary-foreground text-xs font-semibold">总关联文件数</div>
                      <div className="text-2xl font-black mt-1 tabular-nums">{healthReport.totalFiles}</div>
                    </div>
                    <div className="bg-success-soft border border-success/10 rounded-xl p-4 text-center">
                      <div className="text-success-foreground text-xs font-semibold">健康文件数</div>
                      <div className="text-2xl font-black text-success mt-1 tabular-nums">{healthReport.healthyFiles}</div>
                    </div>
                    <div className="bg-warning-soft border border-warning/10 rounded-xl p-4 text-center">
                      <div className="text-warning-foreground text-xs font-semibold">自愈修复数</div>
                      <div className="text-2xl font-black text-warning mt-1 tabular-nums">{healthReport.recoverableFiles}</div>
                    </div>
                    <div className="bg-destructive-soft border border-destructive/10 rounded-xl p-4 text-center">
                      <div className="text-destructive text-xs font-semibold">已断链文件数</div>
                      <div className="text-2xl font-black text-destructive mt-1 tabular-nums">{healthReport.missingFiles}</div>
                    </div>
                  </div>

                  {/* Report details table */}
                  {healthReport.details.length > 0 && (
                    <div className="space-y-2">
                      <h4 className="text-xs font-bold text-secondary-foreground">文件健康报告明细</h4>
                      <div className="overflow-x-auto max-h-64 rounded-lg bg-muted/20">
                        <table className="w-full text-left text-xs border-collapse">
                          <thead>
                            <tr className="bg-muted/50 text-secondary-foreground font-semibold sticky top-0">
                              <th className="p-2.5">文件名</th>
                              <th className="p-2.5">最后指向物理路径</th>
                              <th className="p-2.5">检查状态</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-muted/30">
                            {healthReport.details.map((detail) => (
                              <tr key={detail.fileId} className="hover:bg-muted/10 transition-colors">
                                <td className="p-2.5 font-medium text-foreground">{detail.fileName}</td>
                                <td className="p-2.5 font-mono text-[10px] text-secondary-foreground break-all max-w-md">
                                  {detail.recoveredPath ? (
                                    <span>
                                      <span className="line-through opacity-50">{detail.currentPath}</span>
                                      <br />
                                      <span className="text-warning-foreground">→ {detail.recoveredPath}</span>
                                    </span>
                                  ) : (
                                    detail.currentPath
                                  )}
                                </td>
                                <td className="p-2.5 select-none">
                                  {detail.status === "healthy" && (
                                    <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-success-soft text-success">健康</span>
                                  )}
                                  {detail.status === "recoverable" && (
                                    <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-warning-soft text-warning-foreground">已自愈修复</span>
                                  )}
                                  {detail.status === "missing" && (
                                    <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-destructive-soft text-destructive">已断链</span>
                                  )}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Bulk Relocation */}
            <div className="rounded-xl bg-card p-6 shadow-sm space-y-4">
              <h3 className="font-bold text-base flex items-center gap-2">
                <span className="w-2 h-5 rounded bg-primary"></span>
                批量路径重定位
              </h3>
              <p className="text-xs text-secondary-foreground">
                当您在另一台电脑上重新部署项目，或手动将包含多个项目的父级目录移位后，可以在这里选择旧根目录，并定位至新的物理路径。系统会自动完成批量重映射并更新底层所有关联。
              </p>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <label className="text-xs font-semibold text-secondary-foreground">选择发生移位的旧根目录</label>
                  <select
                    value={relocateRootId}
                    onChange={(e) => setRelocateRootId(e.target.value)}
                    className="w-full px-3.5 py-2.5 rounded-lg bg-muted border-none text-sm outline-none focus:ring-1 focus:ring-primary font-semibold text-foreground cursor-pointer"
                  >
                    {roots.map(r => (
                      <option key={r.id} value={r.id}>
                        {r.name} ({r.rootAlias || "无别名"})
                      </option>
                    ))}
                  </select>
                </div>

                <div className="space-y-1.5">
                  <label className="text-xs font-semibold text-secondary-foreground">指定该根目录的新物理路径</label>
                  <div className="flex gap-2">
                    <input
                      type="text"
                      placeholder="例如: E:\LamberProjects"
                      value={relocateNewPath}
                      onChange={(e) => setRelocateNewPath(e.target.value)}
                      className="flex-1 px-3.5 py-2 rounded-lg bg-muted border-none text-sm outline-none focus:ring-1 focus:ring-primary"
                    />
                    <button
                      onClick={() => handleSelectFolder(false)}
                      className="px-4 py-2 rounded-lg bg-muted text-xs font-bold hover:bg-muted/80 transition-all flex items-center gap-1 shrink-0"
                    >
                      <AppIcon name="folder" size={14} />
                      选择新位置
                    </button>
                  </div>
                </div>
              </div>

              <div className="flex justify-end gap-2.5 pt-2 select-none">
                <button
                  disabled={isPreviewingRelocation || !relocateNewPath}
                  onClick={handlePreviewRelocation}
                  className="px-4 py-2 rounded-lg text-xs font-bold bg-muted hover:bg-muted/80 disabled:opacity-50 transition-all"
                >
                  {isPreviewingRelocation ? "正在生成预览..." : "预览重定位"}
                </button>
              </div>

              {/* Relocation Preview Result */}
              {relocationPreview && (
                <div className="bg-muted/30 rounded-xl p-5 space-y-4 animate-in slide-in-from-bottom-2 duration-300">
                  <div className="flex items-center justify-between">
                    <h4 className="text-xs font-bold text-foreground">重定位映射预览</h4>
                    <div className="text-xs text-secondary-foreground">
                      匹配成功: <span className="text-success font-bold tabular-nums">{relocationPreview.matchedItems}</span> / 
                      未找到: <span className="text-destructive font-bold tabular-nums">{relocationPreview.missingItems}</span>
                    </div>
                  </div>

                  <div className="overflow-x-auto max-h-56 rounded-lg bg-card">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-muted/40 text-secondary-foreground font-semibold sticky top-0">
                          <th className="p-2.5">类型</th>
                          <th className="p-2.5">名称</th>
                          <th className="p-2.5">旧绝对路径</th>
                          <th className="p-2.5">新映射绝对路径</th>
                          <th className="p-2.5">校验存在</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-muted/30">
                        {relocationPreview.details.map((detail, idx) => (
                          <tr key={idx} className="hover:bg-muted/10 transition-colors">
                            <td className="p-2.5">
                              {detail.itemType === "file" ? (
                                <span className="px-1.5 py-0.5 rounded bg-primary-soft text-primary font-bold">文件</span>
                              ) : (
                                <span className="px-1.5 py-0.5 rounded bg-secondary text-secondary-foreground font-bold">目录</span>
                              )}
                            </td>
                            <td className="p-2.5 font-medium text-foreground">{detail.name}</td>
                            <td className="p-2.5 font-mono text-[10px] text-secondary-foreground max-w-xs truncate" title={detail.oldPath}>
                              {detail.oldPath}
                            </td>
                            <td className="p-2.5 font-mono text-[10px] text-foreground max-w-xs truncate" title={detail.newPath}>
                              {detail.newPath}
                            </td>
                            <td className="p-2.5">
                              {detail.exists ? (
                                <span className="text-success font-bold">存在</span>
                              ) : (
                                <span className="text-destructive font-bold">缺损</span>
                              )}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <div className="flex justify-end gap-2.5 pt-2 select-none">
                    <button
                      onClick={() => setRelocationPreview(null)}
                      className="px-4 py-2 rounded-lg text-xs font-bold bg-muted hover:bg-muted/80 transition-all"
                    >
                      取消
                    </button>
                    <button
                      disabled={isRelocating}
                      onClick={handleExecuteRelocation}
                      className="px-5 py-2 rounded-lg text-xs font-bold bg-primary text-primary-foreground hover:bg-primary/90 disabled:opacity-50 transition-all flex items-center gap-1"
                    >
                      {isRelocating ? (
                        <>
                          <span className="animate-spin rounded-full h-3 w-3 border-2 border-primary-foreground border-t-transparent"></span>
                          重定位中...
                        </>
                      ) : (
                        "执行一键批量重定位"
                      )}
                    </button>
                  </div>
                </div>
              )}
            </div>

          </div>
        )}
      </div>
    </div>
  )
}
