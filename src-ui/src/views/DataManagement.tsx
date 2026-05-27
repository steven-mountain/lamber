import { useState, useEffect } from "react"
import { invoke } from "@tauri-apps/api/core"
import AppIcon from "../components/icons/AppIcon"

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
  const [activeTab, setActiveTab] = useState<"roots" | "health">("roots")
  
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

  useEffect(() => {
    fetchRoots()
  }, [])

  const showNotification = (message: string, type: "success" | "error" = "success") => {
    setNotification({ message, type })
    setTimeout(() => setNotification(null), 4000)
  }

  const fetchRoots = async () => {
    try {
      const data = await invoke<ProjectRoot[]>("get_project_roots")
      setRoots(data)
      if (data.length > 0 && !relocateRootId) {
        setRelocateRootId(data[0].id)
      }
    } catch (err: any) {
      showNotification("加载根目录失败: " + err, "error")
    }
  }

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
            className="p-2 rounded-lg bg-muted text-secondary-foreground hover:bg-muted/80 hover:text-foreground transition-all flex items-center justify-center"
          >
            <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M10 19l-7-7m0 0l7-7m-7 7h18" />
            </svg>
          </button>
          <div>
            <h1 className="text-xl font-bold tracking-tight">数据管理中心</h1>
            <p className="text-xs text-secondary-foreground">项目根目录管理、文件健康度检测与路径批量重定位</p>
          </div>
        </div>

        {/* Tab Selection */}
        <div className="flex rounded-lg bg-muted p-1 text-sm font-semibold select-none">
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
            ? "bg-green-50 border-green-200 text-green-800"
            : "bg-red-50 border-red-200 text-red-800"
        }`}>
          {notification.message}
        </div>
      )}

      {/* Main Content Pane */}
      <div className="flex-1 overflow-y-auto p-6 md:p-8 space-y-6">
        {activeTab === "roots" ? (
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
                    <div className="bg-green-50/50 rounded-xl p-4 text-center">
                      <div className="text-green-700 text-xs font-semibold">健康文件数</div>
                      <div className="text-2xl font-black text-green-600 mt-1 tabular-nums">{healthReport.healthyFiles}</div>
                    </div>
                    <div className="bg-amber-50/50 rounded-xl p-4 text-center">
                      <div className="text-amber-700 text-xs font-semibold">自愈修复数</div>
                      <div className="text-2xl font-black text-amber-500 mt-1 tabular-nums">{healthReport.recoverableFiles}</div>
                    </div>
                    <div className="bg-red-50/50 rounded-xl p-4 text-center">
                      <div className="text-red-700 text-xs font-semibold">已断链文件数</div>
                      <div className="text-2xl font-black text-red-500 mt-1 tabular-nums">{healthReport.missingFiles}</div>
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
                                      <span className="text-amber-600">→ {detail.recoveredPath}</span>
                                    </span>
                                  ) : (
                                    detail.currentPath
                                  )}
                                </td>
                                <td className="p-2.5 select-none">
                                  {detail.status === "healthy" && (
                                    <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-green-500/10 text-green-600">健康</span>
                                  )}
                                  {detail.status === "recoverable" && (
                                    <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-amber-500/10 text-amber-500">已自愈修复</span>
                                  )}
                                  {detail.status === "missing" && (
                                    <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-red-500/10 text-red-500">已断链</span>
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
                      匹配成功: <span className="text-green-600 font-bold tabular-nums">{relocationPreview.matchedItems}</span> / 
                      未找到: <span className="text-red-500 font-bold tabular-nums">{relocationPreview.missingItems}</span>
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
                                <span className="px-1.5 py-0.5 rounded bg-blue-500/10 text-blue-500 font-bold">文件</span>
                              ) : (
                                <span className="px-1.5 py-0.5 rounded bg-purple-500/10 text-purple-500 font-bold">目录</span>
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
                                <span className="text-green-500 font-bold">存在</span>
                              ) : (
                                <span className="text-red-500 font-bold">缺损</span>
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
