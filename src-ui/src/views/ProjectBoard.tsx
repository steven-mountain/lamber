import { useEffect, useState, useRef } from "react";
import { List, LayoutGrid, FolderPlus } from "lucide-react";
import AppIcon from "../components/icons/AppIcon";
import { projectService, type Project, type BenefitAnalysisScheme, type BenefitAnalysisSnapshot } from "../utils/projectService";
import ProjectFilesTab from "../components/project/ProjectFilesTab";
import { projectFileService } from "../services/projectFileService";

interface ProjectBoardProps {
  onBack: () => void;
  onOpenCalc: (projectId: string, schemeId: string | null) => void;
}

const STATUS_COLUMNS = ["立项中", "审批中", "实施中", "已完成"] as const;

export default function ProjectBoard({ onBack, onOpenCalc }: ProjectBoardProps) {
  const [projects, setProjects] = useState<Project[]>([]);
  const [noteDrafts, setNoteDrafts] = useState<Record<string, string>>({});
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [projectStageFilter, setProjectStageFilter] = useState<string>("全部");

  // View Mode: list or grid
  const [viewMode, setViewMode] = useState<"list" | "grid">(
    () => (localStorage.getItem("lamber_project_board_view_mode") as "list" | "grid") || "list"
  );

  // Drawer Ref for Outside Clicks
  const drawerRef = useRef<HTMLDivElement>(null);

  // Creation State
  const [showCreateModal, setShowCreateModal] = useState(false);
  const [newProjectName, setNewProjectName] = useState("");
  const [newCustomerName, setNewCustomerName] = useState("");
  const [createProjectFolderEnabled, setCreateProjectFolderEnabled] = useState(false);
  const [createProjectFolderParent, setCreateProjectFolderParent] = useState<string | null>(null);
  const [createProjectFolderName, setCreateProjectFolderName] = useState("");
  const [showProjectFolderModal, setShowProjectFolderModal] = useState(false);

  // Details Modal State
  const [selectedProject, setSelectedProject] = useState<Project | null>(null);
  const [detailTab, setDetailTab] = useState<'info' | 'files'>('info');
  const [schemes, setSchemes] = useState<BenefitAnalysisScheme[]>([]);
  const [selectedScheme, setSelectedScheme] = useState<BenefitAnalysisScheme | null>(null);
  const [snapshots, setSnapshots] = useState<BenefitAnalysisSnapshot[]>([]);
  const [editingProjectName, setEditingProjectName] = useState("");
  const [editingCustomerName, setEditingCustomerName] = useState("");
  const [editingStatus, setEditingStatus] = useState("");
  const [isNewSchemeModalOpen, setIsNewSchemeModalOpen] = useState(false);
  const [newSchemeName, setNewSchemeName] = useState("");

  useEffect(() => {
    fetchProjects();
  }, []);

  const handleToggleViewMode = (mode: "list" | "grid") => {
    setViewMode(mode);
    localStorage.setItem("lamber_project_board_view_mode", mode);
  };

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (selectedProject && drawerRef.current && !drawerRef.current.contains(event.target as Node)) {
        const target = event.target as HTMLElement;
        if (target.closest('.z-\\[60\\]') || target.closest('.fixed.z-\\[60\\]') || target.closest('.z-60')) {
          return;
        }
        setSelectedProject(null);
      }
    };

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        if (showProjectFolderModal) {
          setShowProjectFolderModal(false);
        } else if (isNewSchemeModalOpen) {
          setIsNewSchemeModalOpen(false);
        } else if (showCreateModal) {
          setShowCreateModal(false);
        } else if (selectedProject) {
          setSelectedProject(null);
        }
      }
    };

    if (selectedProject) {
      document.addEventListener("mousedown", handleClickOutside);
    }
    document.addEventListener("keydown", handleKeyDown);

    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [selectedProject, isNewSchemeModalOpen, showCreateModal, showProjectFolderModal]);

  const fetchProjects = async () => {
    setLoading(true);
    try {
      const projs = await projectService.getProjects();
      setProjects(projs);
      setNoteDrafts(prev => {
        const next: Record<string, string> = {};
        projs.forEach(project => {
          next[project.id] = prev[project.id] ?? project.note ?? "";
        });
        return next;
      });
      setError(null);
    } catch (err) {
      console.error(err);
      setError("获取项目列表失败");
    } finally {
      setLoading(false);
    }
  };

  const handleCreateProject = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newProjectName.trim()) return;

    try {
      let newProj = await projectService.createProject(
        newProjectName,
        newCustomerName || "未知客户"
      );
      let folderBindingWarning: string | null = null;
      if (createProjectFolderEnabled && createProjectFolderParent && createProjectFolderName.trim()) {
        try {
          const folderPath = await projectFileService.createProjectFolder(
            createProjectFolderParent,
            createProjectFolderName.trim()
          );
          await projectFileService.bindProjectFolder(newProj.id, folderPath);
          newProj = (await projectService.getProject(newProj.id)) || { ...newProj, folder_path: folderPath };
        } catch (folderErr) {
          console.error(folderErr);
          folderBindingWarning = String(folderErr);
        }
      }
      setProjects((prev) => [newProj, ...prev]);
      setNoteDrafts(prev => ({ ...prev, [newProj.id]: newProj.note || "" }));
      setShowCreateModal(false);
      setNewProjectName("");
      setNewCustomerName("");
      setCreateProjectFolderEnabled(false);
      setCreateProjectFolderParent(null);
      setCreateProjectFolderName("");
      // Automatically open the details of the newly created project
      handleOpenDetails(newProj);
      if (folderBindingWarning) {
        alert("项目已创建，但项目文件夹创建或绑定失败: " + folderBindingWarning);
      }
    } catch (err) {
      console.error(err);
      alert("创建项目失败: " + err);
    }
  };

  const handleCloseCreateModal = () => {
    setShowCreateModal(false);
    setShowProjectFolderModal(false);
    setNewProjectName("");
    setNewCustomerName("");
    setCreateProjectFolderEnabled(false);
    setCreateProjectFolderParent(null);
    setCreateProjectFolderName("");
  };

  const handleOpenProjectFolderModal = async () => {
    try {
      const selected = createProjectFolderParent || await projectFileService.selectLocalFolder();
      if (!selected) return;

      setCreateProjectFolderParent(selected);
      setCreateProjectFolderName((createProjectFolderName || newProjectName || "新建项目").trim());
      setShowProjectFolderModal(true);
    } catch (err) {
      console.error(err);
      alert("选择文件夹失败: " + err);
    }
  };

  const handleChangeProjectFolderParent = async () => {
    try {
      const selected = await projectFileService.selectLocalFolder();
      if (selected) {
        setCreateProjectFolderParent(selected);
      }
    } catch (err) {
      console.error(err);
      alert("选择父级目录失败: " + err);
    }
  };

  const handleConfirmProjectFolderOption = () => {
    if (!createProjectFolderParent) {
      alert("请先选择父级目录");
      return;
    }
    if (!createProjectFolderName.trim()) {
      alert("文件夹名称不能为空");
      return;
    }

    setCreateProjectFolderEnabled(true);
    setCreateProjectFolderName(createProjectFolderName.trim());
    setShowProjectFolderModal(false);
  };

  const handleOpenDetails = async (project: Project) => {
    setSelectedProject(project);
    setDetailTab('info');
    setEditingProjectName(project.name);
    setEditingCustomerName(project.customer_name);
    setEditingStatus(project.status);

    // Fetch Schemes
    try {
      const projectSchemes = await projectService.getSchemes(project.id);
      setSchemes(projectSchemes);

      // Select the active scheme
      const defaultScheme = projectSchemes.find(s => s.id === project.default_scheme_id) || projectSchemes[0] || null;
      setSelectedScheme(defaultScheme);

      if (defaultScheme) {
        const schemeSnapshots = await projectService.getSnapshots(defaultScheme.id);
        setSnapshots(schemeSnapshots);
      } else {
        setSnapshots([]);
      }
    } catch (err) {
      console.error("加载方案/快照失败", err);
      setSchemes([]);
      setSelectedScheme(null);
      setSnapshots([]);
    }
  };

  const handleSchemeChange = async (scheme: BenefitAnalysisScheme) => {
    setSelectedScheme(scheme);
    try {
      const schemeSnapshots = await projectService.getSnapshots(scheme.id);
      setSnapshots(schemeSnapshots);
    } catch (err) {
      console.error(err);
      setSnapshots([]);
    }
  };

  const handleUpdateProjectDetails = async () => {
    if (!selectedProject || !editingProjectName.trim()) return;

    const updated: Project = {
      ...selectedProject,
      name: editingProjectName,
      customer_name: editingCustomerName,
      status: editingStatus,
      updated_at: new Date().toISOString()
    };

    try {
      const result = await projectService.updateProject(updated);
      setSelectedProject(result);
      setProjects(prev => prev.map(p => p.id === result.id ? result : p));
      alert("项目信息更新成功");
    } catch (err) {
      console.error(err);
      alert("更新失败: " + err);
    }
  };

  const handleDeleteProject = async (projectId: string) => {
    if (!confirm("确定要删除该项目吗？关联的所有方案和快照数据都将丢失。")) return;
    try {
      await projectService.deleteProject(projectId);
      setProjects(prev => prev.filter(p => p.id !== projectId));
      setNoteDrafts(prev => {
        const next = { ...prev };
        delete next[projectId];
        return next;
      });
      setSelectedProject(null);
    } catch (err) {
      console.error(err);
      alert("删除失败: " + err);
    }
  };

  const reloadSchemesForProject = async (project: Project) => {
    const projectSchemes = await projectService.getSchemes(project.id);
    setSchemes(projectSchemes);

    const defaultScheme = projectSchemes.find(s => s.id === project.default_scheme_id) || projectSchemes[0] || null;
    setSelectedScheme(defaultScheme);

    if (defaultScheme) {
      const schemeSnapshots = await projectService.getSnapshots(defaultScheme.id);
      setSnapshots(schemeSnapshots);
    } else {
      setSnapshots([]);
    }
  };

  const handleRefreshSelectedProject = async () => {
    if (!selectedProject) return;

    try {
      const latestProject = await projectService.getProject(selectedProject.id);
      if (!latestProject) {
        setSelectedProject(null);
        await fetchProjects();
        return;
      }

      setSelectedProject(latestProject);
      setEditingProjectName(latestProject.name);
      setEditingCustomerName(latestProject.customer_name);
      setEditingStatus(latestProject.status);
      setProjects(prev => prev.map(p => p.id === latestProject.id ? latestProject : p));
      setNoteDrafts(prev => ({ ...prev, [latestProject.id]: latestProject.note || "" }));
      await reloadSchemesForProject(latestProject);
    } catch (err) {
      console.error("刷新项目详情失败", err);
    }
  };

  const handleDeleteScheme = async (scheme: BenefitAnalysisScheme) => {
    if (!selectedProject) return;
    if (!confirm(`确定要删除测算方案「${scheme.name}」吗？该方案下的所有历史快照也会一并删除。`)) return;

    try {
      const updatedProject = await projectService.deleteBenefitScheme(selectedProject.id, scheme.id);
      setSelectedProject(updatedProject);
      setProjects(prev => prev.map(p => p.id === updatedProject.id ? updatedProject : p));
      await reloadSchemesForProject(updatedProject);
    } catch (err) {
      console.error(err);
      alert("删除方案失败: " + err);
    }
  };

  const handleCreateScheme = async () => {
    if (!selectedProject || !newSchemeName.trim()) return;
    try {
      // To create a scheme, we can just save it. Wait, the saveBenefitScheme Tauri command handles saving a scheme
      // by saving a mock/empty input and output initially, or we can just open it in the calculator and save it there.
      // Alternatively, we can let the user name the scheme first, then immediately launch the calculator.
      setIsNewSchemeModalOpen(false);
      const nameToUse = newSchemeName;
      setNewSchemeName("");
      // Launch calculator with this project, and null schemeId, but passing the scheme name so it creates a new one!
      // To pass the new scheme name, we can store it in local storage or pass it.
      localStorage.setItem("lamber_new_scheme_name", nameToUse);
      onOpenCalc(selectedProject.id, null);
    } catch (err) {
      console.error(err);
      alert("创建方案失败: " + err);
    }
  };

  const handleProjectNoteChange = (projectId: string, value: string) => {
    setNoteDrafts(prev => ({ ...prev, [projectId]: value }));
  };

  const handleProjectNoteBlur = async (project: Project) => {
    const nextNote = noteDrafts[project.id] ?? "";
    if ((project.note || "") === nextNote) return;

    try {
      const updatedProject = await projectService.updateProject({
        ...project,
        note: nextNote,
        updated_at: new Date().toISOString(),
      });
      setProjects(prev => prev.map(p => p.id === updatedProject.id ? updatedProject : p));
      setNoteDrafts(prev => ({ ...prev, [updatedProject.id]: updatedProject.note || "" }));
      if (selectedProject?.id === updatedProject.id) {
        setSelectedProject(updatedProject);
      }
    } catch (err) {
      console.error("保存项目备注失败", err);
      alert("保存项目备注失败: " + err);
      setNoteDrafts(prev => ({ ...prev, [project.id]: project.note || "" }));
    }
  };

  const renderProjectNote = (project: Project, compact = false) => (
    <div
      className={`rounded-xl border border-border/50 bg-background/80 ${compact ? "p-3" : "p-3.5"}`}
      onClick={(e) => e.stopPropagation()}
      onMouseDown={(e) => e.stopPropagation()}
    >
      <div className="mb-1.5 flex items-center justify-between">
        <span className="text-[10px] font-extrabold uppercase tracking-wide text-secondary-foreground">项目备注</span>
      </div>
      <textarea
        value={noteDrafts[project.id] ?? project.note ?? ""}
        onChange={(e) => handleProjectNoteChange(project.id, e.target.value)}
        onBlur={() => handleProjectNoteBlur(project)}
        rows={compact ? 2 : 3}
        placeholder="填写客户背景、推进风险、下一步动作..."
        className="block w-full resize-none rounded-lg border border-transparent bg-muted/30 px-3 py-2 text-xs leading-5 text-foreground outline-none transition-all placeholder:text-secondary-foreground/50 focus:border-primary/30 focus:bg-background focus:ring-2 focus:ring-primary/10"
      />
    </div>
  );

  const getStatusBadge = (status: Project["benefit_status"]) => {
    switch (status) {
      case "normal":
        return <span className="inline-flex items-center gap-1 text-[10px] bg-green-500/10 text-green-500 px-2 py-0.5 rounded-full font-bold border border-green-500/20"><span className="w-1.5 h-1.5 bg-green-500 rounded-full animate-pulse" /> 测算已更新</span>;
      case "outdated":
        return <span className="inline-flex items-center gap-1 text-[10px] bg-amber-500/10 text-amber-500 px-2 py-0.5 rounded-full font-bold border border-amber-500/20"><span className="w-1.5 h-1.5 bg-amber-500 rounded-full animate-pulse" /> 测算已失效</span>;
      default:
        return <span className="inline-flex items-center gap-1 text-[10px] bg-gray-500/10 text-gray-500 px-2 py-0.5 rounded-full font-bold border border-gray-500/20">未测算</span>;
    }
  };

  const parseMetricNumber = (value: string | number | null | undefined) => {
    if (typeof value === "number") return Number.isFinite(value) ? value : null;
    const cleaned = String(value ?? "").replace(/[%￥¥,\s]/g, "");
    if (!cleaned || cleaned === "--") return null;
    const numeric = Number(cleaned);
    return Number.isFinite(numeric) ? numeric : null;
  };

  const formatMetricNumber = (value: string | number | null | undefined) => {
    const numeric = parseMetricNumber(value);
    if (numeric === null) return String(value ?? "--");
    return numeric.toLocaleString("zh-CN", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  };

  const formatMetricPercent = (value: string | number | null | undefined) => {
    const numeric = parseMetricNumber(value);
    if (numeric === null) return String(value ?? "--");
    const percent = typeof value === "string" && value.includes("%")
      ? numeric
      : Math.abs(numeric) <= 1
        ? numeric * 100
        : numeric;
    return `${percent.toFixed(2)}%`;
  };

  const filteredProjects = projects.filter(p => {
    if (projectStageFilter === "全部") return true;
    return p.status === projectStageFilter;
  });

  return (
    <div className="flex flex-col flex-1 h-full overflow-hidden bg-background text-foreground animate-in fade-in duration-300">
      {/* Top Header */}
      <header className="flex items-center justify-between px-6 py-4 border-b border-border shrink-0 bg-card">
        <div className="flex items-center gap-3">
          <button
            id="board_back_btn"
            onClick={onBack}
            className="p-2 hover:bg-secondary rounded-lg transition-colors text-secondary-foreground hover:text-primary"
          >
            <AppIcon name="chevronDown" size={20} className="rotate-90" />
          </button>
          <div>
            <h1 className="text-xl font-extrabold flex items-center gap-2 text-foreground">
              <AppIcon name="project" size={22} className="text-primary" /> 项目看板
            </h1>
            <p className="text-xs text-secondary-foreground mt-0.5">管理项目生命周期及其关联的效益分析测算</p>
          </div>
        </div>

        <button
          id="board_create_project_btn"
          onClick={() => setShowCreateModal(true)}
          className="inline-flex items-center gap-1.5 bg-primary text-primary-foreground font-bold px-4 py-2 rounded-lg text-sm hover:bg-primary/95 transition-all shadow-sm active:scale-[0.98]"
        >
          <AppIcon name="save" size={16} /> 创建新项目
        </button>
      </header>

      {/* Kanban Board Container */}
      {loading ? (
        <div className="flex-1 flex flex-col items-center justify-center gap-3 text-secondary-foreground">
          <AppIcon name="loading" size={36} className="animate-spin text-primary" />
          <span className="text-sm font-semibold">正在载入项目库...</span>
        </div>
      ) : error ? (
        <div className="flex-1 flex flex-col items-center justify-center gap-2 text-red-500">
          <AppIcon name="error" size={40} />
          <span className="font-bold">{error}</span>
          <button onClick={fetchProjects} className="mt-2 text-sm text-primary underline">重试</button>
        </div>
      ) : (
        <>
          {/* Filter Bar */}
          <div className="px-6 py-4 bg-muted/20 border-b border-border/80 flex items-center justify-between shrink-0">
            <div className="flex items-center gap-2 overflow-x-auto pb-1 md:pb-0 scrollbar-none">
              <span className="text-xs text-secondary-foreground font-extrabold mr-2 uppercase tracking-wider shrink-0">项目阶段筛选:</span>
              {["全部", ...STATUS_COLUMNS].map((stage) => {
                const isActive = projectStageFilter === stage;
                const count = stage === "全部"
                  ? projects.length
                  : projects.filter(p => p.status === stage).length;

                return (
                  <button
                    key={stage}
                    onClick={() => setProjectStageFilter(stage)}
                    className={`px-3.5 py-1.5 rounded-full text-xs font-bold transition-all flex items-center gap-1.5 shrink-0 active:scale-[0.98] ${
                      isActive
                        ? "bg-primary text-primary-foreground shadow-sm font-extrabold"
                        : "bg-card border border-border text-secondary-foreground hover:bg-secondary hover:text-primary"
                    }`}
                  >
                    {stage}
                    <span className={`text-[10px] px-1.5 py-0.5 rounded-full ${
                      isActive ? "bg-primary-foreground/20 text-primary-foreground" : "bg-muted text-secondary-foreground"
                    }`}>
                      {count}
                    </span>
                  </button>
                );
              })}
            </div>

            <div className="flex items-center border border-border bg-card rounded-lg p-0.5 shrink-0">
              <button
                onClick={() => handleToggleViewMode("list")}
                className={`p-1.5 rounded-md transition-all ${
                  viewMode === "list"
                    ? "bg-secondary text-primary font-bold shadow-sm"
                    : "text-secondary-foreground hover:text-foreground opacity-60 hover:opacity-100"
                }`}
                title="列表视图"
              >
                <List className="w-4 h-4" />
              </button>
              <button
                onClick={() => handleToggleViewMode("grid")}
                className={`p-1.5 rounded-md transition-all ${
                  viewMode === "grid"
                    ? "bg-secondary text-primary font-bold shadow-sm"
                    : "text-secondary-foreground hover:text-foreground opacity-60 hover:opacity-100"
                }`}
                title="卡片视图"
              >
                <LayoutGrid className="w-4 h-4" />
              </button>
            </div>
          </div>

          {/* Project List */}
          {filteredProjects.length === 0 ? (
            <div className="flex-1 flex flex-col items-center justify-center p-12 text-secondary-foreground">
              <div className="bg-muted p-4 rounded-full text-secondary-foreground/60 mb-3">
                <AppIcon name="project" size={32} />
              </div>
              <span className="text-sm font-semibold">该阶段下暂无项目</span>
              {projectStageFilter !== "全部" && (
                <button
                  onClick={() => setProjectStageFilter("全部")}
                  className="mt-2 text-xs text-primary underline"
                >
                  查看全部项目
                </button>
              )}
            </div>
          ) : viewMode === "list" ? (
            <div className="flex-1 overflow-y-auto p-6 space-y-4">
              {filteredProjects.map((project) => {
                const metrics = project.summary_metrics;
                return (
                  <div
                    key={project.id}
                    id={`project_card_${project.id}`}
                    onClick={() => handleOpenDetails(project)}
                    className="bg-card border border-border hover:border-primary/40 rounded-xl p-4 shadow-sm hover:shadow-md transition-all duration-200 cursor-pointer flex flex-col md:flex-row md:items-center justify-between gap-4 group relative overflow-hidden animate-in fade-in slide-in-from-bottom-2"
                  >
                    {/* Left: Info */}
                    <div className="flex-1 min-w-0 flex items-start gap-4">
                      <div className="bg-primary/10 p-2.5 rounded-lg text-primary shrink-0 mt-0.5 group-hover:bg-primary/20 transition-colors">
                        <AppIcon name="project" size={20} />
                      </div>
                      <div className="min-w-0">
                        <div className="flex items-center gap-2 flex-wrap">
                          <h3 className="font-extrabold text-sm text-foreground group-hover:text-primary transition-colors truncate">
                            {project.name}
                          </h3>
                          <span className="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-secondary text-secondary-foreground border border-border/60">
                            {project.status}
                          </span>
                          {getStatusBadge(project.benefit_status)}
                        </div>
                        <p className="text-xs text-secondary-foreground mt-1.5">
                          客户名称: <span className="font-medium text-foreground">{project.customer_name || "未填写"}</span>
                        </p>
                      </div>
                    </div>

                    {/* Middle: Metrics + Notes */}
                    <div className="w-full shrink-0 md:w-[430px] lg:w-[520px] flex flex-col gap-2.5">
                      <div className="flex flex-wrap items-center gap-4 sm:gap-6 bg-muted/20 border border-border/30 px-4 py-2.5 rounded-xl">
                        {metrics ? (
                          <>
                            <div className="flex flex-col gap-0.5 min-w-[50px]">
                              <span className="text-[10px] text-secondary-foreground">毛利率</span>
                              <span className="text-xs font-bold text-foreground">{formatMetricPercent(metrics.margin_rate)}</span>
                            </div>
                            <div className="w-px h-6 bg-border/40" />
                            <div className="flex flex-col gap-0.5 min-w-[70px]">
                              <span className="text-[10px] text-secondary-foreground">净现值 NPV</span>
                              <span className="text-xs font-bold text-foreground">{formatMetricNumber(metrics.npv)}</span>
                            </div>
                            <div className="w-px h-6 bg-border/40" />
                            <div className="flex flex-col gap-0.5 min-w-[70px]">
                              <span className="text-[10px] text-secondary-foreground">净现值率 NPVR</span>
                              <span className="text-xs font-bold text-foreground">{formatMetricPercent(metrics.npv_rate)}</span>
                            </div>
                            <div className="w-px h-6 bg-border/40" />
                            <div className="flex flex-col gap-0.5 min-w-[50px]">
                              <span className="text-[10px] text-secondary-foreground">IRR</span>
                              <span className="text-xs font-bold text-foreground">{formatMetricPercent(metrics.irr)}</span>
                            </div>
                            <div className="w-px h-6 bg-border/40" />
                            <div className="flex flex-col gap-0.5 min-w-[60px]">
                              <span className="text-[10px] text-secondary-foreground">风险等级</span>
                              <span className={`text-xs font-bold ${
                                metrics.risk_level === '低风险' ? 'text-green-500' :
                                metrics.risk_level === '中风险' ? 'text-amber-500' : 'text-red-500'
                              }`}>
                                {metrics.risk_level}
                              </span>
                            </div>
                          </>
                        ) : (
                          <span className="text-xs text-secondary-foreground italic py-1">暂无效益分析指标</span>
                        )}
                      </div>
                      {renderProjectNote(project, true)}
                    </div>

                    {/* Right: Actions */}
                    <div className="shrink-0 flex items-center gap-4 self-end md:self-auto">
                      <span className="text-[10px] text-secondary-foreground hidden lg:inline">
                        {new Date(project.updated_at).toLocaleDateString()} 更新
                      </span>
                      <button
                        id={`open_calc_btn_${project.id}`}
                        onClick={(e) => {
                          e.stopPropagation();
                          onOpenCalc(project.id, project.default_scheme_id || null);
                        }}
                        className="text-xs font-bold text-primary hover:text-primary-foreground hover:bg-primary px-3 py-1.5 rounded-lg transition-all border border-primary/20 hover:border-transparent flex items-center gap-1.5 bg-background shadow-sm active:scale-[0.98]"
                      >
                        <AppIcon name="calculator" size={13} /> 打开效益分析
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
            <div className="flex-1 overflow-y-auto p-6">
              <div
                className="grid gap-6 animate-in fade-in duration-300"
                style={{ gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))' }}
              >
                {filteredProjects.map((project) => {
                  const metrics = project.summary_metrics;
                  return (
                    <div
                      key={project.id}
                      id={`project_card_${project.id}`}
                      onClick={() => handleOpenDetails(project)}
                      className="bg-card border border-border hover:border-primary/40 rounded-2xl p-5 shadow-sm hover:shadow-md transition-all duration-200 cursor-pointer flex flex-col justify-between gap-4 group relative overflow-hidden animate-in fade-in slide-in-from-bottom-2"
                    >
                      {/* Top: Name and Badges */}
                      <div className="space-y-3 shrink-0">
                        <div className="flex items-start justify-between gap-2">
                          <div className="bg-primary/10 p-2.5 rounded-xl text-primary shrink-0 group-hover:bg-primary/20 transition-colors">
                            <AppIcon name="project" size={20} />
                          </div>
                          <div className="flex flex-col items-end gap-1.5 shrink-0">
                            <span className="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-secondary text-secondary-foreground border border-border/60">
                              {project.status}
                            </span>
                            {getStatusBadge(project.benefit_status)}
                          </div>
                        </div>

                        <div className="space-y-1">
                          <h3 className="font-extrabold text-base text-foreground group-hover:text-primary transition-colors truncate" title={project.name}>
                            {project.name}
                          </h3>
                          <p className="text-xs text-secondary-foreground truncate">
                            客户名称: <span className="font-medium text-foreground">{project.customer_name || "未填写"}</span>
                          </p>
                        </div>
                      </div>

                      {/* Middle: Metrics */}
                      <div className="flex-1 bg-muted/20 border border-border/30 rounded-xl p-4 flex flex-col justify-center min-h-[100px]">
                        {metrics ? (
                          <div className="grid grid-cols-2 gap-x-4 gap-y-3">
                            <div className="flex flex-col gap-0.5">
                              <span className="text-[10px] text-secondary-foreground">毛利率</span>
                              <span className="text-xs font-bold text-foreground">{formatMetricPercent(metrics.margin_rate)}</span>
                            </div>
                            <div className="flex flex-col gap-0.5">
                              <span className="text-[10px] text-secondary-foreground">净现值 NPV</span>
                              <span className="text-xs font-bold text-foreground">{formatMetricNumber(metrics.npv)}</span>
                            </div>
                            <div className="flex flex-col gap-0.5">
                              <span className="text-[10px] text-secondary-foreground">净现值率 NPVR</span>
                              <span className="text-xs font-bold text-foreground">{formatMetricPercent(metrics.npv_rate)}</span>
                            </div>
                            <div className="flex flex-col gap-0.5">
                              <span className="text-[10px] text-secondary-foreground">IRR</span>
                              <span className="text-xs font-bold text-foreground">{formatMetricPercent(metrics.irr)}</span>
                            </div>
                            <div className="flex flex-col gap-0.5 col-span-2 border-t border-border/30 pt-2 flex flex-row justify-between items-center">
                              <span className="text-[10px] text-secondary-foreground">风险等级</span>
                              <span className={`text-xs font-bold ${
                                metrics.risk_level === '低风险' ? 'text-green-500' :
                                metrics.risk_level === '中风险' ? 'text-amber-500' : 'text-red-500'
                              }`}>
                                {metrics.risk_level}
                              </span>
                            </div>
                          </div>
                        ) : (
                          <div className="text-center py-4 flex flex-col items-center justify-center gap-2">
                            <span className="text-xs text-secondary-foreground opacity-60 leading-relaxed">
                              暂无效益分析指标，点击下方按钮开始测算
                            </span>
                          </div>
                        )}
                      </div>

                      {renderProjectNote(project)}

                      {/* Bottom: Action and Update date */}
                      <div className="flex items-center justify-between border-t border-border/40 pt-3 shrink-0">
                        <span className="text-[10px] text-secondary-foreground">
                          {new Date(project.updated_at).toLocaleDateString()} 更新
                        </span>
                        <button
                          id={`open_calc_btn_grid_${project.id}`}
                          onClick={(e) => {
                            e.stopPropagation();
                            onOpenCalc(project.id, project.default_scheme_id || null);
                          }}
                          className="text-xs font-bold text-primary hover:text-primary-foreground hover:bg-primary px-3 py-1.5 rounded-lg transition-all border border-primary/20 hover:border-transparent flex items-center gap-1.5 bg-background shadow-sm active:scale-[0.98]"
                        >
                          <AppIcon name="calculator" size={13} /> 打开效益分析
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </>
      )}

      {/* Create Project Modal */}
      {showCreateModal && (
        <div className="fixed inset-0 z-50 bg-background/80 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in">
          <form
            onSubmit={handleCreateProject}
            className="bg-card border border-border rounded-xl shadow-2xl w-full max-w-md overflow-hidden"
          >
            <div className="px-6 py-4 border-b border-border bg-muted/30 flex items-center justify-between">
              <h2 className="font-bold text-lg text-foreground flex items-center gap-2">
                <AppIcon name="project" size={18} className="text-primary" /> 新增项目
              </h2>
              <button
                type="button"
                onClick={handleCloseCreateModal}
                className="text-secondary-foreground hover:bg-secondary p-1 rounded-md"
              >
                <AppIcon name="close" size={16} />
              </button>
            </div>

            <div className="p-6 space-y-4">
              <div className="flex flex-col gap-1.5">
                <label className="text-sm font-semibold text-secondary-foreground">项目名称 <span className="text-red-500">*</span></label>
                <input
                  id="new_project_name_input"
                  type="text"
                  required
                  placeholder="请输入项目名称"
                  value={newProjectName}
                  onChange={(e) => setNewProjectName(e.target.value)}
                  className="bg-background border border-border px-3 py-2 rounded-lg text-sm outline-none focus:border-primary w-full"
                />
              </div>

              <div className="flex flex-col gap-1.5">
                <label className="text-sm font-semibold text-secondary-foreground">客户名称</label>
                <input
                  id="new_customer_name_input"
                  type="text"
                  placeholder="请输入客户名称"
                  value={newCustomerName}
                  onChange={(e) => setNewCustomerName(e.target.value)}
                  className="bg-background border border-border px-3 py-2 rounded-lg text-sm outline-none focus:border-primary w-full"
                />
              </div>

              <div className="rounded-lg border border-border bg-muted/20 p-3">
                <div className="flex items-start justify-between gap-3">
                  <div className="min-w-0">
                    <div className="flex items-center gap-2 text-sm font-bold text-foreground">
                      <FolderPlus className="h-4 w-4 text-primary" />
                      新建并绑定项目文件夹
                    </div>
                    <p className="mt-1 text-xs leading-5 text-secondary-foreground">
                      可选。文件夹名称默认使用项目名称，创建前可以单独修改。
                    </p>
                    {createProjectFolderEnabled && createProjectFolderParent && (
                      <div className="mt-2 space-y-1">
                        <div className="text-xs font-bold text-primary">将创建：{createProjectFolderName}</div>
                        <code className="block rounded-md border border-border/50 bg-background px-2 py-1 font-mono text-[10px] leading-4 text-secondary-foreground break-all">
                          {createProjectFolderParent}
                        </code>
                      </div>
                    )}
                  </div>
                  <div className="flex shrink-0 flex-col gap-2">
                    <button
                      type="button"
                      onClick={handleOpenProjectFolderModal}
                      className="rounded-lg bg-primary/10 px-3 py-1.5 text-xs font-bold text-primary transition-all hover:bg-primary/15"
                    >
                      {createProjectFolderEnabled ? "修改" : "设置"}
                    </button>
                    {createProjectFolderEnabled && (
                      <button
                        type="button"
                        onClick={() => {
                          setCreateProjectFolderEnabled(false);
                          setCreateProjectFolderParent(null);
                          setCreateProjectFolderName("");
                        }}
                        className="rounded-lg px-3 py-1.5 text-xs font-bold text-red-500 transition-all hover:bg-red-500/10"
                      >
                        取消
                      </button>
                    )}
                  </div>
                </div>
              </div>
            </div>

            <div className="border-t border-border p-4 bg-muted/20 flex justify-end gap-3">
              <button
                type="button"
                onClick={handleCloseCreateModal}
                className="px-4 py-2 border border-border hover:bg-secondary rounded-lg text-sm font-semibold text-secondary-foreground transition-all"
              >
                取消
              </button>
              <button
                id="submit_create_project_btn"
                type="submit"
                className="px-4 py-2 bg-primary hover:bg-primary/90 text-primary-foreground font-bold rounded-lg text-sm transition-all"
              >
                确认创建
              </button>
            </div>
          </form>
        </div>
      )}

      {showProjectFolderModal && (
        <div className="fixed inset-0 z-[60] bg-slate-950/35 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in">
          <div className="bg-card border border-border rounded-xl shadow-2xl w-full max-w-md overflow-hidden">
            <div className="px-6 py-4 border-b border-border bg-muted/30 flex items-center justify-between">
              <h2 className="font-bold text-base text-foreground flex items-center gap-2">
                <FolderPlus className="h-4 w-4 text-primary" />
                设置项目文件夹
              </h2>
              <button
                type="button"
                onClick={() => setShowProjectFolderModal(false)}
                className="text-secondary-foreground hover:bg-secondary p-1 rounded-md"
              >
                <AppIcon name="close" size={16} />
              </button>
            </div>

            <div className="p-6 space-y-4">
              <div className="space-y-1.5">
                <div className="flex items-center justify-between gap-3">
                  <label className="text-sm font-semibold text-secondary-foreground">父级目录</label>
                  <button
                    type="button"
                    onClick={handleChangeProjectFolderParent}
                    className="text-xs font-bold text-primary hover:underline"
                  >
                    更换
                  </button>
                </div>
                <code className="block rounded-lg border border-border/60 bg-muted/40 px-3 py-2 font-mono text-[10px] leading-4 text-primary break-all">
                  {createProjectFolderParent}
                </code>
              </div>

              <div className="flex flex-col gap-1.5">
                <label htmlFor="new-project-folder-name" className="text-sm font-semibold text-secondary-foreground">
                  文件夹名称
                </label>
                <input
                  id="new-project-folder-name"
                  autoFocus
                  value={createProjectFolderName}
                  onChange={(e) => setCreateProjectFolderName(e.target.value)}
                  placeholder="请输入文件夹名称"
                  className="bg-background border border-border px-3 py-2 rounded-lg text-sm outline-none focus:border-primary w-full"
                />
              </div>
            </div>

            <div className="border-t border-border p-4 bg-muted/20 flex justify-end gap-3">
              <button
                type="button"
                onClick={() => setShowProjectFolderModal(false)}
                className="px-4 py-2 border border-border hover:bg-secondary rounded-lg text-sm font-semibold text-secondary-foreground transition-all"
              >
                取消
              </button>
              <button
                type="button"
                onClick={handleConfirmProjectFolderOption}
                className="px-4 py-2 bg-primary hover:bg-primary/90 text-primary-foreground font-bold rounded-lg text-sm transition-all"
              >
                确认
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Project Details / Schemes & Snapshot Modal */}
      {selectedProject && (
        <div className="fixed inset-0 z-50 bg-background/80 backdrop-blur-sm flex items-center justify-center p-4 sm:p-6 animate-in fade-in">
          <div
            ref={drawerRef}
            className="bg-card border border-border shadow-2xl w-[88vw] max-w-6xl h-[86vh] max-h-[920px] min-h-[560px] rounded-2xl overflow-hidden flex flex-col animate-in zoom-in-95 fade-in duration-200"
          >
            {/* Modal Header */}
            <div className="px-6 py-4 border-b border-border bg-muted/30 flex items-center justify-between shrink-0">
              <div>
                <h2 className="font-extrabold text-lg text-foreground flex items-center gap-2">
                  项目详情：{selectedProject.name}
                </h2>
                <p className="text-xs text-secondary-foreground mt-0.5">查看及修改项目基本信息、历史效益评估版本</p>
              </div>
              <button
                onClick={() => setSelectedProject(null)}
                className="text-secondary-foreground hover:bg-secondary p-1 rounded-md"
              >
                <AppIcon name="close" size={20} />
              </button>
            </div>

            {/* Tab Navigation */}
            <div className="flex border-b border-border bg-muted/10 shrink-0">
              <button
                onClick={() => setDetailTab('info')}
                className={`py-3 px-6 text-xs font-bold transition-all border-b-2 -mb-[2px] ${
                  detailTab === 'info'
                    ? "border-primary text-primary"
                    : "border-transparent text-secondary-foreground opacity-70 hover:opacity-100"
                }`}
              >
                效益分析与日志
              </button>
              <button
                onClick={() => setDetailTab('files')}
                className={`py-3 px-6 text-xs font-bold transition-all border-b-2 -mb-[2px] ${
                  detailTab === 'files'
                    ? "border-primary text-primary"
                    : "border-transparent text-secondary-foreground opacity-70 hover:opacity-100"
                }`}
              >
                项目文件管理
              </button>
            </div>

            {/* Modal Body */}
            {detailTab === 'info' ? (
              <div className="flex-1 overflow-y-auto p-6 space-y-6 scrollbar-thin">
                {/* Part 1: Edit Project Metadata */}
                <section className="bg-muted/20 border border-border/80 rounded-xl p-4 space-y-4">
                  <h3 className="font-bold text-sm text-foreground border-b border-border/60 pb-1.5 flex items-center gap-2">
                    <AppIcon name="settings" size={16} className="text-primary" /> 基本属性
                  </h3>
                  <div className="grid grid-cols-2 gap-4">
                    <div className="flex flex-col gap-1">
                      <label className="text-xs font-semibold text-secondary-foreground">项目名称</label>
                      <input
                        type="text"
                        value={editingProjectName}
                        onChange={(e) => setEditingProjectName(e.target.value)}
                        className="bg-background border border-border px-3 py-2 rounded-lg text-sm outline-none focus:border-primary"
                      />
                    </div>
                    <div className="flex flex-col gap-1">
                      <label className="text-xs font-semibold text-secondary-foreground">客户名称</label>
                      <input
                        type="text"
                        value={editingCustomerName}
                        onChange={(e) => setEditingCustomerName(e.target.value)}
                        className="bg-background border border-border px-3 py-2 rounded-lg text-sm outline-none focus:border-primary"
                      />
                    </div>
                    <div className="flex flex-col gap-1 col-span-2">
                      <label className="text-xs font-semibold text-secondary-foreground">看板阶段</label>
                      <select
                        value={editingStatus}
                        onChange={(e) => setEditingStatus(e.target.value)}
                        className="bg-background border border-border px-3 py-2 rounded-lg text-sm outline-none focus:border-primary"
                      >
                        {STATUS_COLUMNS.map(col => (
                          <option key={col} value={col}>{col}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                  <div className="flex justify-between items-center pt-2">
                    <button
                      type="button"
                      onClick={() => handleDeleteProject(selectedProject.id)}
                      className="inline-flex items-center gap-1 text-xs text-red-500 hover:text-red-700 font-bold"
                    >
                      <AppIcon name="delete" size={14} /> 删除项目
                    </button>
                    <button
                      type="button"
                      onClick={handleUpdateProjectDetails}
                      className="bg-primary text-primary-foreground font-bold px-3 py-1.5 rounded-lg text-xs hover:bg-primary/95 transition-all shadow-sm"
                    >
                      更新基本信息
                    </button>
                  </div>
                </section>

                {/* Part 2: Schemes and Snapshots */}
                <section className="space-y-4">
                  <div className="flex items-center justify-between border-b border-border pb-2">
                    <h3 className="font-bold text-sm text-foreground flex items-center gap-2">
                      <AppIcon name="calculator" size={16} className="text-primary" /> 效益分析测算方案
                    </h3>
                    <button
                      onClick={() => {
                        if (selectedProject) {
                          setNewSchemeName(selectedProject.name);
                        }
                        setIsNewSchemeModalOpen(true);
                      }}
                      className="inline-flex items-center gap-1 text-xs text-primary font-bold hover:underline"
                    >
                      <AppIcon name="save" size={12} /> 新增方案
                    </button>
                  </div>

                  {/* Scheme list tags */}
                  <div className="flex flex-wrap gap-2">
                    {schemes.length === 0 ? (
                      <span className="text-xs text-secondary-foreground">暂无测算方案</span>
                    ) : (
                      schemes.map((scheme) => {
                        const isDefault = selectedProject.default_scheme_id === scheme.id;
                        const isSelected = selectedScheme?.id === scheme.id;
                        return (
                          <div
                            key={scheme.id}
                            className={`rounded-lg text-xs font-semibold border transition-all flex items-center gap-1 overflow-hidden ${
                              isSelected
                                ? "bg-primary/10 text-primary border-primary"
                                : "bg-muted/40 hover:bg-secondary text-secondary-foreground border-border"
                            }`}
                          >
                            <button
                              onClick={() => handleSchemeChange(scheme)}
                              className="px-3 py-1.5 flex items-center gap-1.5 min-w-0"
                              title={scheme.name}
                            >
                              <span className="truncate max-w-[160px]">{scheme.name}</span>
                              {isDefault && <span className="bg-primary text-primary-foreground text-[8px] px-1 rounded shrink-0">默认</span>}
                            </button>
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                handleDeleteScheme(scheme);
                              }}
                              className="px-2 py-1.5 border-l border-border/60 text-secondary-foreground hover:text-red-500 hover:bg-red-500/10 transition-all"
                              title="删除测算方案"
                            >
                              <AppIcon name="delete" size={12} />
                            </button>
                          </div>
                        );
                      })
                    )}
                  </div>

                  {/* Snapshots / Versions List */}
                  {selectedScheme && (
                    <div className="bg-muted/10 border border-border/80 rounded-xl p-4 space-y-4">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold text-foreground">
                          方案 「{selectedScheme.name}」 的历史测算快照 ({snapshots.length} 个版本)
                        </span>
                        <button
                          onClick={() => onOpenCalc(selectedProject.id, selectedScheme.id)}
                          className="text-xs text-primary font-bold flex items-center gap-1 hover:underline"
                        >
                          <AppIcon name="calculator" size={12} /> 开展测算
                        </button>
                      </div>

                      <div className="space-y-3">
                        {snapshots.length === 0 ? (
                          <div className="text-center py-6 text-xs text-secondary-foreground border border-dashed border-border rounded-lg">
                            该方案暂无测算快照版本，请打开测算页面测算并保存。
                          </div>
                        ) : (
                          snapshots.map((snap) => (
                            <div
                              key={snap.id}
                              className="bg-card border border-border/60 hover:border-border rounded-lg p-3 flex flex-col md:flex-row md:items-center justify-between gap-3 text-xs shadow-sm hover:shadow transition-all"
                            >
                              <div>
                                <div className="flex items-center gap-2">
                                  <span className="font-bold text-primary font-mono text-sm">v{snap.version}</span>
                                  <span className="text-secondary-foreground">{new Date(snap.created_at).toLocaleString()}</span>
                                </div>
                                <div className="grid grid-cols-3 gap-x-4 gap-y-1 mt-2 text-secondary-foreground font-mono text-[10px]">
                                  <div>毛利率: <span className="font-bold text-foreground">{formatMetricPercent(snap.output_metrics.margin_rate)}</span></div>
                                  <div>NPV: <span className="font-bold text-foreground">{formatMetricNumber(snap.output_metrics.npv)}</span></div>
                                  <div>IRR: <span className="font-bold text-foreground">{formatMetricPercent(snap.output_metrics.irr)}</span></div>
                                </div>
                              </div>
                              <button
                                onClick={() => onOpenCalc(selectedProject.id, selectedScheme.id)}
                                className="self-end md:self-auto bg-secondary hover:bg-primary hover:text-primary-foreground font-bold px-3 py-1.5 rounded-lg border border-border/50 hover:border-transparent text-xs transition-all flex items-center gap-1"
                              >
                                <AppIcon name="calculator" size={12} /> 加载并打开
                              </button>
                            </div>
                          ))
                        )}
                      </div>
                    </div>
                  )}
                </section>

                {/* Part 3: Project Logs */}
                <section className="space-y-3">
                  <h3 className="font-bold text-sm text-foreground border-b border-border pb-1.5 flex items-center gap-2">
                    <AppIcon name="batch" size={16} className="text-primary" /> 项目日志与流转痕迹
                  </h3>
                  <div className="space-y-3 max-h-60 overflow-y-auto pr-1">
                    {selectedProject.logs.length === 0 ? (
                      <span className="text-xs text-secondary-foreground block py-2">暂无日志</span>
                    ) : (
                      selectedProject.logs.map((log) => (
                        <div key={log.id} className="flex gap-2 text-xs border-l-2 border-primary/20 pl-3 py-0.5">
                          <span className="text-secondary-foreground shrink-0 font-mono text-[10px]">
                            {new Date(log.timestamp).toLocaleString()}
                          </span>
                          <span className="text-foreground">{log.description}</span>
                        </div>
                      ))
                    )}
                  </div>
                </section>
              </div>
            ) : (
              <div className="flex-1 overflow-hidden flex flex-col">
                <ProjectFilesTab
                  projectId={selectedProject.id}
                  onRefreshProject={handleRefreshSelectedProject}
                />
              </div>
            )}

            {/* Modal Footer */}
            <div className="border-t border-border p-4 bg-muted/20 flex justify-end shrink-0">
              <button
                onClick={() => setSelectedProject(null)}
                className="px-4 py-2 bg-secondary hover:bg-secondary/80 font-bold rounded-lg text-sm text-foreground transition-all"
              >
                关闭
              </button>
            </div>
          </div>
        </div>
      )}

      {/* New Scheme Modal */}
      {isNewSchemeModalOpen && (
        <div className="fixed inset-0 z-[60] bg-background/80 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in">
          <div className="bg-card border border-border rounded-xl shadow-2xl w-full max-w-sm overflow-hidden">
            <div className="px-6 py-4 border-b border-border bg-muted/30 flex items-center justify-between">
              <h4 className="font-bold text-sm text-foreground">新增效益分析测算方案</h4>
              <button
                type="button"
                onClick={() => setIsNewSchemeModalOpen(false)}
                className="text-secondary-foreground hover:bg-secondary p-1 rounded-md"
              >
                <AppIcon name="close" size={14} />
              </button>
            </div>
            <div className="p-6">
              <label className="text-xs font-semibold text-secondary-foreground block mb-1.5">方案名称 <span className="text-red-500">*</span></label>
              <input
                type="text"
                required
                placeholder="例如：方案 A、设备扩容二期测算"
                value={newSchemeName}
                onChange={(e) => setNewSchemeName(e.target.value)}
                className="bg-background border border-border px-3 py-2 rounded-lg text-xs outline-none focus:border-primary w-full"
              />
            </div>
            <div className="border-t border-border p-3 bg-muted/10 flex justify-end gap-2">
              <button
                onClick={() => setIsNewSchemeModalOpen(false)}
                className="px-3 py-1.5 border border-border hover:bg-secondary rounded-lg text-xs font-semibold text-secondary-foreground transition-all"
              >
                取消
              </button>
              <button
                onClick={handleCreateScheme}
                disabled={!newSchemeName.trim()}
                className="px-3 py-1.5 bg-primary hover:bg-primary/90 disabled:opacity-50 text-white font-bold rounded-lg text-xs transition-all"
              >
                确认并去测算
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
