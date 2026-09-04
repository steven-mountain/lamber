import { useEffect, useState, useRef } from "react";
import { ArrowRight, BarChart3, FileText, Info, List, LayoutGrid, FolderPlus, FolderOpen, Plus, Search, Settings2, StickyNote, X, ChevronDown, ChevronUp, AlertTriangle, RefreshCw } from "lucide-react";
import AppIcon from "../components/icons/AppIcon";
import { projectService, type Project, type ProjectType, type BenefitAnalysisScheme, type BenefitAnalysisSnapshot, type SummaryMetrics } from "../utils/projectService";
import ProjectFilesTab from "../components/project/ProjectFilesTab";
import { projectFileService } from "../services/projectFileService";
import { invoke } from "@tauri-apps/api/core";
import WorkspaceGate from "../components/workspace/WorkspaceGate";
import { useWorkspaceStore } from "../store/useWorkspaceStore";
import { useProjectStore } from "../store/useProjectStore";
import { useSaveStore } from "../store/useSaveStore";
import { useAiContextStore } from "../store/useAiContextStore";
import { useUnsavedChangesGuard } from "../hooks/useUnsavedChangesGuard";
import { domainSaveService, type ProjectDetailPatch } from "../services/domainSaveService";
import GlobalSaveButton from "../components/GlobalSaveButton";
import { useNavigationStore } from "../store/useNavigationStore";
import { SCHEME_STAGE_OPTIONS, getSchemeStageOption, type SchemeStage } from "../lib/schemeStage";
import {
  projectPresetService,
  type ProjectPresetTemplate,
} from "../services/projectPresetService";

interface CandidateFile {
  name: string;
  path: string;
  fileRole: 'benefit_scheme' | 'budget' | 'proposal' | 'other';
}

interface ImportCandidate {
  folderName: string;
  folderPath: string;
  existsConflict: boolean;
  files: CandidateFile[];
}

interface ProjectBoardProps {
  onBack: () => void;
  onOpenCalc: (projectId: string, schemeId: string | null) => void;
}

const DEFAULT_STATUS_COLUMNS = ["需求导入", "会审纪要", "甄选"];
const PROJECT_STATUS_STORAGE_KEY = "lamber_project_board_status_options";

const dedupeStatusOptions = (options: string[]) => {
  const seen = new Set<string>();
  return options
    .map(option => option.trim())
    .filter(option => {
      if (!option || seen.has(option)) return false;
      seen.add(option);
      return true;
    });
};

const readStoredStatusOptions = () => {
  try {
    const stored = localStorage.getItem(PROJECT_STATUS_STORAGE_KEY);
    if (!stored) return DEFAULT_STATUS_COLUMNS;
    const parsed = JSON.parse(stored);
    if (Array.isArray(parsed)) {
      const options = dedupeStatusOptions(parsed.filter(item => typeof item === "string"));
      return options.length > 0 ? options : DEFAULT_STATUS_COLUMNS;
    }
  } catch (err) {
    console.warn("读取项目阶段配置失败", err);
  }
  return DEFAULT_STATUS_COLUMNS;
};

const persistStatusOptions = (options: string[]) => {
  localStorage.setItem(PROJECT_STATUS_STORAGE_KEY, JSON.stringify(options));
};

const mergeStatusOptions = (baseOptions: string[], projectList: Project[]) => {
  return dedupeStatusOptions([
    ...baseOptions,
    ...projectList.map(project => project.status).filter(Boolean),
  ]);
};

const normalizeProjectName = (value: string) => value.trim().replace(/\s+/g, " ").toLocaleLowerCase();

export default function ProjectBoard({ onBack, onOpenCalc }: ProjectBoardProps) {
  const {
    currentWorkspace,
    isWorkspaceReady,
    isLoading: workspaceLoading,
    selectAndCreateWorkspace,
    scanAndImportAllWorkspaceCalculations,
  } = useWorkspaceStore();
  const markDirty = useSaveStore(state => state.markDirty);
  const clearDirty = useSaveStore(state => state.clearDirty);
  const registerSaveHandler = useSaveStore(state => state.registerSaveHandler);
  const unregisterSaveHandler = useSaveStore(state => state.unregisterSaveHandler);
  const replaceAiBusinessData = useAiContextStore(state => state.replaceBusinessData);
  const { confirmOrSave } = useUnsavedChangesGuard();
  const [showWorkspaceOverview, setShowWorkspaceOverview] = useState(false);
  const [projects, setProjects] = useState<Project[]>([]);
  const [noteDrafts, setNoteDrafts] = useState<Record<string, string>>({});
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [projectStageFilter, setProjectStageFilter] = useState<string>("全部");
  const [projectSearchTerm, setProjectSearchTerm] = useState("");
  const [statusOptions, setStatusOptions] = useState<string[]>(readStoredStatusOptions);
  const [showStatusManager, setShowStatusManager] = useState(false);
  const [statusDrafts, setStatusDrafts] = useState<string[]>([]);
  const [newStatusDraft, setNewStatusDraft] = useState("");

  // View Mode: list or grid
  const [viewMode, setViewMode] = useState<"list" | "grid">(
    () => (localStorage.getItem("lamber_project_board_view_mode") as "list" | "grid") || "list"
  );

  // Density Mode: original, standard, compact
  const [densityMode, setDensityMode] = useState<"original" | "standard" | "compact">(
    () => (localStorage.getItem("lamber_project_board_density_mode") as "original" | "standard" | "compact") || "compact"
  );

  // Drawer Ref for Outside Clicks
  const drawerRef = useRef<HTMLDivElement>(null);

  // Creation State
  const [showCreateModal, setShowCreateModal] = useState(false);
  const [newProjectName, setNewProjectName] = useState("");
  const [newCustomerName, setNewCustomerName] = useState("");
  const [newProjectType, setNewProjectType] = useState<ProjectType>("ict");
  const [newProjectPresetId, setNewProjectPresetId] = useState("");
  const [projectPresetTemplates, setProjectPresetTemplates] = useState<ProjectPresetTemplate[]>([]);


  // Import Scanner State
  const [showImportModal, setShowImportModal] = useState(false);
  const [importParentPath, setImportParentPath] = useState("");
  const [importCandidates, setImportCandidates] = useState<ImportCandidate[]>([]);
  const [selectedCandidates, setSelectedCandidates] = useState<Record<string, boolean>>({});
  const [conflictActions, setConflictActions] = useState<Record<string, "merge" | "new" | "skip">>({});
  const [expandedCandidates, setExpandedCandidates] = useState<Record<string, boolean>>({});
  const [importLoading, setImportLoading] = useState(false);
  const [scanLoading, setScanLoading] = useState(false);
  const [globalImportLoading, setGlobalImportLoading] = useState(false);

  const handleScanAndImportAllCalculations = async () => {
    if (globalImportLoading) return;
    setGlobalImportLoading(true);
    try {
      const count = await scanAndImportAllWorkspaceCalculations();
      alert(`刷新完成！已为工作区内 ${count} 个符合条件的项目自动导入测算方案。`);
      fetchProjects();
    } catch (err: any) {
      alert(`导入失败: ${err?.message || err}`);
    } finally {
      setGlobalImportLoading(false);
    }
  };

  // Details Modal State
  const [selectedProject, setSelectedProject] = useState<Project | null>(null);
  const [detailTab, setDetailTab] = useState<'info' | 'files'>('info');
  const [schemes, setSchemes] = useState<BenefitAnalysisScheme[]>([]);
  const [selectedScheme, setSelectedScheme] = useState<BenefitAnalysisScheme | null>(null);
  const [snapshots, setSnapshots] = useState<BenefitAnalysisSnapshot[]>([]);
  const [editingProjectName, setEditingProjectName] = useState("");
  const [editingCustomerName, setEditingCustomerName] = useState("");
  const [editingProjectType, setEditingProjectType] = useState<ProjectType>("ict");
  const [editingStatus, setEditingStatus] = useState("");
  const [isNewSchemeModalOpen, setIsNewSchemeModalOpen] = useState(false);
  const [newSchemeName, setNewSchemeName] = useState("");

  useEffect(() => {
    if (!isWorkspaceReady || !currentWorkspace) {
      replaceAiBusinessData("project_board.core", {
        workspaceReady: false,
        workspaceId: null,
        workspaceName: null,
        projectCount: 0,
        projects: [],
        selectedProject: null,
      });
      return;
    }

    replaceAiBusinessData("project_board.core", {
      workspaceReady: true,
      workspaceId: currentWorkspace.workspaceId,
      workspaceName: currentWorkspace.workspaceName,
      projectCount: projects.length,
      projects: projects.slice(0, 50).map(project => ({
        id: project.id,
        name: project.name,
        customerName: project.customer_name,
        projectType: project.project_type,
        status: project.status,
        benefitStatus: project.benefit_status,
        updatedAt: project.updated_at,
        marginRate: project.summary_metrics?.margin_rate ?? null,
        npv: project.summary_metrics?.npv ?? null,
        npvRate: project.summary_metrics?.npv_rate ?? null,
        riskLevel: project.summary_metrics?.risk_level ?? null,
      })),
      truncatedProjectCount: Math.max(projects.length - 50, 0),
      selectedProject: selectedProject ? {
        id: selectedProject.id,
        name: editingProjectName || selectedProject.name,
        customerName: editingCustomerName || selectedProject.customer_name,
        projectType: editingProjectType,
        status: editingStatus || selectedProject.status,
        note: noteDrafts[selectedProject.id] ?? selectedProject.note ?? "",
        progress: selectedProject.progress ?? 0,
        deadline: selectedProject.deadline ?? null,
      } : null,
    });
  }, [
    isWorkspaceReady,
    currentWorkspace,
    projects,
    selectedProject,
    editingProjectName,
    editingCustomerName,
    editingProjectType,
    editingStatus,
    noteDrafts,
    replaceAiBusinessData,
  ]);

  useEffect(() => {
    if (isWorkspaceReady) {
      setShowWorkspaceOverview(false);
      setSelectedProject(null);
      setSelectedScheme(null);
      setSchemes([]);
      setSnapshots([]);
      fetchProjects();
    } else {
      setProjects([]);
      setSelectedProject(null);
      setSelectedScheme(null);
      setSchemes([]);
      setSnapshots([]);
      setLoading(false);
    }
  }, [isWorkspaceReady, currentWorkspace?.workspaceId]);

  useEffect(() => {
    if (projectStageFilter !== "全部" && !statusOptions.includes(projectStageFilter)) {
      setProjectStageFilter("全部");
    }
  }, [projectStageFilter, statusOptions]);

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
        if (showStatusManager) {
          setShowStatusManager(false);
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
  }, [selectedProject, isNewSchemeModalOpen, showCreateModal, showStatusManager]);

  const fetchProjects = async () => {
    setLoading(true);
    try {
      const workspaceProjects = await projectService.listWorkspaceProjects();
      const projs = workspaceProjects.map(wp => ({
        ...wp.project,
        directoryExists: wp.directoryExists
      }));
      setProjects(projs);
      setStatusOptions(prev => {
        const next = mergeStatusOptions(prev, projs);
        persistStatusOptions(next);
        return next;
      });
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

  const projectNameExists = (name: string, excludeProjectId?: string) => {
    const normalized = normalizeProjectName(name);
    return projects.some(project => {
      if (project.id === excludeProjectId) return false;
      return normalizeProjectName(project.name) === normalized;
    });
  };

  useEffect(() => {
    registerSaveHandler("project-detail", async () => {
      const patches = new Map<string, ProjectDetailPatch>();
      const ensurePatch = (projectId: string) => {
        const existing = patches.get(projectId);
        if (existing) return existing;
        const next: ProjectDetailPatch = {};
        patches.set(projectId, next);
        return next;
      };

      projects.forEach(project => {
        const draftNote = noteDrafts[project.id];
        if (draftNote !== undefined && (project.note || "") !== draftNote) {
          ensurePatch(project.id).note = draftNote;
        }
      });

      if (selectedProject?.id) {
        const nextName = editingProjectName.trim();
        if (!nextName) {
          return { success: false, savedScopes: [], error: "项目名称不能为空" };
        }
        const normalizedNextName = normalizeProjectName(nextName);
        const duplicated = projects.some(project =>
          project.id !== selectedProject.id && normalizeProjectName(project.name) === normalizedNextName
        );
        if (duplicated) {
          return { success: false, savedScopes: [], error: `项目名称「${nextName}」已存在` };
        }

        const nextCustomerName = editingCustomerName.trim() || "未知客户";
        const nextStatus = editingStatus.trim() || statusOptions[0] || DEFAULT_STATUS_COLUMNS[0];
        const patch = ensurePatch(selectedProject.id);
        if (nextName !== selectedProject.name) patch.name = nextName;
        if (nextCustomerName !== selectedProject.customer_name) patch.customerName = nextCustomerName;
        if (editingProjectType !== selectedProject.project_type) patch.projectType = editingProjectType;
        if (nextStatus !== selectedProject.status) patch.status = nextStatus;
      }

      const entries = Array.from(patches.entries())
        .filter(([, patch]) => Object.keys(patch).length > 0);
      if (entries.length === 0) {
        return { success: true, savedScopes: ["project-detail"] };
      }

      const savedProjects: Project[] = [];
      for (const [projectId, patch] of entries) {
        const result = await domainSaveService.saveProjectDetail(projectId, patch);
        savedProjects.push(result);
      }

      const savedById = new Map(savedProjects.map(project => [project.id, project]));
      setProjects(prev => prev.map(project => {
        const saved = savedById.get(project.id);
        return saved ? { ...saved, directoryExists: project.directoryExists } : project;
      }));
      setNoteDrafts(prev => {
        const next = { ...prev };
        savedProjects.forEach(project => {
          next[project.id] = project.note || "";
        });
        return next;
      });
      setSelectedProject(current => {
        if (!current) return current;
        const saved = savedById.get(current.id);
        return saved ? { ...saved, directoryExists: current.directoryExists } : current;
      });
      const currentProject = useProjectStore.getState().currentProject;
      if (currentProject) {
        const saved = savedById.get(currentProject.id);
        if (saved) {
          useProjectStore.getState().setCurrentProject({
            ...saved,
            directoryExists: currentProject.directoryExists,
          });
        }
      }
      if (savedProjects.length > 0) {
        setStatusOptions(prev => {
          const next = mergeStatusOptions(prev, savedProjects);
          persistStatusOptions(next);
          return next;
        });
      }

      return { success: true, savedScopes: ["project-detail"] };
    });

    return () => unregisterSaveHandler("project-detail");
  }, [
    editingCustomerName,
    editingProjectName,
    editingProjectType,
    editingStatus,
    noteDrafts,
    projects,
    registerSaveHandler,
    selectedProject,
    statusOptions,
    unregisterSaveHandler,
  ]);

  const openCreateProjectModal = () => {
    setShowCreateModal(true);
    void projectPresetService.list(false)
      .then(setProjectPresetTemplates)
      .catch(error => {
        console.error("Failed to load project presets for project creation", error);
        setProjectPresetTemplates([]);
      });
  };

  const handleCreateProject = async (e: React.FormEvent) => {
    e.preventDefault();
    const projectName = newProjectName.trim();
    if (!projectName) return;
    if (projectNameExists(projectName)) {
      alert(`项目名称「${projectName}」已存在，请换一个名称。`);
      return;
    }

    try {
      const newProj = await projectService.createProjectInWorkspace(
        projectName,
        newCustomerName.trim() || "未知客户",
        newProjectType,
        newProjectPresetId || null,
      );
      const projWithExists = { ...newProj, directoryExists: true };
      setProjects((prev) => [projWithExists, ...prev]);
      setStatusOptions(prev => {
        const next = mergeStatusOptions(prev, [newProj]);
        persistStatusOptions(next);
        return next;
      });
      setNoteDrafts(prev => ({ ...prev, [newProj.id]: newProj.note || "" }));
      setShowCreateModal(false);
      setNewProjectName("");
      setNewCustomerName("");
      setNewProjectType("ict");
      setNewProjectPresetId("");
      // Automatically open the details of the newly created project
      handleOpenDetails(projWithExists);
    } catch (err) {
      console.error(err);
      alert("创建项目失败: " + err);
    }
  };

  const handleOpenImportScanner = async () => {
    try {
      const selected = await projectFileService.selectLocalFolder();
      if (!selected) return;

      setImportParentPath(selected);
      setShowImportModal(true);
      setScanLoading(true);
      setImportCandidates([]);
      
      const candidates = await invoke<ImportCandidate[]>("scan_import_candidates", { parentPath: selected });
      setImportCandidates(candidates);
      
      // Initialize selections and conflict actions
      const initialSelected: Record<string, boolean> = {};
      const initialConflicts: Record<string, "merge" | "new" | "skip"> = {};
      candidates.forEach(c => {
        initialSelected[c.folderPath] = true;
        if (c.existsConflict) {
          initialConflicts[c.folderPath] = "merge";
        }
      });
      setSelectedCandidates(initialSelected);
      setConflictActions(initialConflicts);
      setScanLoading(false);
    } catch (err: any) {
      console.error(err);
      alert("扫描导入目录失败: " + err);
      setScanLoading(false);
      setShowImportModal(false);
    }
  };

  const handleExecuteImport = async () => {
    setImportLoading(true);
    try {
      const selections = importCandidates
        .filter(c => selectedCandidates[c.folderPath])
        .map(c => {
          let conflictAction = "merge";
          if (c.existsConflict) {
            conflictAction = conflictActions[c.folderPath] || "merge";
          }
          return {
            folderPath: c.folderPath,
            conflictAction
          };
        })
        .filter(sel => sel.conflictAction !== "skip");

      if (selections.length === 0) {
        alert("未选择任何要导入的项目目录。");
        setImportLoading(false);
        return;
      }

      await invoke("execute_bulk_import", { selections });
      alert("批量导入完成！");
      setShowImportModal(false);
      await fetchProjects();
    } catch (err: any) {
      console.error(err);
      alert("批量导入失败: " + err);
    } finally {
      setImportLoading(false);
    }
  };

  const handleCloseCreateModal = () => {
    setShowCreateModal(false);
    setNewProjectName("");
    setNewCustomerName("");
    setNewProjectType("ict");
    setNewProjectPresetId("");
  };

  const handleOpenDetails = async (project: Project) => {
    if (selectedProject?.id && selectedProject.id !== project.id) {
      const canProceed = await confirmOrSave();
      if (!canProceed) return;
    }
    setSelectedProject(project);
    useProjectStore.getState().setCurrentProject(project);
    setDetailTab('info');
    setEditingProjectName(project.name);
    setEditingCustomerName(project.customer_name);
    setEditingProjectType(project.project_type);
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

  const handleSchemeStageChange = async (
    scheme: BenefitAnalysisScheme,
    stage: SchemeStage | null
  ) => {
    if (!selectedProject) return;
    // 点击已选中的阶段则取消标注（切换为未标注）。
    const nextStage = (scheme.stage ?? null) === stage ? null : stage;
    try {
      const updated = await projectService.updateSchemeStage(
        selectedProject.id,
        scheme.id,
        nextStage
      );
      setSchemes(prev => prev.map(s => (s.id === updated.id ? updated : s)));
      setSelectedScheme(prev => (prev && prev.id === updated.id ? updated : prev));
    } catch (err) {
      console.error(err);
      alert("更新甄选阶段失败: " + err);
    }
  };

  const handleUpdateProjectDetails = async () => {
    if (!selectedProject) return;
    const nextProjectName = editingProjectName.trim();
    if (!nextProjectName) return;
    if (projectNameExists(nextProjectName, selectedProject.id)) {
      alert(`项目名称「${nextProjectName}」已存在，请换一个名称。`);
      return;
    }

    const updated: Project = {
      ...selectedProject,
      name: nextProjectName,
      customer_name: editingCustomerName.trim() || "未知客户",
      project_type: editingProjectType,
      status: editingStatus.trim() || statusOptions[0] || DEFAULT_STATUS_COLUMNS[0],
      updated_at: new Date().toISOString()
    };

    try {
      const result = await domainSaveService.saveProjectDetail(updated.id, {
        name: updated.name,
        customerName: updated.customer_name,
        projectType: updated.project_type,
        status: updated.status,
        progress: updated.progress || 0,
        deadline: updated.deadline || null,
        note: updated.note || null,
      });
      const projWithExists = { ...result, directoryExists: selectedProject.directoryExists };
      setSelectedProject(projWithExists);
      useProjectStore.getState().setCurrentProject(projWithExists);
      setProjects(prev => prev.map(p => p.id === result.id ? projWithExists : p));
      clearDirty("project-detail");
      setStatusOptions(prev => {
        const next = mergeStatusOptions(prev, [result]);
        persistStatusOptions(next);
        return next;
      });
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
      if (useProjectStore.getState().currentProject?.id === projectId) {
        useProjectStore.getState().clearCurrentProject();
      }
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
      const workspaceProjects = await projectService.listWorkspaceProjects();
      const latestInfo = workspaceProjects.find(wp => wp.project.id === selectedProject.id);
      
      const projs = workspaceProjects.map(wp => ({
        ...wp.project,
        directoryExists: wp.directoryExists
      }));
      setProjects(projs);

      if (!latestInfo) {
        setSelectedProject(null);
        useProjectStore.getState().clearCurrentProject();
        return;
      }

      const latestProject = {
        ...latestInfo.project,
        directoryExists: latestInfo.directoryExists
      };

      setSelectedProject(latestProject);
      useProjectStore.getState().setCurrentProject(latestProject);
      setEditingProjectName(latestProject.name);
      setEditingCustomerName(latestProject.customer_name);
      setEditingProjectType(latestProject.project_type);
      setEditingStatus(latestProject.status);
      setNoteDrafts(prev => ({ ...prev, [latestProject.id]: latestProject.note || "" }));
      await reloadSchemesForProject(latestProject);
    } catch (err) {
      console.error("刷新项目详情失败", err);
    }
  };

  const handleOpenStatusManager = () => {
    setStatusDrafts(statusOptions);
    setNewStatusDraft("");
    setShowStatusManager(true);
  };

  const handleAddStatusDraft = () => {
    const nextStatus = newStatusDraft.trim();
    if (!nextStatus) return;
    if (statusDrafts.some(status => status.trim() === nextStatus)) {
      alert(`阶段「${nextStatus}」已存在。`);
      return;
    }
    setStatusDrafts(prev => [...prev, nextStatus]);
    setNewStatusDraft("");
  };

  const handleSaveStatusManager = async () => {
    const nextOptions = dedupeStatusOptions(statusDrafts);
    if (nextOptions.length === 0) {
      alert("至少保留一个项目阶段。");
      return;
    }

    const renamePairs = statusOptions
      .map((oldStatus, index) => ({ from: oldStatus, to: statusDrafts[index]?.trim() || "" }))
      .filter(pair => pair.to && pair.from !== pair.to);

    try {
      let updatedProjects = projects;
      for (const pair of renamePairs) {
        const affectedProjects = updatedProjects.filter(project => project.status === pair.from);
        if (affectedProjects.length === 0) continue;
        const savedProjects = await Promise.all(
          affectedProjects.map(project => projectService.updateProject({
            ...project,
            status: pair.to,
            updated_at: new Date().toISOString(),
          }))
        );
        updatedProjects = updatedProjects.map(project => {
          return savedProjects.find(saved => saved.id === project.id) || project;
        });
      }

      const finalOptions = mergeStatusOptions(nextOptions, updatedProjects);
      persistStatusOptions(finalOptions);
      setStatusOptions(finalOptions);
      setProjects(updatedProjects);
      if (selectedProject) {
        const updatedSelected = updatedProjects.find(project => project.id === selectedProject.id);
        if (updatedSelected) {
          setSelectedProject(updatedSelected);
          setEditingStatus(updatedSelected.status);
        }
      }
      setShowStatusManager(false);
    } catch (err) {
      console.error(err);
      alert("保存项目阶段失败: " + err);
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
      const canProceed = await confirmOrSave();
      if (!canProceed) return;
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
    const project = projects.find(candidate => candidate.id === projectId);
    if (project) {
      useProjectStore.getState().setCurrentProject(project);
    }
    markDirty("project-detail");
  };

  const handleProjectNoteBlur = async (project: Project) => {
    const nextNote = noteDrafts[project.id] ?? "";
    if ((project.note || "") === nextNote) return;

    try {
      const updatedProject = await domainSaveService.saveProjectDetail(project.id, {
        note: nextNote,
      });
      setProjects(prev => prev.map(p => p.id === updatedProject.id ? updatedProject : p));
      setNoteDrafts(prev => ({ ...prev, [updatedProject.id]: updatedProject.note || "" }));
      if (selectedProject?.id === updatedProject.id) {
        setSelectedProject(updatedProject);
      }
      clearDirty("project-detail");
    } catch (err) {
      console.error("保存项目备注失败", err);
      alert("保存项目备注失败: " + err);
      setNoteDrafts(prev => ({ ...prev, [project.id]: project.note || "" }));
    }
  };

  const renderProjectNote = (project: Project, compact = false) => (
    <div
      className="flex items-start gap-2.5"
      onClick={(e) => e.stopPropagation()}
      onMouseDown={(e) => e.stopPropagation()}
    >
      <div className="mt-0.5 rounded-md bg-warning-soft p-1 text-warning">
        <StickyNote className="h-4 w-4" />
      </div>
      <div className="min-w-0 flex-1">
        <span className="block text-caption font-bold text-foreground">项目备注</span>
        <textarea
          value={noteDrafts[project.id] ?? project.note ?? ""}
          onChange={(e) => handleProjectNoteChange(project.id, e.target.value)}
          rows={compact ? 2 : 3}
          placeholder="填写客户背景、推进风险、下一步动作..."
          className={`mt-1 block w-full resize-none rounded-xl border border-border bg-muted px-3 py-2 text-caption leading-5 text-secondary-foreground outline-none transition-colors placeholder:text-muted-foreground/60 focus:border-ring focus:bg-card focus:ring-2 focus:ring-ring/20 ${
            compact ? "min-h-[48px]" : "min-h-[58px]"
          }`}
        />
      </div>
    </div>
  );

  const getStatusBadge = (status: Project["benefit_status"]) => {
    switch (status) {
      case "normal":
        return <span className="inline-flex items-center gap-1 text-[10px] bg-success-soft text-success px-2 py-0.5 rounded-md font-bold border border-success-soft"><span className="w-1.5 h-1.5 bg-success rounded-full" /> 测算已更新</span>;
      case "outdated":
        return <span className="inline-flex items-center gap-1 text-[10px] bg-warning-soft text-warning px-2 py-0.5 rounded-md font-bold border border-warning-soft"><span className="w-1.5 h-1.5 bg-warning rounded-full" /> 测算已失效</span>;
      default:
        return <span className="inline-flex items-center gap-1 text-[10px] bg-muted text-muted-foreground px-2 py-0.5 rounded-md font-bold border border-border">未测算</span>;
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

  const getRiskTone = (level?: string) => {
    if (level === "高风险") {
      return {
        badge: "border-destructive-soft bg-destructive-soft text-destructive",
        dot: "bg-destructive",
      };
    }
    if (level === "中风险") {
      return {
        badge: "border-warning-soft bg-warning-soft text-warning",
        dot: "bg-warning",
      };
    }
    return {
      badge: "border-success-soft bg-success-soft text-success",
      dot: "bg-success",
    };
  };

  const getRiskBorderStyles = (level?: string) => {
    if (level === "高风险") {
      return "border-l-4 border-l-destructive";
    }
    if (level === "中风险") {
      return "border-l-4 border-l-warning";
    }
    return "border-l-4 border-l-success";
  };

  const getRiskTopBorderStyles = (level?: string) => {
    if (level === "高风险") {
      return "border-t-4 border-t-destructive";
    }
    if (level === "中风险") {
      return "border-t-4 border-t-warning";
    }
    return "border-t-4 border-t-success";
  };

  const renderMetricValue = (
    value: string,
    options: { unit?: string; prefix?: string; compact?: boolean } = {}
  ) => {
    const { unit, prefix, compact = false } = options;
    const hasUnit = Boolean(unit && value.endsWith(unit));
    const displayValue = hasUnit && unit ? value.slice(0, -unit.length) : value;
    const isEmpty = displayValue === "--";

    return (
      <div className="flex min-w-0 items-baseline gap-1 numeric-value">
        {prefix && !isEmpty && <span className="text-[10px] font-bold text-muted-foreground">{prefix}</span>}
        <span className={`${compact ? "text-xl" : "text-2xl"} min-w-0 truncate font-extrabold tracking-tight text-foreground`}>
          {displayValue}
        </span>
        {unit && !isEmpty && <span className="text-xs font-bold text-muted-foreground">{unit}</span>}
      </div>
    );
  };

  const renderMetricTile = (
    label: string,
    value: string,
    options: { unit?: string; prefix?: string; compact?: boolean } = {}
  ) => (
    <div className={`${options.compact ? "p-3" : "p-3.5"} min-w-0 rounded-xl border border-border bg-card transition-colors hover:border-primary/50`}>
      <span className="mb-1 block truncate text-caption font-semibold text-muted-foreground">{label}</span>
      {renderMetricValue(value, options)}
    </div>
  );

  const renderRiskAssessment = (metrics: SummaryMetrics, compact = false) => {
    const tone = getRiskTone(metrics.risk_level);
    return (
      <div className={`${compact ? "mt-3 px-3 py-2" : "mt-4 px-4 py-2.5"} flex items-center justify-between gap-3 rounded-xl border border-border bg-muted/50`}>
        <span className="flex min-w-0 items-center gap-1.5 text-caption font-semibold text-muted-foreground">
          <Info className="h-3.5 w-3.5 shrink-0 text-muted-foreground/60" />
          <span className="truncate">风险综合评估</span>
        </span>
        <span className={`${tone.badge} inline-flex shrink-0 items-center gap-1.5 rounded-lg border px-3 py-1 text-caption font-bold`}>
          <span className={`${tone.dot} h-1.5 w-1.5 rounded-full`} />
          {metrics.risk_level}
        </span>
      </div>
    );
  };

  const renderMetricPanel = (metrics: SummaryMetrics | null | undefined, compact = false) => {
    if (!metrics) {
      return (
        <div className="rounded-xl border border-dashed border-border bg-card/70 px-4 py-6 text-center text-caption font-medium leading-5 text-muted-foreground">
          暂无效益分析指标，点击下方按钮开始测算
        </div>
      );
    }

    return (
      <>
        <div className={`grid grid-cols-2 ${compact ? "gap-2.5" : "gap-3"}`}>
          {renderMetricTile("毛利率", formatMetricPercent(metrics.margin_rate), { unit: "%", compact })}
          {renderMetricTile("净现值 NPV", formatMetricNumber(metrics.npv), { prefix: "¥", compact })}
          {renderMetricTile("净现值率 NPVR", formatMetricPercent(metrics.npv_rate), { unit: "%", compact })}
          {renderMetricTile("内部收益率 IRR", formatMetricPercent(metrics.irr), { unit: "%", compact })}
        </div>
        {renderRiskAssessment(metrics, compact)}
      </>
    );
  };

  const renderProjectCardHeader = (project: Project, compact = false) => (
    <div className={`${compact ? "p-4" : "p-5"} border-b border-border`}>
      <div className="flex items-start justify-between gap-3">
        <div className="flex min-w-0 items-start gap-3">
          <div className="rounded-xl border border-border bg-muted p-2.5 text-secondary-foreground transition-colors group-hover:bg-primary-soft group-hover:text-primary">
            <FileText className="h-5 w-5" />
          </div>
          <div className="min-w-0">
            <h3 className={`${compact ? "text-base" : "text-lg"} truncate font-extrabold leading-snug tracking-tight text-foreground`} title={project.name}>
              {project.name}
            </h3>
            <p className="mt-1 truncate text-caption leading-5 text-muted-foreground">
              客户: <span className="font-medium text-secondary-foreground">{project.customer_name || "未填写"}</span>
            </p>
          </div>
        </div>

        <div className="flex shrink-0 flex-col items-end gap-1.5">
          <span className="rounded-md bg-muted px-2 py-0.5 text-[10px] font-bold text-muted-foreground">
            {project.status}
          </span>
          {renderProjectTypeBadge(project)}
          {getStatusBadge(project.benefit_status)}
          {project.directoryExists === false && (
            <span className="rounded-md bg-destructive-soft border border-destructive-soft px-2 py-0.5 text-[10px] font-bold text-destructive flex items-center gap-1 animate-pulse" title="在磁盘中找不到项目对应的文件夹">
              <AlertTriangle className="h-3 w-3 shrink-0" />
              目录缺失
            </span>
          )}
        </div>
      </div>
    </div>
  );

  const renderOpenCalcButton = (project: Project, id: string, compact = false) => (
    <button
      id={id}
      onClick={async (e) => {
        e.stopPropagation();
        const canProceed = await confirmOrSave();
        if (!canProceed) return;
        onOpenCalc(project.id, project.default_scheme_id || null);
      }}
      className={`${compact ? "px-3 py-1.5" : "px-4 py-2"} group inline-flex shrink-0 items-center justify-center rounded-xl bg-primary-soft hover:bg-primary text-primary hover:text-primary-foreground border border-primary-soft text-caption font-bold shadow-sm transition-all active:scale-[0.98] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/20`}
    >
      <BarChart3 className="h-4 w-4" />
      <span className="ml-2">ICT 测算</span>
      <ArrowRight className="ml-1.5 h-4 w-4 transition-transform duration-200 group-hover:translate-x-0.5" />
    </button>
  );

  const renderProjectTypeBadge = (project: Project) => (
    <span className={`rounded-md px-2 py-0.5 text-[10px] font-bold ${
      project.project_type === "intelligent_compute"
        ? "bg-primary-soft text-primary"
        : "bg-muted text-muted-foreground"
    }`}>
      {project.project_type === "intelligent_compute" ? "智算项目" : "ICT 项目"}
    </span>
  );

  const renderOpenIntelligentComputeButton = (project: Project, id: string, compact = false) => {
    if (project.project_type !== "intelligent_compute") return null;
    return (
      <button
        id={id}
        onClick={async event => {
          event.stopPropagation();
          const canProceed = await confirmOrSave();
          if (!canProceed) return;
          useProjectStore.getState().setCurrentProject(project);
          useNavigationStore.getState().navigateTo("ai_compute_quote", project.id);
        }}
        className={`${compact ? "px-3 py-1.5" : "px-4 py-2"} inline-flex shrink-0 items-center justify-center rounded-xl bg-muted text-caption font-bold text-secondary-foreground transition-all hover:bg-primary-soft hover:text-primary active:scale-[0.98]`}
      >
        <AppIcon name="calculator" size={16} />
        <span className="ml-2">智算测算</span>
      </button>
    );
  };

  const renderOpenFolderButton = (project: Project, compact = false) => {
    const hasFolder = !!project.folder_path;
    return (
      <button
        onClick={async (e) => {
          e.stopPropagation();
          if (!hasFolder) {
            alert("该项目尚未绑定项目文件夹，请在项目详情中设置。");
            return;
          }
          try {
            await projectFileService.openProjectFolder(project.id);
          } catch (err) {
            console.error("打开项目文件夹失败:", err);
            alert("无法打开项目文件夹: " + err);
          }
        }}
        className={`group inline-flex shrink-0 items-center justify-center rounded-xl border border-border bg-card hover:bg-muted text-secondary-foreground text-caption font-bold transition-all active:scale-[0.98] ${
          compact ? "p-1.5" : "px-3 py-2"
        } ${!hasFolder ? "opacity-40 cursor-not-allowed" : ""}`}
        title={hasFolder ? `打开项目文件夹: ${project.folder_path}` : "未绑定项目文件夹"}
      >
        <FolderOpen className="h-4 w-4" />
        {!compact && <span className="ml-1.5">打开文件夹</span>}
      </button>
    );
  };

  const matchesProjectSearch = (project: Project) => {
    const keyword = projectSearchTerm.trim().toLocaleLowerCase();
    if (!keyword) return true;
    return [
      project.name,
      project.customer_name,
      project.status,
      project.note || "",
    ].some(value => String(value || "").toLocaleLowerCase().includes(keyword));
  };

  const filteredProjects = projects.filter(project => {
    const matchesStage = projectStageFilter === "全部" || project.status === projectStageFilter;
    return matchesStage && matchesProjectSearch(project);
  });

  const getStageCount = (stage: string) => {
    return projects.filter(project => {
      const matchesStage = stage === "全部" || project.status === stage;
      return matchesStage && matchesProjectSearch(project);
    }).length;
  };

  const renderCreateProjectEntry = (mode: "list" | "grid") => {
    let heightClass = "min-h-[360px]";
    if (mode === "list") {
      heightClass = "min-h-[104px] w-full rounded-xl p-4";
    } else {
      if (densityMode === "standard") {
        heightClass = "min-h-[320px] rounded-xl p-4";
      } else if (densityMode === "compact") {
        heightClass = "min-h-[220px] rounded-xl p-4";
      } else {
        heightClass = "min-h-[360px] rounded-2xl p-5";
      }
    }

    return (
      <button
        id={`board_create_project_entry_${mode}`}
        type="button"
        onClick={openCreateProjectModal}
        className={`group flex items-center justify-center gap-3 border-2 border-dashed border-border bg-card text-secondary-foreground shadow-sm transition-all hover:border-primary/50 hover:bg-primary-soft hover:text-primary active:scale-[0.99] ${heightClass}`}
      >
        <span className="flex h-11 w-11 items-center justify-center rounded-xl bg-muted text-muted-foreground shadow-sm transition-all group-hover:scale-105 group-hover:bg-primary group-hover:text-primary-foreground">
          <Plus className="h-5 w-5" />
        </span>
        <span className="text-body font-extrabold">创建新项目</span>
      </button>
    );
  };

  if (!isWorkspaceReady || showWorkspaceOverview) {
    return (
      <div className="flex flex-col flex-1 h-full overflow-hidden bg-background text-foreground animate-in fade-in duration-300">
        <header className="flex items-center justify-between px-6 py-4 shrink-0 bg-card shadow-sm">
          <div className="flex items-center gap-3">
            <button
              id="board_back_btn"
              onClick={onBack}
              className="text-secondary-foreground hover:text-primary hover:bg-secondary font-semibold flex items-center gap-1.5 px-3 py-2 rounded-lg transition-colors text-body"
            >
              <span>←</span> 返回集市
            </button>
            <div>
              <h1 className="text-page-title font-extrabold flex items-center gap-2 text-foreground">
                <AppIcon name="project" size={22} className="text-muted-foreground" /> 项目工作区
              </h1>
              <p className="text-caption text-secondary-foreground mt-0.5">选择工作区后进入对应的项目看板</p>
            </div>
          </div>
        </header>
        <WorkspaceGate
          onBack={isWorkspaceReady ? () => setShowWorkspaceOverview(false) : onBack}
          backLabel={isWorkspaceReady ? "返回当前项目列表" : "返回集市"}
          onCurrentWorkspaceSelected={() => setShowWorkspaceOverview(false)}
          onWorkspaceChanged={() => setShowWorkspaceOverview(false)}
        />
      </div>
    );
  }

  return (
    <div className="flex flex-col flex-1 h-full overflow-hidden bg-background text-foreground animate-in fade-in duration-300">
      <header className="flex items-center justify-between px-6 py-4 border-b border-border shrink-0 bg-card">
        <div className="flex items-center gap-3">
          <button
            id="board_back_btn"
            onClick={onBack}
            className="text-secondary-foreground hover:text-primary hover:bg-secondary font-semibold flex items-center gap-1.5 px-3 py-2 rounded-lg transition-colors text-body"
          >
            <span>←</span> 返回集市
          </button>
          <div>
            <h1 className="text-page-title font-extrabold flex items-center gap-2 text-foreground">
              <AppIcon name="project" size={22} className="text-muted-foreground" /> 项目看板
            </h1>
            <p className="text-caption text-secondary-foreground mt-0.5">管理项目生命周期及其关联的效益分析测算</p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <GlobalSaveButton />
          <button
            onClick={() => useNavigationStore.getState().navigateTo("settings")}
            className="flex h-9 w-9 items-center justify-center rounded-lg border border-border bg-card text-secondary-foreground hover:bg-secondary hover:text-foreground transition-all shadow-sm"
            title="系统设置"
          >
            <AppIcon name="settings" size={18} />
          </button>
          <button
            onClick={handleOpenImportScanner}
            className="inline-flex items-center gap-1.5 bg-secondary hover:bg-secondary/80 text-secondary-foreground border border-input font-bold px-4 py-2 rounded-lg text-body transition-all shadow-sm active:scale-[0.98] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/20"
          >
            <FolderPlus className="h-4 w-4" /> 批量扫描导入
          </button>
          <button
            id="board_create_project_btn"
            onClick={openCreateProjectModal}
            className="inline-flex items-center gap-1.5 bg-primary text-primary-foreground font-bold px-4 py-2 rounded-lg text-body hover:bg-primary/95 transition-all shadow-sm active:scale-[0.98] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/20"
          >
            <Plus className="h-4 w-4" /> 创建新项目
          </button>
        </div>
      </header>

      <section className="shrink-0 bg-muted/35 px-6 py-3">
        <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
          <div className="min-w-0">
            <div className="text-caption font-extrabold uppercase tracking-wide text-secondary-foreground">项目工作区</div>
            <div className="mt-1 flex min-w-0 items-center gap-2">
              <span className="shrink-0 rounded-md bg-primary-soft px-2 py-1 text-caption font-extrabold text-primary">
                {currentWorkspace?.workspaceName || "当前工作区"}
              </span>
              <span className="truncate font-mono text-caption text-secondary-foreground">
                {currentWorkspace?.workspaceRoot}
              </span>
            </div>
          </div>
          <div className="flex shrink-0 gap-2">
            <button
              type="button"
              disabled={workspaceLoading || loading || globalImportLoading}
              onClick={handleScanAndImportAllCalculations}
              className="rounded-md bg-card px-3 py-2 text-caption font-bold text-foreground shadow-sm disabled:opacity-50 flex items-center gap-1.5 hover:bg-muted transition-all duration-200"
            >
              <AppIcon
                name={globalImportLoading ? "loading" : "reverse"}
                size={13}
                className={`${globalImportLoading ? "animate-spin text-primary" : "text-muted-foreground"}`}
              />
              {globalImportLoading ? "正在导入全区测算..." : "刷新全区测算"}
            </button>
            <button
              type="button"
              disabled={workspaceLoading}
              onClick={async () => {
                const canProceed = await confirmOrSave();
                if (canProceed) setShowWorkspaceOverview(true);
              }}
              className="rounded-md bg-card px-3 py-2 text-caption font-bold text-foreground shadow-sm disabled:opacity-50 hover:bg-muted transition-all duration-200"
            >
              切换工作区
            </button>
            <button
              type="button"
              disabled={workspaceLoading}
              onClick={async () => {
                const canProceed = await confirmOrSave();
                if (canProceed) selectAndCreateWorkspace();
              }}
              className="rounded-md bg-primary px-3 py-2 text-caption font-bold text-primary-foreground shadow-sm disabled:opacity-50 transition-all duration-200"
            >
              新建工作区
            </button>
          </div>
        </div>
      </section>

      {/* Kanban Board Container */}
      {loading ? (
        <div className="flex-1 flex flex-col items-center justify-center gap-3 text-secondary-foreground text-body">
          <AppIcon name="loading" size={36} className="animate-spin text-primary" />
          <span className="text-body font-semibold">正在载入项目库...</span>
        </div>
      ) : error ? (
        <div className="flex-1 flex flex-col items-center justify-center gap-2 text-red-500 text-body">
          <AppIcon name="error" size={40} />
          <span className="font-bold">{error}</span>
          <button onClick={fetchProjects} className="mt-2 text-body text-primary underline">重试</button>
        </div>
      ) : (
        <>
          {/* Filter Bar */}
          <div className="px-6 py-4 bg-muted/20 border-b border-border/80 flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between shrink-0">
            <div className="flex items-center gap-2 overflow-x-auto pb-1 md:pb-0 scrollbar-none">
              <span className="text-caption text-secondary-foreground font-extrabold mr-2 uppercase tracking-wider shrink-0">项目阶段筛选:</span>
              {["全部", ...statusOptions].map((stage) => {
                const isActive = projectStageFilter === stage;
                const count = getStageCount(stage);

                return (
                  <button
                    key={stage}
                    onClick={() => setProjectStageFilter(stage)}
                    className={`px-3.5 py-1.5 rounded-full text-caption font-bold transition-all flex items-center gap-1.5 shrink-0 active:scale-[0.98] ${
                      isActive
                        ? "bg-primary text-primary-foreground shadow-sm font-extrabold"
                        : "bg-card border border-border text-secondary-foreground hover:bg-muted hover:text-foreground"
                    }`}
                  >
                    {stage}
                    <span className={`text-[10px] px-1.5 py-0.5 rounded-full ${
                      isActive ? "bg-primary-foreground/20 text-primary-foreground" : "bg-muted text-muted-foreground"
                    }`}>
                      {count}
                    </span>
                  </button>
                );
              })}
              <button
                type="button"
                onClick={handleOpenStatusManager}
                className="inline-flex shrink-0 items-center gap-1.5 rounded-full border border-border bg-card px-3 py-1.5 text-caption font-bold text-secondary-foreground transition-all hover:bg-muted hover:text-foreground"
                title="管理项目阶段"
              >
                <Settings2 className="h-3.5 w-3.5" />
                管理阶段
              </button>
            </div>

            <div className="flex items-center gap-2 shrink-0">
              <div className="relative w-full min-w-[220px] lg:w-72">
                <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                <input
                  type="search"
                  value={projectSearchTerm}
                  onChange={(event) => setProjectSearchTerm(event.target.value)}
                  placeholder="搜索项目、客户或备注..."
                  className="w-full rounded-lg border border-input bg-card py-2 pl-9 pr-9 text-caption font-semibold text-foreground outline-none transition-all focus:border-ring focus:ring-2 focus:ring-ring/20"
                />
                {projectSearchTerm && (
                  <button
                    type="button"
                    onClick={() => setProjectSearchTerm("")}
                    className="absolute right-2 top-1/2 rounded-md p-1 -translate-y-1/2 text-muted-foreground transition-all hover:bg-muted hover:text-foreground"
                    title="清空搜索"
                  >
                    <X className="h-3.5 w-3.5" />
                  </button>
                )}
              </div>

              <div className="flex items-center border border-border bg-card rounded-lg p-0.5 shrink-0">
                <button
                  onClick={() => {
                    setDensityMode("original");
                    localStorage.setItem("lamber_project_board_density_mode", "original");
                  }}
                  className={`px-2 py-1 rounded-md text-[10px] transition-all ${
                    densityMode === "original"
                      ? "bg-muted text-foreground font-extrabold shadow-sm"
                      : "text-secondary-foreground hover:text-foreground opacity-60 hover:opacity-100 font-bold"
                  }`}
                  title="低密度 - 原版高卡"
                >
                  低密度
                </button>
                <button
                  onClick={() => {
                    setDensityMode("standard");
                    localStorage.setItem("lamber_project_board_density_mode", "standard");
                  }}
                  className={`px-2 py-1 rounded-md text-[10px] transition-all ${
                    densityMode === "standard"
                      ? "bg-muted text-foreground font-extrabold shadow-sm"
                      : "text-secondary-foreground hover:text-foreground opacity-60 hover:opacity-100 font-bold"
                  }`}
                  title="中密度 - 标准精致"
                >
                  中密度
                </button>
                <button
                  onClick={() => {
                    setDensityMode("compact");
                    localStorage.setItem("lamber_project_board_density_mode", "compact");
                  }}
                  className={`px-2 py-1 rounded-md text-[10px] transition-all ${
                    densityMode === "compact"
                      ? "bg-muted text-foreground font-extrabold shadow-sm"
                      : "text-secondary-foreground hover:text-foreground opacity-60 hover:opacity-100 font-bold"
                  }`}
                  title="高密度 - 极致紧凑"
                >
                  高密度
                </button>
              </div>

              <div className="flex items-center border border-border bg-card rounded-lg p-0.5 shrink-0">
                <button
                  onClick={() => handleToggleViewMode("list")}
                  className={`p-1.5 rounded-md transition-all ${
                    viewMode === "list"
                      ? "bg-muted text-foreground font-bold shadow-sm"
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
                      ? "bg-muted text-foreground font-bold shadow-sm"
                      : "text-secondary-foreground hover:text-foreground opacity-60 hover:opacity-100"
                  }`}
                  title="卡片视图"
                >
                  <LayoutGrid className="w-4 h-4" />
                </button>
              </div>
            </div>
          </div>

          {/* Project List */}
          {viewMode === "list" ? (
            <div className="flex-1 overflow-y-auto p-6 space-y-4">
              {renderCreateProjectEntry("list")}
              {filteredProjects.length === 0 && (
                <div className="rounded-xl border border-dashed border-border bg-card/60 p-8 text-center text-body font-semibold text-secondary-foreground">
                  {projectSearchTerm ? "没有匹配的项目" : "该阶段下暂无项目"}
                </div>
              )}
              {filteredProjects.map((project) => {
                const metrics = project.summary_metrics;

                // ==================== MODE 1: ORIGINAL (低密度 - 原图大卡片) ====================
                if (densityMode === "original") {
                  return (
                    <div
                      key={project.id}
                      id={`project_card_${project.id}`}
                      onClick={() => handleOpenDetails(project)}
                      className="group relative cursor-pointer overflow-hidden rounded-2xl border border-border bg-card shadow-sm transition-all duration-200 hover:border-border/80 hover:shadow-md animate-in fade-in slide-in-from-bottom-2"
                    >
                      <div className="flex flex-col xl:flex-row">
                        <div className="min-w-0 flex-1">
                          {renderProjectCardHeader(project, true)}
                          <div className="p-4">
                            {renderProjectNote(project, true)}
                          </div>
                        </div>

                        <div className="flex w-full flex-col border-t border-border bg-muted/60 xl:w-[560px] xl:border-l xl:border-t-0">
                          <div className="p-4">
                            {renderMetricPanel(metrics, true)}
                          </div>

                          <div className="flex items-center justify-between gap-4 border-t border-border bg-card/70 px-4 py-3">
                            <span className="shrink-0 text-caption font-medium text-muted-foreground">
                              更新于 {new Date(project.updated_at).toLocaleDateString()}
                            </span>
                            <div className="flex items-center gap-2">
                              {renderOpenFolderButton(project, true)}
                              {renderOpenIntelligentComputeButton(project, `open_intelligent_${project.id}`, true)}
                              {renderOpenCalcButton(project, `open_calc_btn_${project.id}`, true)}
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                }

                // ==================== MODE 2: STANDARD (中密度 - 标准配置) ====================
                if (densityMode === "standard") {
                  return (
                    <div
                      key={project.id}
                      id={`project_card_${project.id}`}
                      onClick={() => handleOpenDetails(project)}
                      className="group relative cursor-pointer overflow-hidden bg-card rounded-2xl border border-border/85 p-4 shadow-sm hover:shadow-md transition-all duration-200 flex flex-col md:flex-row items-stretch gap-4 animate-in fade-in slide-in-from-bottom-2"
                    >
                      {/* Left Block (30% Width): Meta info & notes */}
                      <div className="md:w-[30%] flex flex-col justify-between border-b md:border-b-0 md:border-r border-border pb-3 md:pb-0 md:pr-4">
                        <div className="space-y-2">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className="rounded-md bg-muted px-2 py-0.5 text-[10px] font-bold text-muted-foreground">
                              {project.status}
                            </span>
                            {renderProjectTypeBadge(project)}
                            {getStatusBadge(project.benefit_status)}
                            {project.directoryExists === false && (
                              <span className="rounded-md bg-destructive-soft border border-destructive-soft px-2 py-0.5 text-[10px] font-bold text-destructive flex items-center gap-1 animate-pulse" title="在磁盘中找不到项目对应的文件夹">
                                <AlertTriangle className="h-3 w-3 shrink-0" />
                                目录缺失
                              </span>
                            )}
                          </div>
                          <div>
                            <h3 className="text-body font-black text-foreground tracking-tight truncate" title={project.name}>{project.name}</h3>
                            <p className="text-caption text-muted-foreground mt-0.5 truncate" title={project.customer_name || "未填写"}>
                              客户: {project.customer_name || "未填写"}
                            </p>
                          </div>
                        </div>

                        {/* Note Box */}
                        <div className="mt-3 bg-muted/50 p-2 rounded-lg border border-border/60 flex items-start gap-1.5" onClick={(e) => e.stopPropagation()}>
                          <StickyNote className="h-3.5 w-3.5 text-warning flex-shrink-0 mt-0.5" />
                          <div className="min-w-0 flex-1">
                            <textarea
                              value={noteDrafts[project.id] ?? project.note ?? ""}
                              onChange={(e) => handleProjectNoteChange(project.id, e.target.value)}
                              onBlur={() => handleProjectNoteBlur(project)}
                              rows={2}
                              placeholder="填写备注..."
                              className="w-full resize-none bg-transparent text-caption text-muted-foreground outline-none leading-normal placeholder:text-muted-foreground/60 focus:ring-0"
                            />
                          </div>
                        </div>
                      </div>

                      {/* Middle Block (50% Width): 4 Metrics horizontally */}
                      <div className="md:w-[50%] flex flex-col justify-center px-2">
                        <div className="grid grid-cols-4 gap-2 text-left">
                          <div className="bg-muted/30 p-2.5 rounded-xl border border-border/40">
                            <span className="text-[9px] text-muted-foreground block font-bold uppercase truncate">毛利率</span>
                            <span className="text-body font-black text-foreground truncate block">
                              {metrics ? formatMetricPercent(metrics.margin_rate) : "--"}
                            </span>
                          </div>
                          <div className="bg-muted/30 p-2.5 rounded-xl border border-border/40">
                            <span className="text-[9px] text-muted-foreground block font-bold uppercase truncate">净现值 NPV</span>
                            <span className="text-body font-black text-foreground truncate block">
                              {metrics ? `¥${formatMetricNumber(metrics.npv)}` : "--"}
                            </span>
                          </div>
                          <div className="bg-muted/30 p-2.5 rounded-xl border border-border/40">
                            <span className="text-[9px] text-muted-foreground block font-bold uppercase truncate">NPVR</span>
                            <span className="text-body font-black text-foreground truncate block">
                              {metrics ? formatMetricPercent(metrics.npv_rate) : "--"}
                            </span>
                          </div>
                          <div className="bg-muted/30 p-2.5 rounded-xl border border-border/40">
                            <span className="text-[9px] text-muted-foreground block font-bold uppercase truncate">IRR</span>
                            <span className="text-body font-black text-foreground truncate block">
                              {metrics ? formatMetricPercent(metrics.irr) : "--"}
                            </span>
                          </div>
                        </div>

                        {/* Combined Risk tag */}
                        {metrics && metrics.risk_level && (
                          <div className="mt-2 flex items-center justify-between text-caption text-muted-foreground">
                            <span className="font-semibold">风险评估:</span>
                            <span className={`px-2 py-0.5 font-bold border rounded-md text-[9px] ${getRiskTone(metrics.risk_level).badge}`}>
                              <span className={`w-1 h-1 rounded-full ${getRiskTone(metrics.risk_level).dot} mr-1 inline-block`} />
                              {metrics.risk_level}
                            </span>
                          </div>
                        )}
                      </div>

                      {/* Right Block (20% Width): Action */}
                      <div className="md:w-[20%] border-t md:border-t-0 md:border-l border-border pt-3 md:pt-0 md:pl-4 flex flex-col justify-between items-end gap-2 text-right">
                        <span className="text-[9px] text-muted-foreground">
                          更新于 {new Date(project.updated_at).toLocaleDateString()}
                        </span>
                        <div className="flex items-center gap-2">
                          {renderOpenFolderButton(project, true)}
                          {renderOpenIntelligentComputeButton(project, `open_intelligent_standard_${project.id}`, true)}
                          {renderOpenCalcButton(project, `open_calc_btn_standard_${project.id}`, true)}
                        </div>
                      </div>
                    </div>
                  );
                }

                // ==================== MODE 3: ULTRA-COMPACT (高密度 - 极致对齐单行) ====================
                return (
                  <div
                    key={project.id}
                    id={`project_card_${project.id}`}
                    onClick={() => handleOpenDetails(project)}
                    className={`bg-card rounded-xl border border-border/85 shadow-sm hover:shadow-md transition-all duration-200 flex flex-col lg:flex-row lg:items-center justify-between p-3 gap-3 min-h-[64px] ${getRiskBorderStyles(metrics?.risk_level)} animate-in fade-in slide-in-from-bottom-2`}
                  >
                    {/* Column 1: Project Identity & Stage (26% width) */}
                    <div className="lg:w-[26%] flex items-center gap-2.5 min-w-0">
                      <div className="w-8 h-8 bg-muted rounded-lg flex items-center justify-center text-secondary-foreground border border-border flex-shrink-0">
                        <FileText className="h-4 w-4" />
                      </div>
                      <div className="min-w-0 flex-1">
                        <div className="flex items-center gap-1.5 min-w-0 flex-wrap">
                          <h3 className="text-caption font-black text-foreground truncate leading-tight hover:text-primary cursor-pointer" title={project.name}>
                            {project.name}
                          </h3>
                          <span className="px-1 py-0.5 text-[8px] font-bold bg-muted text-muted-foreground rounded flex-shrink-0">
                            {project.status}
                          </span>
                          {renderProjectTypeBadge(project)}
                          {project.directoryExists === false && (
                            <span className="px-1 py-0.5 text-[8px] font-bold bg-destructive-soft border border-destructive-soft text-destructive rounded flex-shrink-0 flex items-center gap-0.5 animate-pulse" title="在磁盘中找不到项目对应的文件夹">
                              <AlertTriangle className="h-2.5 w-2.5 shrink-0" />
                              目录缺失
                            </span>
                          )}
                        </div>
                        <p className="text-caption text-muted-foreground truncate mt-0.5" title={project.customer_name || "未填写"}>
                          {project.customer_name || "未填写"}
                        </p>
                      </div>
                    </div>

                    {/* Column 2: Risk Indicator (8% width) */}
                    <div className="lg:w-[8%] flex items-center justify-start lg:justify-center flex-shrink-0">
                      {metrics && metrics.risk_level ? (
                        <span className={`inline-flex items-center gap-1 px-1.5 py-0.5 text-[9px] font-extrabold border rounded-md ${getRiskTone(metrics.risk_level).badge}`}>
                          <span className={`w-1 h-1 rounded-full ${getRiskTone(metrics.risk_level).dot}`} />
                          {metrics.risk_level}
                        </span>
                      ) : (
                        <span className="text-[9px] text-muted-foreground">无风险评估</span>
                      )}
                    </div>

                    {/* Column 3: High Density Financial Metrics (44% width) */}
                    <div className="lg:w-[44%] bg-muted/50 rounded-lg p-1.5 border border-border/60">
                      <div className="grid grid-cols-4 gap-1 text-center divide-x divide-border/40">
                        <div className="px-1 text-left sm:text-center">
                          <span className="text-[8px] text-muted-foreground font-bold block scale-90 origin-left sm:origin-center">毛利率</span>
                          <span className="text-caption font-black text-foreground truncate block">
                            {metrics ? formatMetricPercent(metrics.margin_rate) : "--"}
                          </span>
                        </div>

                        <div className="px-1 text-left sm:text-center">
                          <span className="text-[8px] text-muted-foreground font-bold block scale-90 origin-left sm:origin-center">NPV</span>
                          <span className="text-caption font-black text-foreground truncate block">
                            {metrics ? `¥${formatMetricNumber(metrics.npv)}` : "--"}
                          </span>
                        </div>

                        <div className="px-1 text-left sm:text-center">
                          <span className="text-[8px] text-muted-foreground font-bold block scale-90 origin-left sm:origin-center">NPVR</span>
                          <span className="text-caption font-black text-foreground truncate block">
                            {metrics ? formatMetricPercent(metrics.npv_rate) : "--"}
                          </span>
                        </div>

                        <div className="px-1 text-left sm:text-center">
                          <span className="text-[8px] text-muted-foreground font-bold block scale-90 origin-left sm:origin-center">IRR</span>
                          <span className="text-caption font-black text-foreground truncate block">
                            {metrics ? formatMetricPercent(metrics.irr) : "--"}
                          </span>
                        </div>
                      </div>
                    </div>

                    {/* Column 4: Truncated Remarks with tooltip (12% width) */}
                    <div className="lg:w-[12%] min-w-0 flex items-center gap-1 bg-warning-soft px-2 py-1 rounded-lg border border-warning/20">
                      <StickyNote className="h-3 w-3 text-warning flex-shrink-0" />
                      <span
                        className="text-[10px] text-foreground font-semibold truncate flex-1"
                        title={project.note ? `项目备注: ${project.note}` : "暂无备注"}
                      >
                        {project.note || "暂无备注"}
                      </span>
                    </div>

                    {/* Column 5: Single Row Actions (10% width) */}
                    <div className="lg:w-[10%] flex lg:flex-col items-center lg:items-end justify-between lg:justify-center gap-1 flex-shrink-0 text-right border-t lg:border-t-0 border-border pt-2 lg:pt-0">
                      <span className="text-[8px] text-muted-foreground scale-90 origin-right">
                        更新于 {new Date(project.updated_at).toLocaleDateString()}
                      </span>
                      <div className="flex items-center gap-1.5 mt-1 lg:mt-0">
                        {renderOpenFolderButton(project, true)}
                        {renderOpenIntelligentComputeButton(project, `open_intelligent_compact_${project.id}`, true)}
                        {renderOpenCalcButton(project, `open_calc_btn_compact_${project.id}`, true)}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
            <div className="flex-1 overflow-y-auto p-6">
              <div
                className="grid gap-6 animate-in fade-in duration-300"
                style={{ gridTemplateColumns: 'repeat(auto-fill, minmax(360px, 1fr))' }}
              >
                {renderCreateProjectEntry("grid")}
                {filteredProjects.length === 0 && (
                  <div className="rounded-2xl border border-dashed border-border bg-card/60 p-8 text-center text-sm font-semibold text-secondary-foreground">
                    {projectSearchTerm ? "没有匹配的项目" : "该阶段下暂无项目"}
                  </div>
                )}
                {filteredProjects.map((project) => {
                  const metrics = project.summary_metrics;

                  // ==================== MODE 1: ORIGINAL (低密度 - 原版大卡片) ====================
                  if (densityMode === "original") {
                    return (
                      <div
                        key={project.id}
                        id={`project_card_${project.id}`}
                        onClick={() => handleOpenDetails(project)}
                        className="group relative flex cursor-pointer flex-col overflow-hidden rounded-2xl border border-border bg-card shadow-sm transition-all duration-200 hover:border-border/80 hover:shadow-md animate-in fade-in slide-in-from-bottom-2"
                      >
                        {renderProjectCardHeader(project)}

                        <div className="bg-muted/30 px-5 py-4">
                          {renderMetricPanel(metrics)}
                        </div>

                        <div className="border-t border-border p-5">
                          {renderProjectNote(project)}
                        </div>

                        <div className="mt-auto flex items-center justify-between gap-4 border-t border-border bg-muted/40 px-5 py-3.5">
                          <span className="text-caption font-medium text-muted-foreground">
                            更新于 {new Date(project.updated_at).toLocaleDateString()}
                          </span>
                          <div className="flex items-center gap-2">
                            {renderOpenFolderButton(project, true)}
                            {renderOpenIntelligentComputeButton(project, `open_intelligent_grid_${project.id}`, true)}
                            {renderOpenCalcButton(project, `open_calc_btn_grid_${project.id}`)}
                          </div>
                        </div>
                      </div>
                    );
                  }

                  // ==================== MODE 2: STANDARD (中密度 - 标准配置) ====================
                  if (densityMode === "standard") {
                    return (
                      <div
                        key={project.id}
                        id={`project_card_${project.id}`}
                        onClick={() => handleOpenDetails(project)}
                        className="group relative flex cursor-pointer flex-col overflow-hidden rounded-xl border border-border bg-card shadow-sm transition-all duration-200 hover:border-border/80 hover:shadow-md animate-in fade-in slide-in-from-bottom-2"
                      >
                        {renderProjectCardHeader(project, true)}

                        <div className="bg-muted/20 px-4 py-3">
                          {renderMetricPanel(metrics, true)}
                        </div>

                        <div className="border-t border-border p-4">
                          {renderProjectNote(project, true)}
                        </div>

                        <div className="mt-auto flex items-center justify-between gap-4 border-t border-border bg-muted/30 px-4 py-2.5">
                          <span className="text-caption font-medium text-muted-foreground">
                            更新于 {new Date(project.updated_at).toLocaleDateString()}
                          </span>
                          <div className="flex items-center gap-2">
                            {renderOpenFolderButton(project, true)}
                            {renderOpenIntelligentComputeButton(project, `open_intelligent_grid_standard_${project.id}`, true)}
                            {renderOpenCalcButton(project, `open_calc_btn_grid_standard_${project.id}`, true)}
                          </div>
                        </div>
                      </div>
                    );
                  }

                  // ==================== MODE 3: ULTRA-COMPACT (高密度 - 极致紧凑卡片) ====================
                  return (
                    <div
                      key={project.id}
                      id={`project_card_${project.id}`}
                      onClick={() => handleOpenDetails(project)}
                      className={`group relative flex cursor-pointer flex-col justify-between overflow-hidden rounded-xl border border-border/90 bg-card shadow-sm hover:shadow-md transition-all duration-200 p-4 min-h-[220px] ${getRiskTopBorderStyles(metrics?.risk_level)} animate-in fade-in slide-in-from-bottom-2`}
                    >
                      {/* Header: Title + Client + Stage Tag Inline */}
                      <div className="border-b border-border/80 pb-2 flex items-start justify-between gap-2">
                        <div className="min-w-0 flex-1">
                          <div className="flex items-center gap-1.5 flex-wrap min-w-0">
                            <h3 className="text-caption font-black text-foreground truncate" title={project.name}>
                              {project.name}
                            </h3>
                            <span className="px-1 py-0.5 text-[8px] font-bold bg-muted text-muted-foreground rounded flex-shrink-0">
                              {project.status}
                            </span>
                            {renderProjectTypeBadge(project)}
                            {project.directoryExists === false && (
                              <span className="px-1 py-0.5 text-[8px] font-bold bg-destructive-soft border border-destructive-soft text-destructive rounded flex-shrink-0 flex items-center gap-0.5 animate-pulse" title="在磁盘中找不到项目对应的文件夹">
                                <AlertTriangle className="h-2.5 w-2.5 shrink-0" />
                                目录缺失
                              </span>
                            )}
                          </div>
                          <p className="text-caption text-muted-foreground truncate mt-0.5" title={project.customer_name || "未填写"}>
                            {project.customer_name || "未填写"}
                          </p>
                        </div>

                        {/* Risk Badge on Top Right */}
                        {metrics && metrics.risk_level ? (
                          <span className={`inline-flex items-center gap-1 px-1.5 py-0.5 text-[9px] font-extrabold border rounded-md flex-shrink-0 ${getRiskTone(metrics.risk_level).badge}`}>
                            <span className={`w-1 h-1 rounded-full ${getRiskTone(metrics.risk_level).dot}`} />
                            {metrics.risk_level}
                          </span>
                        ) : (
                          <span className="text-[9px] text-muted-foreground">无风险评估</span>
                        )}
                      </div>

                      {/* Compact Metrics Row */}
                      <div className="py-2.5">
                        <div className="grid grid-cols-2 gap-x-4 gap-y-1.5 text-xs font-mono">
                          <div className="flex justify-between items-baseline border-b border-border/50 pb-1">
                            <span className="text-[10px] text-muted-foreground font-semibold font-sans">毛利率</span>
                            <span className="text-caption font-black text-foreground numeric-value">
                              {metrics ? formatMetricPercent(metrics.margin_rate) : "--"}
                            </span>
                          </div>
                          <div className="flex justify-between items-baseline border-b border-border/50 pb-1">
                            <span className="text-[10px] text-muted-foreground font-semibold font-sans">NPV</span>
                            <span className="text-caption font-black text-foreground numeric-value">
                              {metrics ? `¥${formatMetricNumber(metrics.npv)}` : "--"}
                            </span>
                          </div>
                          <div className="flex justify-between items-baseline">
                            <span className="text-[10px] text-muted-foreground font-semibold font-sans">NPVR</span>
                            <span className="text-caption font-black text-foreground numeric-value">
                              {metrics ? formatMetricPercent(metrics.npv_rate) : "--"}
                            </span>
                          </div>
                          <div className="flex justify-between items-baseline">
                            <span className="text-[10px] text-muted-foreground font-semibold font-sans">IRR</span>
                            <span className="text-caption font-black text-foreground numeric-value">
                              {metrics ? formatMetricPercent(metrics.irr) : "--"}
                            </span>
                          </div>
                        </div>
                      </div>

                      {/* Compressed Remarks Box */}
                      <div className="bg-muted/50 p-2 rounded-xl text-[10px] flex items-center gap-1.5 border border-border/50 mb-2">
                        <StickyNote className="h-3 w-3 text-warning flex-shrink-0" />
                        <span
                          className="text-muted-foreground font-semibold truncate flex-1 cursor-help"
                          title={project.note ? `项目备注: ${project.note}` : "暂无备注"}
                        >
                          {project.note || "暂无备注"}
                        </span>
                      </div>

                      {/* Footer Actions */}
                      <div className="border-t border-border/80 pt-2 flex items-center justify-between text-[10px]">
                        <span className="text-muted-foreground">
                          更新于 {new Date(project.updated_at).toLocaleDateString()}
                        </span>
                        <div className="flex items-center gap-1.5">
                          {renderOpenFolderButton(project, true)}
                          {renderOpenIntelligentComputeButton(project, `open_intelligent_grid_compact_${project.id}`, true)}
                          {renderOpenCalcButton(project, `open_calc_btn_grid_compact_${project.id}`, true)}
                        </div>
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
            className="bg-card border border-border rounded-xl shadow-xl w-full max-w-md overflow-hidden"
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
                  className="bg-card border border-input px-3 py-2 rounded-lg text-sm outline-none focus:border-ring w-full"
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
                  className="bg-card border border-input px-3 py-2 rounded-lg text-sm outline-none focus:border-ring w-full"
                />
              </div>
              <div className="flex flex-col gap-1.5">
                <label className="text-sm font-semibold text-secondary-foreground">项目类型</label>
                <select
                  value={newProjectType}
                  onChange={event => setNewProjectType(event.target.value as ProjectType)}
                  className="bg-card border border-input px-3 py-2 rounded-lg text-sm outline-none focus:border-ring w-full"
                >
                  <option value="ict">ICT 项目</option>
                  <option value="intelligent_compute">智算项目</option>
                </select>
                <p className="text-xs text-muted-foreground">
                  智算项目可维护独立金额来源，并在确认后同步至 ICT 测算。
                </p>
              </div>

              <div className="flex flex-col gap-1.5">
                <label className="text-sm font-semibold text-secondary-foreground">项目预设</label>
                <select
                  value={newProjectPresetId}
                  onChange={event => setNewProjectPresetId(event.target.value)}
                  className="bg-card border border-input px-3 py-2 rounded-lg text-sm outline-none focus:border-ring w-full"
                >
                  <option value="">空白项目</option>
                  {projectPresetTemplates.map(template => (
                    <option key={template.id} value={template.id}>
                      {template.name}
                    </option>
                  ))}
                </select>
                <span className="text-xs text-muted-foreground">
                  创建失败时会回滚项目记录和目录，不会留下半初始化项目。
                </span>
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



      {showStatusManager && (
        <div className="fixed inset-0 z-[60] bg-background/80 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in">
          <div className="bg-card border border-border rounded-xl shadow-xl w-full max-w-md overflow-hidden">
            <div className="px-6 py-4 border-b border-border bg-muted/30 flex items-center justify-between">
              <h2 className="font-bold text-base text-foreground flex items-center gap-2">
                <Settings2 className="h-4 w-4 text-primary" />
                管理项目阶段
              </h2>
              <button
                type="button"
                onClick={() => setShowStatusManager(false)}
                className="text-secondary-foreground hover:bg-secondary p-1 rounded-md"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <div className="p-6 space-y-4">
              <div className="space-y-2">
                {statusDrafts.map((status, index) => (
                  <div key={`${status}-${index}`} className="flex items-center gap-2">
                    <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg bg-muted text-xs font-extrabold text-secondary-foreground">
                      {index + 1}
                    </span>
                    <input
                      value={status}
                      onChange={(event) => {
                        const next = [...statusDrafts];
                        next[index] = event.target.value;
                        setStatusDrafts(next);
                      }}
                      className="min-w-0 flex-1 rounded-lg border border-input bg-card px-3 py-2 text-sm font-semibold text-foreground outline-none transition-all focus:border-ring focus:ring-2 focus:ring-ring/20"
                    />
                  </div>
                ))}
              </div>

              <div className="flex gap-2 border-t border-border/60 pt-4">
                <input
                  value={newStatusDraft}
                  onChange={(event) => setNewStatusDraft(event.target.value)}
                  onKeyDown={(event) => {
                    if (event.key === "Enter") {
                      event.preventDefault();
                      handleAddStatusDraft();
                    }
                  }}
                  placeholder="新增阶段名称"
                  className="min-w-0 flex-1 rounded-lg border border-input bg-card px-3 py-2 text-sm text-foreground outline-none transition-all focus:border-ring focus:ring-2 focus:ring-ring/20"
                />
                <button
                  type="button"
                  onClick={handleAddStatusDraft}
                  className="inline-flex shrink-0 items-center gap-1.5 rounded-lg bg-primary/10 px-3 py-2 text-xs font-bold text-primary transition-all hover:bg-primary/15"
                >
                  <Plus className="h-3.5 w-3.5" />
                  添加
                </button>
              </div>
            </div>

            <div className="border-t border-border p-4 bg-muted/20 flex justify-end gap-3">
              <button
                type="button"
                onClick={() => setShowStatusManager(false)}
                className="px-4 py-2 border border-border hover:bg-secondary rounded-lg text-sm font-semibold text-secondary-foreground transition-all"
              >
                取消
              </button>
              <button
                type="button"
                onClick={handleSaveStatusManager}
                className="px-4 py-2 bg-primary hover:bg-primary/90 text-primary-foreground font-bold rounded-lg text-sm transition-all"
              >
                保存阶段
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
            className="bg-card border border-border shadow-xl w-[88vw] max-w-6xl h-[86vh] max-h-[920px] min-h-[560px] rounded-2xl overflow-hidden flex flex-col animate-in zoom-in-95 fade-in duration-200"
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
                        onChange={(e) => {
                          setEditingProjectName(e.target.value);
                          markDirty("project-detail");
                        }}
                        className="bg-card border border-input px-3 py-2 rounded-lg text-sm outline-none focus:border-ring"
                      />
                    </div>
                    <div className="flex flex-col gap-1">
                      <label className="text-xs font-semibold text-secondary-foreground">客户名称</label>
                      <input
                        type="text"
                        value={editingCustomerName}
                        onChange={(e) => {
                          setEditingCustomerName(e.target.value);
                          markDirty("project-detail");
                        }}
                        className="bg-card border border-input px-3 py-2 rounded-lg text-sm outline-none focus:border-ring"
                      />
                    </div>
                    <div className="flex flex-col gap-1 col-span-2">
                      <label className="text-xs font-semibold text-secondary-foreground">项目类型</label>
                      <select
                        value={editingProjectType}
                        onChange={event => {
                          const nextType = event.target.value as ProjectType;
                          if (
                            selectedProject.project_type === "intelligent_compute"
                            && nextType === "ict"
                            && !window.confirm("转为 ICT 项目后将隐藏智算入口，但保留智算金额来源和已同步的 ICT 数据。确定继续吗？")
                          ) {
                            return;
                          }
                          setEditingProjectType(nextType);
                          markDirty("project-detail");
                        }}
                        className="bg-card border border-input px-3 py-2 rounded-lg text-sm outline-none focus:border-ring"
                      >
                        <option value="ict">ICT 项目</option>
                        <option value="intelligent_compute">智算项目</option>
                      </select>
                    </div>
                    <div className="flex flex-col gap-1 col-span-2">
                      <div className="flex items-center justify-between gap-3">
                        <label className="text-xs font-semibold text-secondary-foreground">看板阶段</label>
                        <button
                          type="button"
                          onClick={handleOpenStatusManager}
                          className="text-[11px] font-bold text-primary hover:underline"
                        >
                          管理阶段
                        </button>
                      </div>
                      <select
                        value={editingStatus}
                        onChange={(e) => {
                          setEditingStatus(e.target.value);
                          markDirty("project-detail");
                        }}
                        className="bg-card border border-input px-3 py-2 rounded-lg text-sm outline-none focus:border-ring"
                      >
                        {!statusOptions.includes(editingStatus) && editingStatus && (
                          <option value={editingStatus}>{editingStatus}</option>
                        )}
                        {statusOptions.map(col => (
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
                              {(() => {
                                const stageOption = getSchemeStageOption(scheme.stage);
                                return stageOption ? (
                                  <span className={`text-[8px] px-1 py-0.5 rounded shrink-0 ${stageOption.chipClass}`}>
                                    {stageOption.short}
                                  </span>
                                ) : null;
                              })()}
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
                          onClick={async () => {
                            const canProceed = await confirmOrSave();
                            if (!canProceed) return;
                            onOpenCalc(selectedProject.id, selectedScheme.id);
                          }}
                          className="text-xs text-primary font-bold flex items-center gap-1 hover:underline"
                        >
                          <AppIcon name="calculator" size={12} /> 开展测算
                        </button>
                      </div>

                      {/* 甄选阶段标签 */}
                      <div className="flex items-center gap-2">
                        <span className="text-xs text-secondary-foreground shrink-0">甄选阶段</span>
                        <div className="flex gap-1.5">
                          {SCHEME_STAGE_OPTIONS.map((option) => {
                            const active = selectedScheme.stage === option.value;
                            return (
                              <button
                                key={option.value}
                                onClick={() => handleSchemeStageChange(selectedScheme, option.value)}
                                className={`text-xs px-2.5 py-1 rounded-md font-semibold transition-all ${
                                  active
                                    ? option.chipClass
                                    : "bg-muted/40 text-secondary-foreground hover:bg-secondary"
                                }`}
                                title={active ? "点击取消标注" : `标注为${option.label}`}
                              >
                                {option.label}
                              </button>
                            );
                          })}
                        </div>
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
                                  <div>毛利率: <span className="font-bold text-foreground numeric-value">{formatMetricPercent(snap.output_metrics.margin_rate)}</span></div>
                                  <div>NPV: <span className="font-bold text-foreground numeric-value">{formatMetricNumber(snap.output_metrics.npv)}</span></div>
                                  <div>IRR: <span className="font-bold text-foreground numeric-value">{formatMetricPercent(snap.output_metrics.irr)}</span></div>
                                </div>
                              </div>
                              <button
                                onClick={async () => {
                                  const canProceed = await confirmOrSave();
                                  if (!canProceed) return;
                                  onOpenCalc(selectedProject.id, selectedScheme.id);
                                }}
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
          <div className="bg-card border border-border rounded-xl shadow-xl w-full max-w-sm overflow-hidden">
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
                className="bg-card border border-input px-3 py-2 rounded-lg text-xs outline-none focus:border-ring w-full"
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

      {/* Import Candidates Scanner Modal */}
      {showImportModal && (
        <div className="fixed inset-0 z-[60] bg-background/80 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in">
          <div className="bg-card border border-border rounded-xl shadow-xl w-full max-w-4xl max-h-[85vh] overflow-hidden flex flex-col">
            <div className="px-6 py-4 border-b border-border bg-muted/30 flex items-center justify-between shrink-0 font-bold">
              <div className="flex items-center gap-2">
                <FolderPlus className="h-5 w-5 text-primary" />
                <h4 className="font-extrabold text-sm text-foreground">批量扫描与项目导入</h4>
              </div>
              <button
                type="button"
                onClick={() => setShowImportModal(false)}
                className="text-secondary-foreground hover:bg-secondary p-1 rounded-md transition-colors"
                disabled={importLoading}
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <div className="p-6 overflow-y-auto flex-1 space-y-4">
              <div className="bg-muted/30 p-4 rounded-xl flex items-center justify-between gap-4 border border-border/40">
                <div className="min-w-0">
                  <span className="text-[10px] uppercase font-extrabold text-secondary-foreground opacity-70 block">当前扫描目录</span>
                  <code className="text-xs font-mono text-primary truncate block mt-0.5 select-all">{importParentPath}</code>
                </div>
                <button
                  type="button"
                  onClick={handleOpenImportScanner}
                  disabled={importLoading || scanLoading}
                  className="px-3 py-1.5 bg-secondary hover:bg-muted text-secondary-foreground rounded-lg text-xs font-bold transition-all border border-input shrink-0"
                >
                  重新选择目录
                </button>
              </div>

              {scanLoading ? (
                <div className="py-20 flex flex-col items-center justify-center gap-3 text-secondary-foreground">
                  <RefreshCw className="animate-spin text-primary h-8 w-8" />
                  <span className="text-sm font-semibold">正在扫描子文件夹，并智能分析项目文件...</span>
                </div>
              ) : importCandidates.length === 0 ? (
                <div className="py-20 text-center text-secondary-foreground">
                  <span className="text-sm">在该目录下未扫描到任何有效的候选项目子文件夹。</span>
                </div>
              ) : (
                <div className="space-y-3">
                  <div className="flex justify-between items-center text-xs font-bold text-secondary-foreground px-1 pb-1">
                    <span>候选目录数量: {importCandidates.length}</span>
                    <div className="flex gap-3">
                      <button
                        type="button"
                        onClick={() => {
                          const next: Record<string, boolean> = {};
                          importCandidates.forEach(c => next[c.folderPath] = true);
                          setSelectedCandidates(next);
                        }}
                        className="text-primary hover:underline font-bold"
                      >
                        全选
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          const next: Record<string, boolean> = {};
                          importCandidates.forEach(c => next[c.folderPath] = false);
                          setSelectedCandidates(next);
                        }}
                        className="text-primary hover:underline font-bold"
                      >
                        取消全选
                      </button>
                    </div>
                  </div>

                  <div className="space-y-3">
                    {importCandidates.map(c => {
                      const isSelected = !!selectedCandidates[c.folderPath];
                      const isExpanded = !!expandedCandidates[c.folderPath];
                      const conflictAction = conflictActions[c.folderPath] || "merge";

                      return (
                        <div
                          key={c.folderPath}
                          className={`rounded-xl border transition-all ${
                            isSelected
                              ? "bg-card border-border/80"
                              : "bg-muted/10 border-border/40 opacity-70"
                          }`}
                        >
                          <div className="p-4 flex flex-col md:flex-row md:items-center justify-between gap-4 select-none">
                            <div className="flex items-start gap-3 min-w-0">
                              <input
                                type="checkbox"
                                checked={isSelected}
                                onChange={e => {
                                  setSelectedCandidates(prev => ({
                                    ...prev,
                                    [c.folderPath]: e.target.checked
                                  }));
                                }}
                                className="mt-1 h-4.5 w-4.5 rounded border-input text-primary focus:ring-primary/20 accent-primary cursor-pointer shrink-0"
                              />
                              <div className="min-w-0">
                                <span className="font-extrabold text-sm text-foreground flex items-center gap-2">
                                  <FolderOpen className="h-4 w-4 text-muted-foreground shrink-0" />
                                  {c.folderName}
                                  {c.existsConflict && (
                                    <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded-md bg-warning-soft text-warning-foreground text-[9px] font-bold border border-warning/20">
                                      <AlertTriangle className="h-2.5 w-2.5" />
                                      同名项目已存在
                                    </span>
                                  )}
                                </span>
                                <code className="text-[10px] text-muted-foreground font-mono block mt-1 truncate">
                                  {c.folderPath}
                                </code>
                              </div>
                            </div>

                            <div className="flex items-center gap-3 self-end md:self-auto">
                              {c.existsConflict && isSelected && (
                                <div className="flex items-center gap-1.5 shrink-0 bg-warning-soft px-2 py-1 rounded-lg border border-warning/20">
                                  <span className="text-[10px] font-bold text-warning-foreground uppercase">冲突策略:</span>
                                  <select
                                    value={conflictAction}
                                    onChange={e => {
                                      const val = e.target.value as "merge" | "new" | "skip";
                                      setConflictActions(prev => ({
                                        ...prev,
                                        [c.folderPath]: val
                                      }));
                                    }}
                                    className="bg-transparent text-[10px] font-bold text-warning-foreground outline-none cursor-pointer"
                                  >
                                    <option value="merge">覆盖合并 (Merge)</option>
                                    <option value="new">另存为新项目 (New)</option>
                                    <option value="skip">跳过此项 (Skip)</option>
                                  </select>
                                </div>
                              )}

                              <button
                                type="button"
                                onClick={() => {
                                  setExpandedCandidates(prev => ({
                                    ...prev,
                                    [c.folderPath]: !isExpanded
                                  }));
                                }}
                                className="text-[10px] font-bold text-primary hover:underline inline-flex items-center gap-1 shrink-0 px-2 py-1 hover:bg-muted rounded"
                              >
                                {isExpanded ? <ChevronUp className="h-3 w-3" /> : <ChevronDown className="h-3 w-3" />}
                                {isExpanded ? "隐藏文件" : `查看文件 (${c.files.length})`}
                              </button>
                            </div>
                          </div>

                          {isExpanded && c.files.length > 0 && (
                            <div className="px-4 pb-4 pt-1 border-t border-border/20 bg-muted/10 rounded-b-xl">
                              <div className="space-y-1.5 pt-2">
                                {c.files.map(f => {
                                  let pillClass = "bg-muted text-muted-foreground";
                                  let roleLabel = "普通文件";
                                  if (f.fileRole === "benefit_scheme") {
                                    pillClass = "bg-primary-soft text-primary";
                                    roleLabel = "效益测算主方案";
                                  } else if (f.fileRole === "budget") {
                                    pillClass = "bg-warning-soft text-warning-foreground";
                                    roleLabel = "项目预算表";
                                  } else if (f.fileRole === "proposal") {
                                    pillClass = "bg-success-soft text-success-foreground";
                                    roleLabel = "立项申报书";
                                  }

                                  return (
                                    <div key={f.path} className="flex items-center justify-between text-xs py-1 px-2 hover:bg-card rounded transition-colors font-mono">
                                      <span className="truncate text-secondary-foreground pr-4" title={f.name}>{f.name}</span>
                                      <span className={`px-2 py-0.5 rounded text-[8px] font-bold shrink-0 uppercase ${pillClass}`}>
                                        {roleLabel}
                                      </span>
                                    </div>
                                  );
                                })}
                              </div>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>

            <div className="border-t border-border p-4 bg-muted/20 flex justify-end gap-2 shrink-0">
              <button
                type="button"
                onClick={() => setShowImportModal(false)}
                disabled={importLoading}
                className="px-4 py-2 border border-border bg-card hover:bg-secondary rounded-lg text-xs font-semibold text-secondary-foreground transition-all disabled:opacity-50"
              >
                取消
              </button>
              <button
                type="button"
                onClick={handleExecuteImport}
                disabled={importLoading || scanLoading || importCandidates.length === 0}
                className="px-4 py-2 bg-primary hover:bg-primary/95 disabled:opacity-50 text-white font-extrabold rounded-lg text-xs transition-all flex items-center gap-1.5 shadow-sm active:scale-[0.98]"
              >
                {importLoading && <RefreshCw className="animate-spin h-3.5 w-3.5 shrink-0" />}
                开始批量导入
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
