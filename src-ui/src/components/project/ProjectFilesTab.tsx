import { useEffect, useState } from "react";
import {
  FileText,
  FileSpreadsheet,
  Folder,
  FolderPlus,
  FolderOpen,
  RefreshCw,
  Trash2,
  ExternalLink,
  Plus,
  Check,
  AlertTriangle,
  Pin,
  File,
  Search,
  Eye,
  FileWarning
} from "lucide-react";
import { projectFileService, type ProjectFile } from "../../services/projectFileService";
import { projectService, type Project } from "../../utils/projectService";
import { invoke } from "@tauri-apps/api/core";

interface ProjectFilesTabProps {
  projectId: string;
  onRefreshProject?: () => void;
}

export default function ProjectFilesTab({ projectId, onRefreshProject }: ProjectFilesTabProps) {
  const [project, setProject] = useState<Project | null>(null);
  const [files, setFiles] = useState<ProjectFile[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState("");
  const [pendingBindFolder, setPendingBindFolder] = useState<{ folderPath: string; folderName: string } | null>(null);
  const [createFolderParentPath, setCreateFolderParentPath] = useState<string | null>(null);
  const [createFolderName, setCreateFolderName] = useState("");
  const [isCreateFolderModalOpen, setIsCreateFolderModalOpen] = useState(false);
  const [notInRootFolder, setNotInRootFolder] = useState<{ folderPath: string; renameProject: boolean } | null>(null);

  const loadData = async () => {
    setLoading(true);
    setError(null);
    try {
      const proj = await projectService.getProject(projectId);
      setProject(proj);
      const fileList = await projectFileService.getProjectFiles(projectId);
      // Sort files: main documents first, then recently updated
      const sorted = [...fileList].sort((a, b) => {
        if (a.isMainDocument && !b.isMainDocument) return -1;
        if (!a.isMainDocument && b.isMainDocument) return 1;
        if (a.isMainBudgetFile && !b.isMainBudgetFile) return -1;
        if (!a.isMainBudgetFile && b.isMainBudgetFile) return 1;
        return new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime();
      });
      setFiles(sorted);
    } catch (err: any) {
      setError(err?.toString() || "加载项目文件失败");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, [projectId]);

  const showTemporaryMessage = (msg: string, isError = false) => {
    if (isError) {
      setError(msg);
      setTimeout(() => setError(null), 5000);
    } else {
      setSuccessMsg(msg);
      setTimeout(() => setSuccessMsg(null), 3000);
    }
  };

  const getFolderName = (path: string) => {
    return path.split(/[/\\]/).filter(Boolean).pop() || "";
  };

  const importExcelData = async (filePath: string) => {
    setLoading(true);
    try {
      const parsedData = await projectFileService.parseBenefitExcel(filePath);
      const latestProject = await projectService.getProject(projectId);
      const latestSchemes = await projectService.getSchemes(projectId);
      const importScheme = latestSchemes.find(s => s.name === "Excel导入测算方案");

      const ctIncome = parsedData.ct_income_incl || 0;
      const itIncome = Math.max(0, parsedData.total_income_incl - ctIncome);
      const toTaxPercent = (value: number | undefined, fallback: number) => {
        const numeric = Number(value);
        if (!Number.isFinite(numeric) || numeric <= 0) return fallback;
        return numeric > 0 && numeric < 1 ? numeric * 100 : numeric;
      };
      const itTax = toTaxPercent(parsedData.it_tax, 6);
      const ctTax = toTaxPercent(parsedData.ct_tax, 6);
      const parsedItems = parsedData.items || {};
      const hasDetailedItems = Object.values(parsedItems).some(item => {
        return Math.abs(Number(item?.incl_tax || 0)) > 0 || Math.abs(Number(item?.excl_tax || 0)) > 0;
      });

      const makeItem = (incl: number, tax: number, customSubjectName?: string | null, billingSubjectName?: string | null) => {
        const custom = String(customSubjectName || "").trim();
        const billing = String(billingSubjectName || "").trim();
        return {
          incl_tax: String(Number.isFinite(incl) ? Number(incl.toFixed(2)) : 0),
          tax_rate: String(Number.isFinite(tax) ? Number(tax.toFixed(4)) : 0),
          ...(custom ? { custom_subject_name: custom } : {}),
          ...(billing ? { billing_subject_name: billing } : {}),
        };
      };

      const makeParsedItem = (key: string, defaultTax: number, fallbackIncl = 0) => {
        const item = parsedItems[key];
        if (item) {
          const tax = toTaxPercent(item.tax_rate, defaultTax);
          const parsedIncl = Number(item.incl_tax);
          if (Number.isFinite(parsedIncl) && Math.abs(parsedIncl) > 0) {
            return makeItem(parsedIncl, tax, item.custom_subject_name, item.billing_subject_name);
          }

          const parsedExcl = Number(item.excl_tax);
          if (Number.isFinite(parsedExcl) && Math.abs(parsedExcl) > 0) {
            return makeItem(parsedExcl * (1 + tax / 100), tax, item.custom_subject_name, item.billing_subject_name);
          }

          return makeItem(0, tax, item.custom_subject_name, item.billing_subject_name);
        }

        return makeItem(hasDetailedItems ? 0 : fallbackIncl, defaultTax);
      };

      const payload = {
        project_name: parsedData.project_name || latestProject?.name || project?.name || "未命名项目",
        customer_name: parsedData.customer_name || latestProject?.customer_name || project?.customer_name || "未指定客户",
        property_rights: "客户",
        discount_rate: String(parsedData.discount_rate || 0.055),
        project_years: parsedData.project_years > 0 ? parsedData.project_years : 1,
        cashflow_model: "model_a",
        cashflow_calculation_source: "subject_funding_plans",
        subject_funding_plan_migration_version: 1,
        cashflow_segment_value_mode: "ratio",
        cashflow_segments: [],
        ignore_tail_difference: false,
        tail_difference_value: "0",
        rev_distribution: [1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
        cost_distribution: [1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
        rev_cashflow_excl: null,
        cost_cashflow_excl: null,
        it_rev_cashflow_excl: null,
        it_cost_cashflow_excl: null,

        rev_it_integration: makeParsedItem("rev_it_integration", itTax, itIncome),
        rev_it_maintenance: makeParsedItem("rev_it_maintenance", 6),
        rev_it_device_sales: makeParsedItem("rev_it_device_sales", 13),
        rev_it_device_lease: makeParsedItem("rev_it_device_lease", 13),
        rev_it_other: makeParsedItem("rev_it_other", 6),
        rev_it_cloud: makeParsedItem("rev_it_cloud", 6),

        rev_ct_line: makeParsedItem("rev_ct_line", 9),
        rev_ct_product: makeParsedItem("rev_ct_product", ctTax, ctIncome),

        rev_non_it_ct: makeParsedItem("rev_non_it_ct", 9),

        cost_it_device: makeParsedItem("cost_it_device", 13, parsedData.total_cost_incl),
        cost_it_construction: makeParsedItem("cost_it_construction", 9),
        cost_it_survey: makeParsedItem("cost_it_survey", 6),
        cost_it_integration: makeParsedItem("cost_it_integration", 6),
        cost_it_other: makeParsedItem("cost_it_other", 6),
        cost_it_maintenance: makeParsedItem("cost_it_maintenance", 6),
        cost_it_running: makeParsedItem("cost_it_running", 13),
        cost_it_bidding: makeParsedItem("cost_it_bidding", 6),
        cost_it_design_eval: makeParsedItem("cost_it_design_eval", 6),
        cost_it_audit: makeParsedItem("cost_it_audit", 6),

        cost_ct_construction: makeParsedItem("cost_ct_construction", 6),
        cost_ct_maintenance: makeParsedItem("cost_ct_maintenance", 9),
        cost_ct_other: makeParsedItem("cost_ct_other", ctTax),
        cost_ct_bandwidth: makeParsedItem("cost_ct_bandwidth", 9),
        cost_ct_renewal: makeParsedItem("cost_ct_renewal", 9),

        cost_non_it_ct: makeParsedItem("cost_non_it_ct", 9),
        cost_mix_marketing: makeParsedItem("cost_mix_marketing", 6),
        cost_mix_channel: makeParsedItem("cost_mix_channel", 6),
        cost_mix_other: makeParsedItem("cost_mix_other", 6),
      };

      const res: any = await invoke('calculate_ict_benefit', { input: payload });

      const updatedProject = await projectService.saveBenefitScheme(
        projectId,
        latestProject?.default_scheme_id || importScheme?.id || null,
        "Excel导入测算方案",
        payload,
        res,
        false
      );

      setProject(updatedProject);
      showTemporaryMessage("成功导入测算数据！");
      await loadData();
      if (onRefreshProject) onRefreshProject();
    } catch (err: any) {
      showTemporaryMessage(err?.toString() || "导入数据失败", true);
    } finally {
      setLoading(false);
    }
  };

  const executeBindFolder = async (folderPath: string, renameProject: boolean, forceMode?: string) => {
    setLoading(true);
    try {
      const folderName = getFolderName(folderPath);
      let finalProjectName = project?.name || "";
      let renamed = false;
      let renameWarning: string | null = null;

      await projectFileService.bindProjectFolder(projectId, folderPath, forceMode);
      const latestProject = await projectService.getProject(projectId);
      if (latestProject) {
        setProject(latestProject);
        finalProjectName = latestProject.name;
      }

      const projectForRename = latestProject || project;
      if (renameProject && projectForRename && folderName && folderName !== projectForRename.name) {
        try {
          const updatedProject = await projectService.updateProject({
            ...projectForRename,
            name: folderName
          });
          setProject(updatedProject);
          finalProjectName = updatedProject.name;
          renamed = true;
        } catch (renameErr: any) {
          console.error(renameErr);
          renameWarning = renameErr?.toString() || "同步项目名称失败";
        }
      }

      const bindResultMessage = renameWarning
        ? `目录已绑定，但项目名称同步失败：${renameWarning}`
        : renamed
          ? `成功绑定本地项目目录，项目名称已更新为「${finalProjectName}」`
          : "成功绑定本地项目目录！";

      const scannedFiles = await projectFileService.scanProjectFolder(projectId, false);
      const excelFiles = scannedFiles.filter(f => f.fileType === 'excel' && f.exists);
      if (excelFiles.length > 0) {
        const expectedFileName = `效益分析表-${finalProjectName}.xlsx`;
        const expectedFileNameOld = `效益分析表-${finalProjectName}.xls`;
        const matchedFile = excelFiles.find(f => f.fileName === expectedFileName || f.fileName === expectedFileNameOld);

        if (matchedFile) {
          const confirmImport = window.confirm(`检测到当前绑定目录下存在效益分析文件 "${matchedFile.fileName}"，是否直接导入测算数据？`);
          if (confirmImport) {
            await importExcelData(matchedFile.filePath);
          }
        } else if (excelFiles.length === 1) {
          const confirmImport = window.confirm(`检测到目录下存在Excel文件 "${excelFiles[0].fileName}"，是否将其作为效益分析数据导入？`);
          if (confirmImport) {
            await importExcelData(excelFiles[0].filePath);
          }
        }
      }

      await loadData();
      showTemporaryMessage(bindResultMessage, Boolean(renameWarning));
      if (onRefreshProject) onRefreshProject();
    } catch (err: any) {
      const errMsg = err?.toString() || "";
      if (errMsg.includes("NOT_IN_ROOT")) {
        setNotInRootFolder({ folderPath, renameProject });
      } else {
        showTemporaryMessage(errMsg || "绑定目录失败", true);
      }
    } finally {
      setPendingBindFolder(null);
      setLoading(false);
    }
  };

  const handleBindFolder = async () => {
    try {
      const selected = await projectFileService.selectLocalFolder();
      if (!selected) return;

      const folderName = getFolderName(selected);
      if (project && folderName && folderName !== project.name) {
        setPendingBindFolder({ folderPath: selected, folderName });
        return;
      }

      await executeBindFolder(selected, false);
    } catch (err: any) {
      showTemporaryMessage(err?.toString() || "选择目录失败", true);
    }
  };

  const handleCreateFolderClick = async () => {
    try {
      const selected = await projectFileService.selectLocalFolder();
      if (!selected) return;

      setCreateFolderParentPath(selected);
      setCreateFolderName((project?.name || "新建项目").trim());
      setIsCreateFolderModalOpen(true);
    } catch (err: any) {
      showTemporaryMessage(err?.toString() || "选择父级目录失败", true);
    }
  };

  const handleConfirmCreateFolder = async () => {
    if (!createFolderParentPath) return;

    const folderName = createFolderName.trim();
    if (!folderName) {
      showTemporaryMessage("文件夹名称不能为空", true);
      return;
    }

    setLoading(true);
    try {
      const createdPath = await projectFileService.createProjectFolder(createFolderParentPath, folderName);
      setIsCreateFolderModalOpen(false);
      setCreateFolderParentPath(null);
      setCreateFolderName("");
      setLoading(false);
      await executeBindFolder(createdPath, false);
    } catch (err: any) {
      showTemporaryMessage(err?.toString() || "创建项目文件夹失败", true);
    } finally {
      setLoading(false);
    }
  };

  const handleUnbindFolder = async () => {
    if (!window.confirm("确定要解绑当前本地目录吗？这将会清除目录绑定和所有从该目录中扫描出来的文件关联，但不会删除你磁盘上的任何实际文件。")) {
      return;
    }
    setLoading(true);
    try {
      await projectFileService.unbindProjectFolder(projectId);
      showTemporaryMessage("已成功解绑本地目录！");
      await loadData();
      if (onRefreshProject) onRefreshProject();
    } catch (err: any) {
      showTemporaryMessage(err?.toString() || "解绑目录失败", true);
    } finally {
      setLoading(false);
    }
  };

  const handleScanFolder = async () => {
    if (!project?.folder_path) return;
    setLoading(true);
    try {
      await projectFileService.scanProjectFolder(projectId, false);
      showTemporaryMessage("项目目录扫描完成");
      await loadData();
    } catch (err: any) {
      showTemporaryMessage(err?.toString() || "扫描目录失败", true);
    } finally {
      setLoading(false);
    }
  };

  const handleAddFile = async (mode: 'linked' | 'copied') => {
    try {
      const selected = await projectFileService.selectLocalFile(
        mode === 'copied' ? "选择文件导入到托管沙盒" : "选择本地文件关联到项目",
        ["docx", "doc", "xlsx", "xls", "pdf", "pptx", "ppt", "png", "jpg", "jpeg", "gif"]
      );
      if (!selected) return;

      setLoading(true);
      await projectFileService.addProjectFile(projectId, selected, mode);
      showTemporaryMessage(mode === 'copied' ? "文件已导入到托管沙盒！" : "本地文件已关联！");
      await loadData();
      if (onRefreshProject) onRefreshProject();
    } catch (err: any) {
      showTemporaryMessage(err?.toString() || "添加文件失败", true);
    } finally {
      setLoading(false);
    }
  };

  const handleDeleteFile = async (file: ProjectFile) => {
    const confirmMsg = file.storageMode === 'copied'
      ? `确定要删除托管文件 "${file.fileName}" 吗？这将会从应用沙盒中物理删除该文件，且不可恢复。`
      : `确定要取消关联文件 "${file.fileName}" 吗？这仅会清除项目索引，不会删除你本地磁盘的实际文件。`;

    if (!window.confirm(confirmMsg)) return;

    setLoading(true);
    try {
      if (file.storageMode === 'copied') {
        await projectFileService.deleteManagedProjectFile(projectId, file.id);
      } else {
        await projectFileService.removeProjectFileRecord(projectId, file.id);
      }
      showTemporaryMessage("文件处理成功");
      await loadData();
      if (onRefreshProject) onRefreshProject();
    } catch (err: any) {
      showTemporaryMessage(err?.toString() || "删除文件失败", true);
    } finally {
      setLoading(false);
    }
  };

  const handleToggleMainDoc = async (file: ProjectFile) => {
    setLoading(true);
    try {
      const newValue = file.isMainDocument ? null : file.id;
      await projectFileService.markMainDocument(projectId, newValue);
      showTemporaryMessage(newValue ? "主效益测算文档设置成功" : "已清除主效益测算文档");
      await loadData();
      if (onRefreshProject) onRefreshProject();
    } catch (err: any) {
      showTemporaryMessage(err?.toString() || "设置主测算文档失败", true);
    } finally {
      setLoading(false);
    }
  };

  const handleToggleMainBudget = async (file: ProjectFile) => {
    setLoading(true);
    try {
      const newValue = file.isMainBudgetFile ? null : file.id;
      await projectFileService.markMainBudgetFile(projectId, newValue);
      showTemporaryMessage(newValue ? "主预算文件设置成功" : "已清除主预算文件");
      await loadData();
      if (onRefreshProject) onRefreshProject();
    } catch (err: any) {
      showTemporaryMessage(err?.toString() || "设置主预算文件失败", true);
    } finally {
      setLoading(false);
    }
  };

  const handleOpenFile = async (file: ProjectFile) => {
    try {
      await projectFileService.openProjectFile(file.id);
    } catch (err: any) {
      showTemporaryMessage(err?.toString() || "打开文件失败，可能文件已被移动或删除", true);
    }
  };

  const handleRevealFile = async (file: ProjectFile) => {
    try {
      await projectFileService.revealProjectFile(file.id);
    } catch (err: any) {
      showTemporaryMessage(err?.toString() || "定位文件失败，可能文件已被移动或删除", true);
    }
  };

  const handleOpenFolder = async () => {
    try {
      await projectFileService.openProjectFolder(projectId);
    } catch (err: any) {
      showTemporaryMessage(err?.toString() || "打开文件夹失败，可能文件夹已被移动或删除", true);
    }
  };

  const formatSize = (bytes: number) => {
    if (bytes === 0) return "0 Bytes";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
  };

  const getFileIcon = (type: string) => {
    switch (type) {
      case "word":
        return <FileText className="w-5 h-5 text-blue-600" />;
      case "excel":
        return <FileSpreadsheet className="w-5 h-5 text-emerald-500" />;
      case "pdf":
        return <File className="w-5 h-5 text-rose-500" />;
      case "ppt":
        return <FileText className="w-5 h-5 text-orange-500" />;
      default:
        return <File className="w-5 h-5 text-slate-500" />;
    }
  };

  const filteredFiles = files.filter(f =>
    f.fileName.toLowerCase().includes(searchTerm.toLowerCase()) ||
    f.filePath.toLowerCase().includes(searchTerm.toLowerCase())
  );

  return (
    <div className="flex flex-col flex-1 h-full overflow-hidden p-6 bg-background">
      {/* Top Banner Message */}
      {error && (
        <div className="mb-4 p-4 rounded-xl bg-red-50 text-red-700 text-sm font-semibold flex items-center gap-2.5 animate-in slide-in-from-top-4 duration-300">
          <AlertTriangle className="w-4 h-4 shrink-0" />
          <span>{error}</span>
        </div>
      )}
      {successMsg && (
        <div className="mb-4 p-4 rounded-xl bg-emerald-50 text-emerald-700 text-sm font-semibold flex items-center gap-2.5 animate-in slide-in-from-top-4 duration-300">
          <Check className="w-4 h-4 shrink-0" />
          <span>{successMsg}</span>
        </div>
      )}

      {/* Header Panel */}
      <div className="mb-6 bg-muted/30 p-5 rounded-2xl border border-border/40 flex flex-col gap-4">
        <div className="flex items-start justify-between gap-4">
          <div className="flex flex-col gap-1 min-w-0">
            <h2 className="text-sm font-extrabold text-foreground flex items-center gap-2">
              <FolderOpen className="w-4 h-4 text-primary shrink-0" />
              本地工作目录
            </h2>
            <p className="text-xs text-secondary-foreground leading-normal">
              {project?.folder_path ? (
                <span className="block">
                  <span className="font-semibold text-foreground">当前绑定:</span>
                  <code className="mt-1 block w-full max-w-full px-2.5 py-1.5 rounded bg-muted/80 border border-border/50 font-mono text-primary text-[10px] leading-5 select-all whitespace-normal break-all">
                    {project.folder_path}
                  </code>
                  {project.linked_folder_type && (
                    <div className="mt-1.5 flex flex-col gap-1">
                      <div>
                        <span className={`inline-flex items-center px-2 py-0.5 rounded text-[10px] font-bold ${
                          project.linked_folder_type === 'internal'
                            ? 'bg-emerald-500/10 text-emerald-600 border border-emerald-200/20'
                            : project.linked_folder_type === 'external'
                              ? 'bg-amber-500/10 text-amber-600 border border-amber-200/20'
                              : 'bg-slate-100 text-slate-500'
                        }`}>
                          {project.linked_folder_type === 'internal' ? '工作区内部关联 (可迁移)' : '工作区外部关联'}
                        </span>
                      </div>
                      {project.linked_folder_type === 'external' && (
                        <p className="text-[10px] font-medium text-amber-600 leading-normal">
                          ⚠️ 该文件夹位于当前工作区外部，复制或导出整个 Workspace 时不会自动迁移。
                        </p>
                      )}
                    </div>
                  )}
                </span>
              ) : (
                "绑定项目文件夹后，可一键扫描同步该文件夹下的效益测算文件，且支持直接定位和打开文件。"
              )}
            </p>
          </div>

          {/* Main Folder Actions if bound */}
          {project?.folder_path && (
            <button
              onClick={handleOpenFolder}
              className="shrink-0 px-3 py-1.5 bg-primary/10 hover:bg-blue-50 text-primary rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all"
              title="在系统文件管理器中打开此文件夹"
            >
              <FolderOpen className="w-3.5 h-3.5" />
              打开目录
            </button>
          )}
        </div>

        {/* Directory & File Management Buttons */}
        <div className="flex flex-wrap gap-2 pt-2 border-t border-border/40 justify-start items-center">
          {project?.folder_path ? (
            <>
              <button
                onClick={handleScanFolder}
                disabled={loading}
                className="px-3.5 py-1.5 bg-primary text-primary-foreground hover:bg-primary/95 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all disabled:opacity-50 shadow-sm"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${loading ? 'animate-spin' : ''}`} />
                同步扫描
              </button>
              <button
                onClick={handleBindFolder}
                className="px-3.5 py-1.5 bg-secondary text-secondary-foreground hover:bg-card border border-input rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all"
              >
                更换绑定
              </button>
              <button
                onClick={handleCreateFolderClick}
                disabled={loading}
                className="px-3.5 py-1.5 bg-secondary text-secondary-foreground hover:bg-card border border-input rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all disabled:opacity-50"
              >
                <FolderPlus className="w-3.5 h-3.5" />
                新建并绑定
              </button>
              <button
                onClick={handleUnbindFolder}
                className="px-3.5 py-1.5 hover:bg-red-500/10 text-red-500 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all"
              >
                解绑目录
              </button>
            </>
          ) : (
            <>
              <button
                onClick={handleBindFolder}
                disabled={loading}
                className="px-3.5 py-1.5 bg-primary text-primary-foreground hover:opacity-90 rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-sm transition-all disabled:opacity-50"
              >
                <Folder className="w-3.5 h-3.5" />
                关联本地目录
              </button>
              <button
                onClick={handleCreateFolderClick}
                disabled={loading}
                className="px-3.5 py-1.5 bg-secondary text-secondary-foreground hover:bg-card border border-input rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all disabled:opacity-50"
              >
                <FolderPlus className="w-3.5 h-3.5" />
                新建项目文件夹
              </button>
            </>
          )}

          {/* Spacer */}
          <div className="flex-1 min-w-[10px]" />

          {/* Add File Button */}
          <div className="relative group">
            <button
              disabled={loading}
              className="px-3.5 py-1.5 bg-foreground text-background hover:opacity-90 rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-sm transition-all disabled:opacity-50"
            >
              <Plus className="w-3.5 h-3.5" />
              添加文件
            </button>

            {/* Dropdown Menu */}
            <div className="absolute right-0 mt-1 w-44 bg-popover border border-border rounded-xl shadow-lg opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-200 z-50 p-1 flex flex-col gap-0.5">
              <button
                onClick={() => handleAddFile('linked')}
                className="w-full text-left px-2.5 py-1.5 text-xs font-semibold text-foreground hover:bg-muted rounded-lg transition-colors flex items-center gap-1.5"
              >
                <Pin className="w-3 h-3 text-primary rotate-45" />
                关联本地文件 (Linked)
              </button>
              <button
                onClick={() => handleAddFile('copied')}
                className="w-full text-left px-2.5 py-1.5 text-xs font-semibold text-foreground hover:bg-muted rounded-lg transition-colors flex items-center gap-1.5"
              >
                <Plus className="w-3 h-3 text-emerald-500" />
                导入托管文件 (Copied)
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* Filter and Stats Panel */}
      <div className="mb-4 flex flex-col sm:flex-row items-center justify-between gap-4">
        <div className="relative w-full sm:w-80">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
          <input
            type="text"
            placeholder="搜索文件名称或路径..."
            value={searchTerm}
            onChange={e => setSearchTerm(e.target.value)}
            className="w-full pl-9 pr-4 py-2 bg-card border border-input rounded-xl text-xs outline-none focus:bg-card focus:border-ring transition-all text-foreground"
          />
        </div>

        <div className="flex gap-4 text-xs font-medium text-secondary-foreground opacity-80">
          <span>文件总数: <strong className="text-foreground">{files.length}</strong></span>
          <span>已失效文件: <strong className="text-red-500">{files.filter(f => !f.exists).length}</strong></span>
          <span>托管文件: <strong className="text-emerald-500">{files.filter(f => f.storageMode === 'copied').length}</strong></span>
        </div>
      </div>

      {/* File List Table */}
      <div className="flex-1 overflow-y-auto rounded-2xl bg-muted/10 border border-transparent">
        {filteredFiles.length === 0 ? (
          <div className="h-full flex flex-col items-center justify-center text-center p-8 gap-3">
            <File className="w-12 h-12 text-muted-foreground opacity-30" />
            <p className="text-sm font-semibold text-secondary-foreground opacity-60">
              {searchTerm ? "未找到符合条件的文件" : "暂无项目文件，请通过绑定目录扫描或手动添加"}
            </p>
          </div>
        ) : (
            <div className="w-full min-w-[780px] flex flex-col">
              {/* Table Header */}
              <div className="flex items-center px-6 py-3.5 border-b border-border bg-muted/20 text-xs font-bold text-secondary-foreground uppercase tracking-wider select-none shrink-0">
              <div className="w-[34%]">文件名称</div>
              <div className="w-[14%]">类型/大小</div>
              <div className="w-[17%]">存储模式</div>
              <div className="w-[16%]">角色标识</div>
              <div className="w-[19%] text-right">操作</div>
            </div>

            {/* Table Body */}
            <div className="flex flex-col divide-y divide-border/40">
              {filteredFiles.map(file => {
                const isLinkedMissing = !file.exists;
                return (
                  <div
                    key={file.id}
                    className={`flex items-center px-6 py-4 hover:bg-muted/30 transition-all duration-150 ${isLinkedMissing ? 'bg-red-500/[0.02]' : ''}`}
                  >
                    {/* File Name & Path */}
                    <div className="w-[34%] flex items-start gap-3 min-w-0 pr-4">
                      <div className="mt-0.5 shrink-0">
                        {isLinkedMissing ? (
                          <FileWarning className="w-5 h-5 text-red-500 animate-pulse" />
                        ) : (
                          getFileIcon(file.fileType)
                        )}
                      </div>
                      <div className="flex flex-col min-w-0">
                        <span
                          onClick={() => file.exists && handleOpenFile(file)}
                          className={`text-xs font-semibold text-foreground truncate cursor-pointer ${file.exists ? 'hover:text-primary hover:underline' : 'opacity-60 cursor-not-allowed'}`}
                          title={file.fileName}
                        >
                          {file.fileName}
                        </span>
                        <span
                          className="text-[10px] text-muted-foreground break-all whitespace-normal leading-4 select-all mt-1 font-mono"
                          title={file.filePath}
                        >
                          {file.filePath}
                        </span>
                      </div>
                    </div>

                    {/* File Type and Size */}
                    <div className="w-[14%] flex flex-col gap-1">
                      <span className="text-[11px] font-medium text-foreground uppercase">
                        {file.extension || "未知"}
                      </span>
                      <span className="text-[10px] text-secondary-foreground opacity-75">
                        {formatSize(file.size)}
                      </span>
                    </div>

                    {/* Storage Mode & Existence */}
                    <div className="w-[17%] flex flex-col items-start gap-1.5">
                      <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
                        file.storageMode === 'copied'
                          ? 'bg-emerald-500/10 text-emerald-500'
                           : 'bg-primary/10 text-primary'
                      }`}>
                        {file.storageMode === 'copied' ? '托管 (Copied)' : '关联 (Linked)'}
                      </span>
                      {isLinkedMissing && (
                        <span className="flex items-center gap-1 text-[10px] text-red-500 font-bold">
                          <AlertTriangle className="w-3 h-3" />
                          已丢失 (未在磁盘找到)
                        </span>
                      )}
                    </div>

                    {/* Indicators */}
                    <div className="w-[16%] flex flex-col gap-1.5 items-start">
                      {file.isMainDocument && (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-blue-50 text-blue-700 text-[10px] font-bold">
                          <Pin className="w-3 h-3 rotate-45 shrink-0" />
                          效益测算主文档
                        </span>
                      )}
                      {file.isMainBudgetFile && (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-orange-500/10 text-orange-500 text-[10px] font-bold">
                          <Pin className="w-3 h-3 rotate-45 shrink-0" />
                          预算主文件
                        </span>
                      )}
                      {!file.isMainDocument && !file.isMainBudgetFile && (
                        <span className="text-[10px] text-secondary-foreground opacity-50 italic">
                          普通文件
                        </span>
                      )}
                    </div>

                    {/* Actions */}
                    <div className="w-[19%] flex items-center justify-end gap-1.5 shrink-0">
                      {/* Import Benefit data from Excel */}
                      {file.exists && file.fileType === 'excel' && (
                        <button
                          onClick={() => {
                            if (window.confirm(`确定要从 Excel "${file.fileName}" 导入效益测算数据吗？这将会覆盖当前项目已有的首选效益方案测算参数。`)) {
                              importExcelData(file.filePath);
                            }
                          }}
                          title="导入测算数据"
                          className="group/import relative inline-flex h-8 w-8 items-center justify-center rounded-lg border border-emerald-200 bg-emerald-50 text-emerald-700 shadow-sm transition-all hover:-translate-y-px hover:border-emerald-300 hover:bg-emerald-100 hover:text-emerald-800 hover:shadow disabled:opacity-50 disabled:hover:translate-y-0"
                          disabled={loading}
                        >
                          <FileSpreadsheet className="w-3.5 h-3.5 shrink-0" />
                          <span className="pointer-events-none absolute -top-8 right-0 hidden whitespace-nowrap rounded-md bg-foreground px-2 py-1 text-[10px] font-bold text-background shadow-md group-hover/import:block">
                            导入测算
                          </span>
                        </button>
                      )}

                      {/* Set role pins */}
                      {file.exists && (
                        <>
                          <button
                            onClick={() => handleToggleMainDoc(file)}
                            title={file.isMainDocument ? "取消设定为效益测算主文档" : "设定为效益测算主文档"}
                            className={`p-1.5 rounded-lg hover:bg-muted text-secondary-foreground transition-all ${file.isMainDocument ? 'text-blue-700 bg-blue-50' : 'opacity-60 hover:opacity-100'}`}
                          >
                            <Pin className="w-3.5 h-3.5 rotate-45" />
                          </button>
                          <button
                            onClick={() => handleToggleMainBudget(file)}
                            title={file.isMainBudgetFile ? "取消设定为项目预算主文件" : "设定为项目预算主文件"}
                            className={`p-1.5 rounded-lg hover:bg-muted text-secondary-foreground transition-all ${file.isMainBudgetFile ? 'text-orange-500 bg-orange-500/10' : 'opacity-60 hover:opacity-100'}`}
                          >
                            <Pin className="w-3.5 h-3.5" />
                          </button>
                        </>
                      )}

                      {/* Open file operations */}
                      {file.exists && (
                        <>
                          <button
                            onClick={() => handleOpenFile(file)}
                            title="系统打开"
                            className="p-1.5 rounded-lg hover:bg-muted text-secondary-foreground opacity-60 hover:opacity-100 transition-all"
                          >
                            <Eye className="w-3.5 h-3.5" />
                          </button>
                          <button
                            onClick={() => handleRevealFile(file)}
                            title="定位文件"
                            className="p-1.5 rounded-lg hover:bg-muted text-secondary-foreground opacity-60 hover:opacity-100 transition-all"
                          >
                            <ExternalLink className="w-3.5 h-3.5" />
                          </button>
                        </>
                      )}

                      {/* Delete / Unlink file */}
                      <button
                        onClick={() => handleDeleteFile(file)}
                        title={file.storageMode === 'copied' ? "删除托管文件" : "取消项目关联"}
                        className="p-1.5 rounded-lg hover:bg-red-500/10 text-secondary-foreground hover:text-red-500 opacity-60 hover:opacity-100 transition-all"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {pendingBindFolder && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-950/35 p-4 backdrop-blur-sm">
          <div className="w-full max-w-md rounded-xl border border-border bg-card p-5 shadow-xl">
            <div className="flex items-start gap-3">
              <div className="mt-0.5 flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-primary/10 text-primary">
                <FolderOpen className="h-4 w-4" />
              </div>
              <div className="min-w-0 flex-1">
                <h3 className="text-sm font-extrabold text-foreground">确认更换绑定目录</h3>
                <p className="mt-1 text-xs leading-5 text-secondary-foreground">
                  新目录名称为
                  <span className="mx-1 font-bold text-foreground">「{pendingBindFolder.folderName}」</span>
                  ，与当前项目名称
                  <span className="mx-1 font-bold text-foreground">「{project?.name || "未命名项目"}」</span>
                  不一致。请选择是否同步项目名称。
                </p>
                <code className="mt-3 block rounded-lg border border-border/60 bg-muted/50 px-3 py-2 font-mono text-[10px] leading-4 text-primary break-all">
                  {pendingBindFolder.folderPath}
                </code>
              </div>
            </div>

            <div className="mt-5 flex flex-wrap justify-end gap-2">
              <button
                type="button"
                onClick={() => setPendingBindFolder(null)}
                disabled={loading}
                className="rounded-lg border border-border bg-card px-3 py-2 text-xs font-bold text-secondary-foreground transition-all hover:bg-muted disabled:opacity-50"
              >
                取消
              </button>
              <button
                type="button"
                onClick={() => executeBindFolder(pendingBindFolder.folderPath, false)}
                disabled={loading}
                className="rounded-lg border border-border bg-secondary px-3 py-2 text-xs font-bold text-secondary-foreground transition-all hover:bg-muted disabled:opacity-50"
              >
                仅绑定目录
              </button>
              <button
                type="button"
                onClick={() => executeBindFolder(pendingBindFolder.folderPath, true)}
                disabled={loading}
                className="inline-flex items-center gap-1.5 rounded-lg bg-primary px-3 py-2 text-xs font-bold text-primary-foreground shadow-sm transition-all hover:opacity-95 disabled:opacity-50"
              >
                <Check className="h-3.5 w-3.5" />
                绑定并同步项目名称
              </button>
            </div>
          </div>
        </div>
      )}

      {notInRootFolder && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-950/35 p-4 backdrop-blur-sm">
          <div className="w-full max-w-md rounded-xl border border-border bg-card p-5 shadow-xl">
            <div className="flex items-start gap-3">
              <div className="mt-0.5 flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-amber-500/10 text-amber-600">
                <AlertTriangle className="h-4 w-4" />
              </div>
              <div className="min-w-0 flex-1">
                <h3 className="text-sm font-extrabold text-foreground">未关联项目根目录</h3>
                <p className="mt-1 text-xs leading-5 text-secondary-foreground">
                  所选文件夹不在当前已注册的项目根目录下。为保证项目的“路径韧性”（在不同电脑或移动目录后自动识别），建议将其添加为根目录。
                </p>
                <code className="mt-3 block rounded-lg border border-border/60 bg-muted/50 px-3 py-2 font-mono text-[10px] leading-4 text-primary break-all">
                  {notInRootFolder.folderPath}
                </code>
              </div>
            </div>

            <div className="mt-5 flex flex-wrap justify-end gap-2">
              <button
                type="button"
                onClick={() => setNotInRootFolder(null)}
                disabled={loading}
                className="rounded-lg border border-border bg-card px-3 py-2 text-xs font-bold text-secondary-foreground transition-all hover:bg-muted disabled:opacity-50"
              >
                取消
              </button>
              <button
                type="button"
                onClick={() => {
                  executeBindFolder(notInRootFolder.folderPath, notInRootFolder.renameProject, "absolute_only");
                  setNotInRootFolder(null);
                }}
                disabled={loading}
                className="rounded-lg border border-border bg-secondary px-3 py-2 text-xs font-bold text-secondary-foreground transition-all hover:bg-muted disabled:opacity-50"
              >
                仅保留绝对路径
              </button>
              <button
                type="button"
                onClick={() => {
                  executeBindFolder(notInRootFolder.folderPath, notInRootFolder.renameProject, "create_root");
                  setNotInRootFolder(null);
                }}
                disabled={loading}
                className="inline-flex items-center gap-1.5 rounded-lg bg-primary px-3 py-2 text-xs font-bold text-primary-foreground shadow-sm transition-all hover:opacity-95 disabled:opacity-50"
              >
                <Check className="h-3.5 w-3.5" />
                注册为新根目录 (推荐)
              </button>
            </div>
          </div>
        </div>
      )}

      {isCreateFolderModalOpen && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-950/35 p-4 backdrop-blur-sm">
          <form
            onSubmit={(event) => {
              event.preventDefault();
              handleConfirmCreateFolder();
            }}
            className="w-full max-w-md rounded-xl border border-border bg-card p-5 shadow-xl"
          >
            <div className="flex items-start gap-3">
              <div className="mt-0.5 flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-emerald-500/10 text-emerald-600">
                <FolderPlus className="h-4 w-4" />
              </div>
              <div className="min-w-0 flex-1">
                <h3 className="text-sm font-extrabold text-foreground">新建项目文件夹</h3>
                <p className="mt-1 text-xs leading-5 text-secondary-foreground">
                  系统会在所选父级目录下创建文件夹，并在创建成功后自动绑定到当前项目。
                </p>
                <div className="mt-3">
                  <label className="text-[11px] font-bold text-secondary-foreground">父级目录</label>
                  <code className="mt-1 block rounded-lg border border-border/60 bg-muted/50 px-3 py-2 font-mono text-[10px] leading-4 text-primary break-all">
                    {createFolderParentPath}
                  </code>
                </div>
                <div className="mt-3">
                  <label htmlFor="create-project-folder-name" className="text-[11px] font-bold text-secondary-foreground">
                    文件夹名称
                  </label>
                  <input
                    id="create-project-folder-name"
                    autoFocus
                    value={createFolderName}
                    onChange={(event) => setCreateFolderName(event.target.value)}
                    disabled={loading}
                    className="mt-1 w-full rounded-lg border border-input bg-card px-3 py-2 text-sm font-semibold text-foreground outline-none transition-all focus:border-ring focus:ring-2 focus:ring-ring/20 disabled:opacity-50"
                    placeholder="请输入文件夹名称"
                  />
                </div>
              </div>
            </div>

            <div className="mt-5 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => {
                  setIsCreateFolderModalOpen(false);
                  setCreateFolderParentPath(null);
                  setCreateFolderName("");
                }}
                disabled={loading}
                className="rounded-lg border border-border bg-card px-3 py-2 text-xs font-bold text-secondary-foreground transition-all hover:bg-muted disabled:opacity-50"
              >
                取消
              </button>
              <button
                type="submit"
                disabled={loading || !createFolderName.trim()}
                className="inline-flex items-center gap-1.5 rounded-lg bg-primary px-3 py-2 text-xs font-bold text-primary-foreground shadow-sm transition-all hover:opacity-95 disabled:opacity-50"
              >
                <FolderPlus className="h-3.5 w-3.5" />
                创建并绑定
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
}
