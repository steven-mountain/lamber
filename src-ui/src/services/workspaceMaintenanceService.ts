import { invoke } from "@tauri-apps/api/core";
import type { WorkspaceInfo } from "../utils/workspaceService";

export interface HealthCheckItem {
  id: string;
  severity: "info" | "warning" | "error";
  category: string;
  message: string;
  detail?: string | null;
  repairable: boolean;
  repairAction?: string | null;
}

export interface HealthCheckResult {
  status: "normal" | "warning" | "error";
  items: HealthCheckItem[];
}

export interface WorkspaceBackupInfo {
  id: string;
  fileName: string;
  path: string;
  createdAt: string;
  sizeBytes: number;
  isDaily: boolean;
}

export interface WorkspaceExportOptions {
  includeBackups?: boolean;
  includeExports?: boolean;
  allowWarnings?: boolean;
}

export interface ExportWorkspaceResult {
  archivePath: string;
  databaseBackupPath: string;
  manifest: {
    exportedAt: string;
    workspaceId: string;
    workspaceName: string;
    workspaceVersion: number;
    externalPathCount: number;
    warnings: string[];
  };
  warnings: HealthCheckItem[];
}

export interface ArchiveValidationResult {
  valid: boolean;
  rootPrefix?: string | null;
  workspaceId?: string | null;
  workspaceName?: string | null;
  workspaceVersion?: number | null;
  errors: string[];
  warnings: string[];
}

export interface ImportWorkspaceResult {
  workspaceRoot: string;
  opened: boolean;
  workspace?: WorkspaceInfo | null;
  warnings: HealthCheckItem[];
}

export interface ImportWorkspaceOptions {
  openAfterImport?: boolean;
  conflictStrategy?: "rename" | "overwrite" | "cancel";
  destinationName?: string | null;
}

export interface ExternalPathInfo {
  path: string;
  projectId?: string | null;
  projectName?: string | null;
  pathType: string;
  exists: boolean;
  impact: string;
}

export interface PathConversionCandidate {
  id: string;
  tableName: string;
  recordId: string;
  columnName: string;
  projectId?: string | null;
  currentPath: string;
  relativePath: string;
  reason: string;
}

export interface PathConversionResult {
  dryRun: boolean;
  candidates: PathConversionCandidate[];
  applied: number;
  backupPath?: string | null;
}

export const workspaceMaintenanceService = {
  exportWorkspace(targetPath: string | null, options: WorkspaceExportOptions) {
    return invoke<ExportWorkspaceResult>("export_workspace", { targetPath, options });
  },

  validateArchive(zipPath: string) {
    return invoke<ArchiveValidationResult>("validate_workspace_archive", { zipPath });
  },

  importWorkspace(
    zipPath: string,
    targetDir: string,
    openAfterImport = false,
    conflictStrategy: ImportWorkspaceOptions["conflictStrategy"] = "rename",
    destinationName: string | null = null,
  ) {
    return invoke<ImportWorkspaceResult>("import_workspace", {
      zipPath,
      targetDir,
      openAfterImport,
      conflictStrategy,
      destinationName,
    });
  },

  revealInFileManager(path: string) {
    return invoke<void>("reveal_in_file_manager", { path });
  },

  createBackup() {
    return invoke<WorkspaceBackupInfo>("create_workspace_backup");
  },

  listBackups() {
    return invoke<WorkspaceBackupInfo[]>("list_workspace_backups");
  },

  restoreBackup(backupId: string) {
    return invoke<WorkspaceInfo>("restore_workspace_backup", { backupId });
  },

  deleteBackup(backupId: string) {
    return invoke<void>("delete_workspace_backup", { backupId });
  },

  runHealthCheck() {
    return invoke<HealthCheckResult>("run_workspace_health_check");
  },

  repairIssues(issueIds: string[]) {
    return invoke<{ repaired: number; backupPath?: string | null; health: HealthCheckResult }>("repair_workspace_issues", { issueIds });
  },

  listExternalPaths() {
    return invoke<ExternalPathInfo[]>("list_external_paths");
  },

  inspectPaths() {
    return invoke<[PathConversionCandidate[], ExternalPathInfo[]]>("inspect_workspace_paths");
  },

  convertInternalAbsolutePathsToRelative(dryRun = true) {
    return invoke<PathConversionResult>("convert_internal_absolute_paths_to_relative", { dryRun });
  },
};
