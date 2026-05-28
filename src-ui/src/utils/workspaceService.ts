import { invoke } from "@tauri-apps/api/core";

export interface WorkspaceManifest {
  app: string;
  workspaceVersion: number;
  workspaceId: string;
  name: string;
  createdAt: string;
  lastOpenedAt: string;
}

export interface WorkspaceInfo {
  workspaceRoot: string;
  workspaceName: string;
  workspaceId: string;
  manifest: WorkspaceManifest;
}

export interface RecentWorkspace {
  path: string;
  name: string;
  workspaceId: string;
  lastOpenedAt: string;
}

export interface WorkspaceStateDto {
  currentWorkspace: WorkspaceInfo | null;
  recentWorkspaces: RecentWorkspace[];
  isWorkspaceReady: boolean;
  startupError?: { code: string; message: string } | null;
}

export interface WorkspacePathStatus {
  status: "workspace" | "legacySuspected" | "emptyOrInitializable" | "nonEmptyNonWorkspace" | "importablePlainDirectory";
  message?: string | null;
}

export interface InitializeWorkspaceOptions {
  workspaceName?: string;
  selectedDirectories: string[];
  createProjectJson?: boolean;
  createSubDirs?: boolean;
}

export function parseWorkspaceError(error: unknown) {
  const fallback = String(error || "工作区操作失败");
  try {
    const parsed = JSON.parse(fallback);
    if (parsed?.code && parsed?.message) {
      return parsed as { code: string; message: string };
    }
  } catch {
    // fall through
  }
  return { code: "Unknown", message: fallback };
}

export const workspaceService = {
  getState() {
    return invoke<WorkspaceStateDto>("get_workspace_state");
  },
  selectFolder() {
    return invoke<string | null>("select_workspace_folder");
  },
  inspectPath(path: string) {
    return invoke<WorkspacePathStatus>("inspect_workspace_path", { path });
  },
  create(path: string, name?: string, allowNonEmpty = false) {
    return invoke<WorkspaceInfo>("create_workspace", { path, name: name || null, allowNonEmpty });
  },
  open(path: string) {
    return invoke<WorkspaceInfo>("open_workspace", { path });
  },
  initializeFromExisting(path: string, options: InitializeWorkspaceOptions) {
    return invoke<WorkspaceInfo>("initialize_workspace_from_existing_directory", { path, options });
  },
  scanAndImportAllCalculations() {
    return invoke<number>("scan_and_import_all_workspace_calculations");
  },
};
