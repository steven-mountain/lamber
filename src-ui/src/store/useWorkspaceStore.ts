import { create } from "zustand";
import { parseWorkspaceError, workspaceService, type RecentWorkspace, type WorkspaceInfo } from "../utils/workspaceService";
import { useNavigationStore } from "./useNavigationStore";
import { useProjectStore } from "./useProjectStore";

interface WorkspaceStore {
  currentWorkspace: WorkspaceInfo | null;
  workspaceRoot: string | null;
  workspaceName: string | null;
  workspaceId: string | null;
  recentWorkspaces: RecentWorkspace[];
  isWorkspaceReady: boolean;
  isLoading: boolean;
  error: string | null;
  refreshWorkspaceState: () => Promise<void>;
  selectAndOpenWorkspace: () => Promise<void>;
  selectAndCreateWorkspace: () => Promise<void>;
  openRecentWorkspace: (path: string) => Promise<void>;
  forgetWorkspace: (path: string) => Promise<void>;
  closeCurrentWorkspace: () => Promise<void>;
  initializeWorkspaceFromExisting: (path: string, options: import("../utils/workspaceService").InitializeWorkspaceOptions) => Promise<void>;
  scanAndImportAllWorkspaceCalculations: () => Promise<number>;
}

function clearWorkspaceScopedProjectContext() {
  useProjectStore.getState().clearCurrentProject();
  useNavigationStore.getState().clearContext();

  if (typeof window !== "undefined") {
    window.localStorage.removeItem("lamber_active_project_id");
    window.localStorage.removeItem("lamber_active_scheme_id");
    window.localStorage.removeItem("lamber_new_scheme_name");
  }
}

function applyWorkspace(workspace: WorkspaceInfo | null) {
  if (!workspace) {
    clearWorkspaceScopedProjectContext();
  }
  return {
    currentWorkspace: workspace,
    workspaceRoot: workspace?.workspaceRoot || null,
    workspaceName: workspace?.workspaceName || null,
    workspaceId: workspace?.workspaceId || null,
    isWorkspaceReady: Boolean(workspace),
  };
}

export const useWorkspaceStore = create<WorkspaceStore>((set, get) => ({
  currentWorkspace: null,
  workspaceRoot: null,
  workspaceName: null,
  workspaceId: null,
  recentWorkspaces: [],
  isWorkspaceReady: false,
  isLoading: false,
  error: null,

  refreshWorkspaceState: async () => {
    set({ isLoading: true });
    try {
      const state = await workspaceService.getState();
      const prevWorkspace = get().currentWorkspace;
      
      // If we are clearing or changing to a different workspace, clear project
      if (!state.currentWorkspace || (prevWorkspace && prevWorkspace.workspaceId !== state.currentWorkspace.workspaceId)) {
        clearWorkspaceScopedProjectContext();
      }
      
      set({
        ...applyWorkspace(state.currentWorkspace),
        recentWorkspaces: state.recentWorkspaces,
        error: state.startupError?.message || null,
        isLoading: false,
      });
    } catch (error) {
      clearWorkspaceScopedProjectContext();
      set({ ...applyWorkspace(null), error: parseWorkspaceError(error).message, isLoading: false });
    }
  },

  selectAndOpenWorkspace: async () => {
    const path = await workspaceService.selectFolder();
    if (!path) return;
    set({ isLoading: true, error: null });
    try {
      clearWorkspaceScopedProjectContext();
      const workspace = await workspaceService.open(path);
      await get().refreshWorkspaceState();
      set({ ...applyWorkspace(workspace), isLoading: false });
    } catch (error) {
      set({ error: parseWorkspaceError(error).message, isLoading: false });
    }
  },

  selectAndCreateWorkspace: async () => {
    const path = await workspaceService.selectFolder();
    if (!path) return;
    set({ isLoading: true, error: null });
    try {
      const status = await workspaceService.inspectPath(path);
      if (status.status === "legacySuspected") {
        set({ error: status.message || "疑似旧版数据目录，本阶段不会覆盖。", isLoading: false });
        return;
      }
      const allowNonEmpty = status.status === "nonEmptyNonWorkspace"
        ? confirm("该目录非空且不是 Lamber 工作区。确认要在此目录初始化工作区吗？")
        : false;
      if (status.status === "nonEmptyNonWorkspace" && !allowNonEmpty) {
        set({ isLoading: false });
        return;
      }
      clearWorkspaceScopedProjectContext();
      const workspace = await workspaceService.create(path, undefined, allowNonEmpty);
      await get().refreshWorkspaceState();
      set({ ...applyWorkspace(workspace), isLoading: false });
    } catch (error) {
      set({ error: parseWorkspaceError(error).message, isLoading: false });
    }
  },

  openRecentWorkspace: async (path: string) => {
    set({ isLoading: true, error: null });
    try {
      clearWorkspaceScopedProjectContext();
      const workspace = await workspaceService.open(path);
      await get().refreshWorkspaceState();
      set({ ...applyWorkspace(workspace), isLoading: false });
    } catch (error) {
      set({ error: parseWorkspaceError(error).message, isLoading: false });
    }
  },

  forgetWorkspace: async (path: string) => {
    set({ isLoading: true, error: null });
    try {
      const prevWorkspace = get().currentWorkspace;
      const state = await workspaceService.forget(path);
      if (!state.currentWorkspace || (prevWorkspace && prevWorkspace.workspaceId !== state.currentWorkspace.workspaceId)) {
        clearWorkspaceScopedProjectContext();
      }
      set({
        ...applyWorkspace(state.currentWorkspace),
        recentWorkspaces: state.recentWorkspaces,
        error: state.startupError?.message || null,
        isLoading: false,
      });
    } catch (error) {
      set({ error: parseWorkspaceError(error).message, isLoading: false });
      throw error;
    }
  },

  closeCurrentWorkspace: async () => {
    set({ isLoading: true, error: null });
    try {
      const state = await workspaceService.closeCurrent();
      clearWorkspaceScopedProjectContext();
      set({
        ...applyWorkspace(state.currentWorkspace),
        recentWorkspaces: state.recentWorkspaces,
        error: state.startupError?.message || null,
        isLoading: false,
      });
    } catch (error) {
      set({ error: parseWorkspaceError(error).message, isLoading: false });
      throw error;
    }
  },

  initializeWorkspaceFromExisting: async (path: string, options: import("../utils/workspaceService").InitializeWorkspaceOptions) => {
    set({ isLoading: true, error: null });
    try {
      clearWorkspaceScopedProjectContext();
      const workspace = await workspaceService.initializeFromExisting(path, options);
      await get().refreshWorkspaceState();
      set({ ...applyWorkspace(workspace), isLoading: false });
    } catch (error) {
      set({ error: parseWorkspaceError(error).message, isLoading: false });
      throw error;
    }
  },

  scanAndImportAllWorkspaceCalculations: async () => {
    set({ isLoading: true, error: null });
    try {
      const count = await workspaceService.scanAndImportAllCalculations();
      set({ isLoading: false });
      return count;
    } catch (error) {
      set({ error: parseWorkspaceError(error).message, isLoading: false });
      throw error;
    }
  },
}));
