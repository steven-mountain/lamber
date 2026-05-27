import { create } from "zustand";
import { parseWorkspaceError, workspaceService, type RecentWorkspace, type WorkspaceInfo } from "../utils/workspaceService";
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
}

function applyWorkspace(workspace: WorkspaceInfo | null) {
  if (!workspace) {
    useProjectStore.getState().clearCurrentProject();
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
        useProjectStore.getState().clearCurrentProject();
      }
      
      set({
        ...applyWorkspace(state.currentWorkspace),
        recentWorkspaces: state.recentWorkspaces,
        error: state.startupError?.message || null,
        isLoading: false,
      });
    } catch (error) {
      useProjectStore.getState().clearCurrentProject();
      set({ ...applyWorkspace(null), error: parseWorkspaceError(error).message, isLoading: false });
    }
  },

  selectAndOpenWorkspace: async () => {
    const path = await workspaceService.selectFolder();
    if (!path) return;
    set({ isLoading: true, error: null });
    try {
      useProjectStore.getState().clearCurrentProject();
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
      useProjectStore.getState().clearCurrentProject();
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
      useProjectStore.getState().clearCurrentProject();
      const workspace = await workspaceService.open(path);
      await get().refreshWorkspaceState();
      set({ ...applyWorkspace(workspace), isLoading: false });
    } catch (error) {
      set({ error: parseWorkspaceError(error).message, isLoading: false });
    }
  },
}));
