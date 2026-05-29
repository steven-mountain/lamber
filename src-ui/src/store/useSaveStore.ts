import { create } from "zustand";
import { useProjectStore } from "./useProjectStore";
import { useWorkspaceStore } from "./useWorkspaceStore";

export type DirtyScope =
  | "project-detail"
  | "lifecycle"
  | "cashflow"
  | "benefit-analysis"
  | "template-forms";

export interface SaveContext {
  workspaceId: string;
  workspaceRoot: string;
  projectId: string;
  dirtyScopes: DirtyScope[];
}

export interface SaveHandlerResult {
  success: boolean;
  savedScopes: DirtyScope[];
  error?: string;
}

export type SaveHandler = (context: SaveContext) => Promise<SaveHandlerResult>;

interface SaveState {
  isDirty: boolean;
  dirtyScopes: DirtyScope[];
  isSaving: boolean;
  lastSavedAt: string | null;
  lastSaveError: string | null;
  handlers: Partial<Record<DirtyScope, SaveHandler>>;
  markDirty: (scope: DirtyScope) => void;
  clearDirty: (scope: DirtyScope) => void;
  clearDirtyScopes: (scopes: DirtyScope[]) => void;
  clearAllDirty: () => void;
  hasUnsavedChanges: () => boolean;
  registerSaveHandler: (scope: DirtyScope, handler: SaveHandler) => void;
  unregisterSaveHandler: (scope: DirtyScope) => void;
  saveCurrentProject: () => Promise<void>;
}

function uniqueScopes(scopes: DirtyScope[]) {
  return Array.from(new Set(scopes));
}

function normalizeSavedScopes(scopes: DirtyScope[], dirtySnapshot: DirtyScope[]) {
  const dirtySet = new Set(dirtySnapshot);
  return uniqueScopes(scopes).filter(scope => dirtySet.has(scope));
}

export const useSaveStore = create<SaveState>((set, get) => ({
  isDirty: false,
  dirtyScopes: [],
  isSaving: false,
  lastSavedAt: null,
  lastSaveError: null,
  handlers: {},

  markDirty: (scope) => {
    const currentScopes = get().dirtyScopes;
    if (currentScopes.includes(scope)) return;
    const nextScopes = [...currentScopes, scope];
    set({ dirtyScopes: nextScopes, isDirty: true, lastSaveError: null });
  },

  clearDirty: (scope) => {
    const nextScopes = get().dirtyScopes.filter(item => item !== scope);
    set({ dirtyScopes: nextScopes, isDirty: nextScopes.length > 0 });
  },

  clearDirtyScopes: (scopes) => {
    const remove = new Set(scopes);
    const nextScopes = get().dirtyScopes.filter(scope => !remove.has(scope));
    set({ dirtyScopes: nextScopes, isDirty: nextScopes.length > 0 });
  },

  clearAllDirty: () => set({ dirtyScopes: [], isDirty: false, lastSaveError: null }),

  hasUnsavedChanges: () => get().dirtyScopes.length > 0,

  registerSaveHandler: (scope, handler) => {
    set(state => ({ handlers: { ...state.handlers, [scope]: handler } }));
  },

  unregisterSaveHandler: (scope) => {
    set(state => {
      const nextHandlers = { ...state.handlers };
      delete nextHandlers[scope];
      return { handlers: nextHandlers };
    });
  },

  saveCurrentProject: async () => {
    const state = get();
    if (state.isSaving) return;

    const workspace = useWorkspaceStore.getState();
    if (!workspace.currentWorkspace || !workspace.workspaceId || !workspace.workspaceRoot) {
      const message = "请先打开工作区";
      set({ lastSaveError: message });
      throw new Error(message);
    }

    const currentProject = useProjectStore.getState().currentProject;
    if (!currentProject?.id) {
      const message = "请先选择或创建项目";
      set({ lastSaveError: message });
      throw new Error(message);
    }

    const dirtySnapshot = uniqueScopes([...state.dirtyScopes]);
    if (dirtySnapshot.length === 0) return;

    const context: SaveContext = {
      workspaceId: workspace.workspaceId,
      workspaceRoot: workspace.workspaceRoot,
      projectId: currentProject.id,
      dirtyScopes: dirtySnapshot,
    };

    set({ isSaving: true, lastSaveError: null });

    const savedScopes: DirtyScope[] = [];
    const errors: string[] = [];

    for (const scope of dirtySnapshot) {
      const latestWorkspace = useWorkspaceStore.getState();
      const latestProject = useProjectStore.getState().currentProject;
      if (latestWorkspace.workspaceId !== context.workspaceId || latestProject?.id !== context.projectId) {
        errors.push("保存过程中项目或工作区已切换，已停止应用保存结果");
        break;
      }

      const handler = get().handlers[scope];
      if (!handler) {
        errors.push(`${scope} 暂未接入统一保存`);
        continue;
      }

      try {
        const result = await handler(context);
        if (!result.success) {
          errors.push(`${scope}: ${result.error || "保存失败"}`);
          continue;
        }

        const afterWorkspace = useWorkspaceStore.getState();
        const afterProject = useProjectStore.getState().currentProject;
        if (afterWorkspace === latestWorkspace && afterWorkspace.workspaceId === context.workspaceId && afterProject?.id === context.projectId) {
          savedScopes.push(...normalizeSavedScopes(result.savedScopes, dirtySnapshot));
        } else if (afterWorkspace.workspaceId === context.workspaceId && afterProject?.id === context.projectId) {
          savedScopes.push(...normalizeSavedScopes(result.savedScopes, dirtySnapshot));
        } else {
          errors.push("保存完成时项目或工作区已切换，未清除 dirty 状态");
          break;
        }
      } catch (error) {
        errors.push(`${scope}: ${error instanceof Error ? error.message : String(error)}`);
      }
    }

    const stillSameContext =
      useWorkspaceStore.getState().workspaceId === context.workspaceId &&
      useProjectStore.getState().currentProject?.id === context.projectId;

    const uniqueSavedScopes = normalizeSavedScopes(savedScopes, dirtySnapshot);
    if (stillSameContext && uniqueSavedScopes.length > 0) {
      get().clearDirtyScopes(uniqueSavedScopes);
    }

    const remainingSnapshotScopes = dirtySnapshot.filter(scope => !uniqueSavedScopes.includes(scope));
    const failed = errors.length > 0 || remainingSnapshotScopes.length > 0;
    set({
      isSaving: false,
      lastSavedAt: failed ? get().lastSavedAt : new Date().toISOString(),
      lastSaveError: failed ? errors.join("; ") || "部分内容未保存" : null,
    });

    if (failed) {
      throw new Error(errors.join("; ") || "部分内容未保存");
    }
  },
}));
