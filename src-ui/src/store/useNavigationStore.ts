import { create } from "zustand";

export type ViewType = "hub" | "project_board" | "ict_lifecycle" | "data_management" | "preset_center" | "settings";

interface NavigationState {
  currentView: ViewType;
  activeProjectId: string | null;
  activeSchemeId: string | null;
  entrySource: "hub" | "project_board" | null;
  settingsReturnView: ViewType | null;
  navigateTo: (view: ViewType, projectId?: string | null, schemeId?: string | null) => void;
  clearContext: () => void;
}

export const NAVIGATION_STORAGE_KEY = "lamber_navigation_state";

export interface StoredNavigation {
  currentView: ViewType;
  activeProjectId: string | null;
  activeSchemeId: string | null;
  entrySource: "hub" | "project_board" | null;
}

export function readStoredNavigationState(): StoredNavigation {
  if (typeof window === "undefined") {
    return { currentView: "hub", activeProjectId: null, activeSchemeId: null, entrySource: null };
  }
  try {
    const raw = localStorage.getItem(NAVIGATION_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      return {
        currentView: "hub",
        activeProjectId: parsed.activeProjectId || null,
        activeSchemeId: parsed.activeSchemeId || null,
        entrySource: parsed.entrySource || null,
      };
    }
  } catch (e) {
    console.warn("Failed to load navigation state from localStorage:", e);
  }
  return { currentView: "hub", activeProjectId: null, activeSchemeId: null, entrySource: null };
}

export const useNavigationStore = create<NavigationState>((set) => {
  const initial = readStoredNavigationState();

  return {
    ...initial,
    settingsReturnView: null,
    navigateTo: (view, projectId = null, schemeId = null) => {
      set((state) => {
        let newEntrySource = state.entrySource;
        let newSettingsReturnView = state.settingsReturnView;

        if (view === "ict_lifecycle") {
          // If we navigate to ict_lifecycle, and we were in hub or project_board, remember it
          if (state.currentView === "hub" || state.currentView === "project_board") {
            newEntrySource = state.currentView as "hub" | "project_board";
          }
        } else if (view === "hub" || view === "project_board") {
          newEntrySource = null;
        }

        if (view === "settings") {
          newSettingsReturnView = state.currentView;
        } else if (state.currentView === "settings") {
          newSettingsReturnView = null;
        }

        const finalProjectId = view === "settings" ? (projectId || state.activeProjectId) : projectId;
        const finalSchemeId = view === "settings" ? (schemeId || state.activeSchemeId) : schemeId;

        const newState = {
          currentView: view,
          activeProjectId: finalProjectId,
          activeSchemeId: finalSchemeId,
          entrySource: newEntrySource,
          settingsReturnView: newSettingsReturnView,
        };

        try {
          const { settingsReturnView: _, ...storedState } = newState;
          localStorage.setItem(NAVIGATION_STORAGE_KEY, JSON.stringify(storedState));
        } catch (e) {
          console.warn("Failed to save navigation state to localStorage:", e);
        }

        return newState;
      });
    },
    clearContext: () => {
      set((state) => {
        const newState = {
          ...state,
          activeProjectId: null,
          activeSchemeId: null,
        };
        try {
          localStorage.setItem(NAVIGATION_STORAGE_KEY, JSON.stringify({
            currentView: state.currentView,
            activeProjectId: null,
            activeSchemeId: null,
            entrySource: state.entrySource,
          }));
        } catch (e) {
          console.warn("Failed to clear navigation context in localStorage:", e);
        }
        return newState;
      });
    },
  };
});
