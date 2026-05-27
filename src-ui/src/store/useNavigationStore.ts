import { create } from "zustand";

export type ViewType = "hub" | "benefit" | "docfill" | "project_board" | "ict_lifecycle" | "data_management";

interface NavigationState {
  currentView: ViewType;
  activeProjectId: string | null;
  activeSchemeId: string | null;
  entrySource: "hub" | "project_board" | null;
  navigateTo: (view: ViewType, projectId?: string | null, schemeId?: string | null) => void;
  clearContext: () => void;
}

const STORAGE_KEY = "lamber_navigation_state";

interface StoredNavigation {
  currentView: ViewType;
  activeProjectId: string | null;
  activeSchemeId: string | null;
  entrySource: "hub" | "project_board" | null;
}

function getInitialState(): StoredNavigation {
  if (typeof window === "undefined") {
    return { currentView: "hub", activeProjectId: null, activeSchemeId: null, entrySource: null };
  }
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      return {
        currentView: parsed.currentView || "hub",
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
  const initial = getInitialState();

  return {
    ...initial,
    navigateTo: (view, projectId = null, schemeId = null) => {
      set((state) => {
        let newEntrySource = state.entrySource;
        if (view === "ict_lifecycle") {
          // If we navigate to ict_lifecycle, and we were in hub or project_board, remember it
          if (state.currentView === "hub" || state.currentView === "project_board") {
            newEntrySource = state.currentView as "hub" | "project_board";
          }
        } else if (view === "hub" || view === "project_board") {
          newEntrySource = null;
        }

        const newState = {
          currentView: view,
          activeProjectId: projectId,
          activeSchemeId: schemeId,
          entrySource: newEntrySource,
        };

        try {
          localStorage.setItem(STORAGE_KEY, JSON.stringify(newState));
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
          localStorage.setItem(STORAGE_KEY, JSON.stringify({
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
