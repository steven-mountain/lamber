import { create } from "zustand";

export type ViewType = "hub" | "project_board" | "ict_lifecycle" | "ai_compute_quote" | "data_management" | "preset_center" | "settings";

export interface IctOrigin {
  type: "intelligent_compute";
  workspaceId: string;
  projectId: string;
  projectName: string;
  amountSourceId: string | null;
}

interface NavigationState {
  currentView: ViewType;
  activeProjectId: string | null;
  activeSchemeId: string | null;
  entrySource: "hub" | "project_board" | "ai_compute_quote" | null;
  activeScenarioId: string | null;
  ictOrigin: IctOrigin | null;
  settingsReturnView: ViewType | null;
  navigateTo: (
    view: ViewType,
    projectId?: string | null,
    schemeId?: string | null,
    scenarioId?: string | null,
  ) => void;
  openIctFromIntelligentCompute: (origin: IctOrigin) => void;
  clearContext: () => void;
}

export const NAVIGATION_STORAGE_KEY = "lamber_navigation_state";

export interface StoredNavigation {
  currentView: ViewType;
  activeProjectId: string | null;
  activeSchemeId: string | null;
  entrySource: "hub" | "project_board" | "ai_compute_quote" | null;
  activeScenarioId: string | null;
  ictOrigin: IctOrigin | null;
}

export function readStoredNavigationState(): StoredNavigation {
  if (typeof window === "undefined") {
    return { currentView: "hub", activeProjectId: null, activeSchemeId: null, entrySource: null, activeScenarioId: null, ictOrigin: null };
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
        activeScenarioId: parsed.activeScenarioId || null,
        ictOrigin: parsed.ictOrigin || null,
      };
    }
  } catch (e) {
    console.warn("Failed to load navigation state from localStorage:", e);
  }
  return { currentView: "hub", activeProjectId: null, activeSchemeId: null, entrySource: null, activeScenarioId: null, ictOrigin: null };
}

export const useNavigationStore = create<NavigationState>((set) => {
  const initial = readStoredNavigationState();

  return {
    ...initial,
    settingsReturnView: null,
    navigateTo: (view, projectId = null, schemeId = null, scenarioId = null) => {
      set((state) => {
        let newEntrySource = state.entrySource;
        let newSettingsReturnView = state.settingsReturnView;
        let newScenarioId = state.activeScenarioId;
        let newIctOrigin = state.ictOrigin;

        if (view === "ict_lifecycle") {
          if (state.currentView === "hub" || state.currentView === "project_board" || state.currentView === "ai_compute_quote") {
            newEntrySource = state.currentView as "hub" | "project_board" | "ai_compute_quote";
          }
          if (state.currentView === "ai_compute_quote") newScenarioId = scenarioId || state.activeScenarioId;
          if (state.currentView !== "ict_lifecycle" || (newIctOrigin && newIctOrigin.projectId !== projectId)) {
            newIctOrigin = null;
          }
        } else if (view === "ai_compute_quote") {
          newEntrySource = null;
          newScenarioId = scenarioId || state.activeScenarioId;
        } else if (view === "hub" || view === "project_board") {
          newEntrySource = null;
          newScenarioId = null;
          newIctOrigin = null;
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
          activeScenarioId: newScenarioId,
          ictOrigin: newIctOrigin,
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
    openIctFromIntelligentCompute: origin => {
      set(state => {
        const newState = {
          ...state,
          currentView: "ict_lifecycle" as const,
          activeProjectId: origin.projectId,
          activeSchemeId: null,
          entrySource: "ai_compute_quote" as const,
          activeScenarioId: origin.amountSourceId,
          ictOrigin: origin,
          settingsReturnView: null,
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
          activeScenarioId: null,
          ictOrigin: null,
        };
        try {
          localStorage.setItem(NAVIGATION_STORAGE_KEY, JSON.stringify({
            currentView: state.currentView,
            activeProjectId: null,
            activeSchemeId: null,
            entrySource: state.entrySource,
            activeScenarioId: null,
            ictOrigin: null,
          }));
        } catch (e) {
          console.warn("Failed to clear navigation context in localStorage:", e);
        }
        return newState;
      });
    },
  };
});
