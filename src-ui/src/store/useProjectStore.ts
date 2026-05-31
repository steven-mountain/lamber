import { create } from "zustand";
import type { Project } from "../utils/projectService";

export const PROJECT_STORAGE_KEY = "lamber_current_project";

export function readStoredCurrentProject(): Project | null {
  if (typeof window === "undefined") return null;

  try {
    const saved = window.localStorage.getItem(PROJECT_STORAGE_KEY);
    return saved ? JSON.parse(saved) : null;
  } catch (e) {
    console.error("Failed to restore current project:", e);
    return null;
  }
}

function clearStoredActiveProjectKeys() {
  if (typeof window === "undefined") return;
  window.localStorage.removeItem("lamber_active_project_id");
  window.localStorage.removeItem("lamber_active_scheme_id");
}

interface ProjectState {
  currentProject: Project | null;
  setCurrentProject: (project: Project | null) => void;
  clearCurrentProject: () => void;
}

export const useProjectStore = create<ProjectState>((set) => {
  const initialProject = readStoredCurrentProject();

  return {
    currentProject: initialProject,
    setCurrentProject: (project) => {
      set({ currentProject: project });
      if (project) {
        localStorage.setItem(PROJECT_STORAGE_KEY, JSON.stringify(project));
      } else {
        localStorage.removeItem(PROJECT_STORAGE_KEY);
        clearStoredActiveProjectKeys();
      }
    },
    clearCurrentProject: () => {
      set({ currentProject: null });
      localStorage.removeItem(PROJECT_STORAGE_KEY);
      clearStoredActiveProjectKeys();
    },
  };
});
