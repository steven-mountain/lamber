import { create } from "zustand";
import { Project } from "../utils/projectService";

interface ProjectState {
  currentProject: Project | null;
  setCurrentProject: (project: Project | null) => void;
  clearCurrentProject: () => void;
}

export const useProjectStore = create<ProjectState>((set) => {
  // Try to restore project from localStorage
  let initialProject: Project | null = null;
  try {
    const saved = localStorage.getItem("lamber_current_project");
    if (saved) {
      initialProject = JSON.parse(saved);
    }
  } catch (e) {
    console.error("Failed to restore current project:", e);
  }

  return {
    currentProject: initialProject,
    setCurrentProject: (project) => {
      set({ currentProject: project });
      if (project) {
        localStorage.setItem("lamber_current_project", JSON.stringify(project));
      } else {
        localStorage.removeItem("lamber_current_project");
      }
    },
    clearCurrentProject: () => {
      set({ currentProject: null });
      localStorage.removeItem("lamber_current_project");
    },
  };
});
