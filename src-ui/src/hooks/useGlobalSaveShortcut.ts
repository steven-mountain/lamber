import { useEffect } from "react";
import { useSaveStore } from "../store/useSaveStore";
import { useProjectStore } from "../store/useProjectStore";
import { useWorkspaceStore } from "../store/useWorkspaceStore";

export function useGlobalSaveShortcut() {
  useEffect(() => {
    const handleKeyDown = async (event: KeyboardEvent) => {
      const isSaveShortcut = (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s";
      if (!isSaveShortcut) return;

      event.preventDefault();
      const saveState = useSaveStore.getState();
      if (saveState.isSaving) return;

      const workspace = useWorkspaceStore.getState();
      if (!workspace.currentWorkspace || !workspace.workspaceId || !workspace.workspaceRoot) {
        alert("请先打开工作区");
        return;
      }

      const currentProject = useProjectStore.getState().currentProject;
      if (!currentProject?.id) {
        alert("请先选择或创建项目");
        return;
      }

      if (!saveState.hasUnsavedChanges()) return;

      try {
        await saveState.saveCurrentProject();
      } catch (error) {
        alert(error instanceof Error ? error.message : String(error));
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, []);
}
