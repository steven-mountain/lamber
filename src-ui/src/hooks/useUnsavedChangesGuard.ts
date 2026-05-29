import { useCallback, useEffect } from "react";
import { useSaveStore } from "../store/useSaveStore";

export function useUnsavedChangesGuard() {
  const hasUnsavedChanges = useSaveStore(state => state.hasUnsavedChanges);
  const saveCurrentProject = useSaveStore(state => state.saveCurrentProject);
  const clearAllDirty = useSaveStore(state => state.clearAllDirty);

  const confirmOrSave = useCallback(async () => {
    if (!hasUnsavedChanges()) return true;

    const shouldSave = window.confirm(
      "当前项目存在未保存修改，是否保存后继续？\n\n确定：保存并继续\n取消：选择放弃修改或停留当前页面",
    );
    if (shouldSave) {
      try {
        await saveCurrentProject();
        return true;
      } catch (error) {
        alert(`保存失败，已停留在当前页面：${error instanceof Error ? error.message : String(error)}`);
        return false;
      }
    }

    const shouldDiscard = window.confirm(
      "是否放弃未保存修改并继续？\n\n确定：放弃修改\n取消：停留当前页面",
    );
    if (shouldDiscard) {
      clearAllDirty();
      return true;
    }

    return false;
  }, [clearAllDirty, hasUnsavedChanges, saveCurrentProject]);

  useEffect(() => {
    const handleBeforeUnload = (event: BeforeUnloadEvent) => {
      if (!useSaveStore.getState().hasUnsavedChanges()) return;
      event.preventDefault();
      event.returnValue = "";
    };
    window.addEventListener("beforeunload", handleBeforeUnload);
    return () => window.removeEventListener("beforeunload", handleBeforeUnload);
  }, []);

  return { confirmOrSave };
}
