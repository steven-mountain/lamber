import { useCallback, useEffect } from "react";
import { useSaveStore } from "../store/useSaveStore";
import { useWorkspaceStore } from "../store/useWorkspaceStore";
import { useProjectStore } from "../store/useProjectStore";

/**
 * 是否具备"保存到当前项目"的前提：工作区已打开且已绑定项目。
 * 与 useSaveStore.saveCurrentProject 的前置校验保持一致，避免在自由测算模式
 * （未绑定项目）下引导用户去做一次注定失败的保存，从而把返回操作卡死。
 */
function canSaveToProject() {
  const workspace = useWorkspaceStore.getState();
  if (!workspace.currentWorkspace || !workspace.workspaceId || !workspace.workspaceRoot) {
    return false;
  }
  return Boolean(useProjectStore.getState().currentProject?.id);
}

export function useUnsavedChangesGuard() {
  const hasUnsavedChanges = useSaveStore(state => state.hasUnsavedChanges);
  const saveCurrentProject = useSaveStore(state => state.saveCurrentProject);
  const clearAllDirty = useSaveStore(state => state.clearAllDirty);

  const confirmOrSave = useCallback(async () => {
    if (!hasUnsavedChanges()) return true;

    // 放弃修改并离开 / 停留当前页面。任何路径都保证用户能退出，绝不死锁。
    const promptDiscardOrStay = (message: string) => {
      const shouldDiscard = window.confirm(message);
      if (shouldDiscard) {
        clearAllDirty();
        return true;
      }
      return false;
    };

    // 自由测算模式（未绑定项目）：改动无法保存到项目，只提供放弃/停留，不再引导失败的保存。
    if (!canSaveToProject()) {
      return promptDiscardOrStay(
        "当前为自由测算模式（未绑定项目），改动无法保存到项目。\n\n确定：放弃改动并返回\n取消：停留当前页面",
      );
    }

    const shouldSave = window.confirm(
      "当前项目存在未保存修改，是否保存后继续？\n\n确定：保存并继续\n取消：选择放弃修改或停留当前页面",
    );
    if (shouldSave) {
      try {
        await saveCurrentProject();
        return true;
      } catch (error) {
        // 保存失败也不能把用户卡死：提示后仍给出放弃/停留的选择。
        alert(`保存失败：${error instanceof Error ? error.message : String(error)}`);
        return promptDiscardOrStay(
          "保存未成功。是否放弃未保存修改并返回？\n\n确定：放弃修改并返回\n取消：停留当前页面",
        );
      }
    }

    return promptDiscardOrStay(
      "是否放弃未保存修改并继续？\n\n确定：放弃修改并返回\n取消：停留当前页面",
    );
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
