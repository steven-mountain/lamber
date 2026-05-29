import AppIcon from "./icons/AppIcon";
import { useSaveStore } from "../store/useSaveStore";

export default function GlobalSaveButton() {
  const isDirty = useSaveStore(state => state.isDirty);
  const isSaving = useSaveStore(state => state.isSaving);
  const lastSaveError = useSaveStore(state => state.lastSaveError);
  const saveCurrentProject = useSaveStore(state => state.saveCurrentProject);

  const label = isSaving
    ? "保存中..."
    : lastSaveError
      ? "保存失败，请重试"
      : isDirty
        ? "保存"
        : "已保存";

  const handleClick = async () => {
    if (isSaving || !isDirty) return;
    try {
      await saveCurrentProject();
    } catch (error) {
      alert(error instanceof Error ? error.message : String(error));
    }
  };

  return (
    <button
      type="button"
      disabled={isSaving || !isDirty}
      onClick={handleClick}
      title={lastSaveError || label}
      className={`inline-flex items-center gap-1.5 rounded-md px-3 py-2 text-xs font-bold shadow-sm transition-colors disabled:cursor-default disabled:opacity-70 ${
        lastSaveError
          ? "bg-amber-50 text-amber-700 hover:bg-amber-100"
          : isDirty
            ? "bg-primary text-primary-foreground hover:bg-primary/90"
            : "bg-secondary text-secondary-foreground"
      }`}
    >
      <AppIcon name={isSaving ? "loading" : "save"} size={14} className={isSaving ? "animate-spin" : ""} />
      {label}
    </button>
  );
}
