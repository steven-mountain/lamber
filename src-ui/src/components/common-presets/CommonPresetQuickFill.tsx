import { type ReactNode, useEffect, useMemo, useRef, useState } from "react";
import AppIcon from "../icons/AppIcon";
import { Button } from "../ui/button";
import { Card } from "../ui/card";
import { Input } from "../ui/input";
import {
  getPresetFieldDisplay,
  type CommonPresetKind,
  type PresetFieldDefinition,
} from "../../lib/presetFieldKeys";
import {
  commonPresetService,
  type CommonPreset,
  type PresetFieldSetting,
} from "../../services/commonPresetService";
import { useWorkspaceStore } from "../../store/useWorkspaceStore";

interface CommonPresetQuickFillProps {
  fieldKey: string;
  kind: CommonPresetKind;
  value: string;
  onApply: (value: string) => void;
  className?: string;
}

interface CommonPresetFieldHeaderProps extends Omit<CommonPresetQuickFillProps, "className"> {
  children: ReactNode;
  className?: string;
  labelClassName?: string;
  actionsClassName?: string;
}

interface CommonPresetLabelHeaderProps {
  children: ReactNode;
  className?: string;
  labelClassName?: string;
}

type PresetPanelView = "select" | "save" | "edit";

const fieldSettingsCache = new Map<string, Map<string, PresetFieldSetting>>();
const fieldSettingsRequests = new Map<string, Promise<Map<string, PresetFieldSetting>>>();
const fieldSettingListeners = new Set<(workspaceId: string, setting: PresetFieldSetting) => void>();

function splitTags(raw: string): string[] {
  return raw
    .split(/[,\uFF0C\s]+/)
    .map(item => item.trim())
    .filter(Boolean);
}

function appendContent(current: string, incoming: string, kind: CommonPresetKind) {
  if (!current.trim()) return incoming;
  const separator = kind === "text_snippet" ? "\n" : " ";
  return `${current.trimEnd()}${separator}${incoming.trimStart()}`;
}

function buildDefaultPresetName(content: string, fallback: string) {
  const summary = content.trim().replace(/\s+/g, " ");
  if (!summary) return fallback;
  return summary.length > 24 ? `${summary.slice(0, 24)}...` : summary;
}

async function loadWorkspaceFieldSettings(workspaceId: string) {
  const cached = fieldSettingsCache.get(workspaceId);
  if (cached) return cached;
  const pending = fieldSettingsRequests.get(workspaceId);
  if (pending) return pending;

  const request = commonPresetService.listFieldSettings().then(settings => {
    const byKey = new Map(settings.map(setting => [setting.fieldKey, setting]));
    fieldSettingsCache.set(workspaceId, byKey);
    fieldSettingsRequests.delete(workspaceId);
    return byKey;
  }).catch(error => {
    fieldSettingsRequests.delete(workspaceId);
    throw error;
  });
  fieldSettingsRequests.set(workspaceId, request);
  return request;
}

function FieldBusinessMeta({
  field,
  compact = false,
}: {
  field: PresetFieldDefinition;
  compact?: boolean;
}) {
  return (
    <div className={compact ? "text-[11px] leading-4 text-muted-foreground" : "space-y-1 text-xs text-muted-foreground"}>
      <div><span className="font-medium text-secondary-foreground">适用模板：</span>{field.templates.join("、")}</div>
      <div><span className="font-medium text-secondary-foreground">所属分组：</span>{field.groups.join("、")}</div>
    </div>
  );
}

export default function CommonPresetQuickFill({
  fieldKey,
  kind,
  value,
  onApply,
  className = "",
}: CommonPresetQuickFillProps) {
  const rootRef = useRef<HTMLDivElement>(null);
  const workspaceId = useWorkspaceStore(state => state.workspaceId);
  const field = useMemo(() => getPresetFieldDisplay(fieldKey), [fieldKey]);
  const [fieldEnabled, setFieldEnabled] = useState(Boolean(field.defaultEnabled));
  const [fieldSettingLoading, setFieldSettingLoading] = useState(field.presetEligible);
  const [panelView, setPanelView] = useState<PresetPanelView | null>(null);
  const [presets, setPresets] = useState<CommonPreset[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [saveError, setSaveError] = useState("");
  const [saveName, setSaveName] = useState("");
  const [saveCategory, setSaveCategory] = useState(field.category);
  const [saveTags, setSaveTags] = useState("");
  const [editingPreset, setEditingPreset] = useState<CommonPreset | null>(null);
  const [editName, setEditName] = useState("");
  const [editContent, setEditContent] = useState("");
  const [editCategory, setEditCategory] = useState("");
  const [editTags, setEditTags] = useState("");
  const [editError, setEditError] = useState("");

  useEffect(() => {
    setSaveCategory(field.category);
  }, [field.category]);

  useEffect(() => {
    let cancelled = false;
    setPanelView(null);
    setError("");
    if (!field.presetEligible || !workspaceId) {
      setFieldEnabled(false);
      setFieldSettingLoading(false);
      return () => {
        cancelled = true;
      };
    }

    setFieldSettingLoading(true);
    void loadWorkspaceFieldSettings(workspaceId)
      .then(settings => {
        if (cancelled) return;
        setFieldEnabled(settings.get(field.fieldKey)?.enabled ?? Boolean(field.defaultEnabled));
      })
      .catch(err => {
        console.warn("Failed to load preset field settings", err);
        if (!cancelled) {
          setFieldEnabled(Boolean(field.defaultEnabled));
        }
      })
      .finally(() => {
        if (!cancelled) setFieldSettingLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [field.defaultEnabled, field.fieldKey, field.presetEligible, workspaceId]);

  useEffect(() => {
    if (!workspaceId) return;
    const listener = (changedWorkspaceId: string, setting: PresetFieldSetting) => {
      if (changedWorkspaceId === workspaceId && setting.fieldKey === field.fieldKey) {
        setFieldEnabled(setting.enabled);
        if (!setting.enabled) {
          setPanelView(null);
        }
      }
    };
    fieldSettingListeners.add(listener);
    return () => {
      fieldSettingListeners.delete(listener);
    };
  }, [field.fieldKey, workspaceId]);

  useEffect(() => {
    if (!panelView) return;

    const handleOutsidePointerDown = (event: PointerEvent) => {
      if (rootRef.current?.contains(event.target as Node)) return;
      setPanelView(null);
    };

    document.addEventListener("pointerdown", handleOutsidePointerDown);
    return () => {
      document.removeEventListener("pointerdown", handleOutsidePointerDown);
    };
  }, [panelView]);

  useEffect(() => {
    if (!panelView) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setPanelView(null);
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, [panelView]);

  if (!field.presetEligible) return null;

  const loadPresets = async () => {
    setLoading(true);
    setError("");
    try {
      const items = await commonPresetService.list({
        kind,
        fieldKey: field.fieldKey,
        includeDisabled: false,
        sortBy: "recent",
      });
      setPresets(items);
    } catch (err) {
      console.error("Failed to load common presets", err);
      setError("无法读取常用资料，请确认已打开工作区。");
    } finally {
      setLoading(false);
    }
  };

  const updateFieldEnabled = async (enabled: boolean) => {
    if (!workspaceId) {
      setError("请先打开工作区。");
      return;
    }
    setFieldSettingLoading(true);
    setError("");
    try {
      const setting = await commonPresetService.setFieldEnabled(field.fieldKey, enabled);
      const workspaceSettings = fieldSettingsCache.get(workspaceId) ?? new Map();
      workspaceSettings.set(field.fieldKey, setting);
      fieldSettingsCache.set(workspaceId, workspaceSettings);
      setFieldEnabled(setting.enabled);
      fieldSettingListeners.forEach(listener => listener(workspaceId, setting));
      if (!setting.enabled) {
        setPanelView(null);
      }
    } catch (err) {
      console.error("Failed to update preset field setting", err);
      setError("字段预设状态更新失败。");
    } finally {
      setFieldSettingLoading(false);
    }
  };

  const handleToggle = async () => {
    if (panelView === "select") {
      setPanelView(null);
      return;
    }
    setEditingPreset(null);
    setPanelView("select");
    await loadPresets();
  };

  const openSaveView = () => {
    setEditingPreset(null);
    setError("");
    setSaveError("");
    setSaveName(buildDefaultPresetName(value, field.label));
    setSaveCategory(field.category || field.label);
    setSaveTags("");
    setPanelView("save");
    void loadPresets();
  };

  const returnToSelect = () => {
    setEditingPreset(null);
    setPanelView("select");
  };

  const openEditView = (preset: CommonPreset) => {
    setError("");
    setEditError("");
    setEditingPreset(preset);
    setEditName(preset.name);
    setEditContent(preset.content);
    setEditCategory(preset.category);
    setEditTags(preset.tags.join(" "));
    setPanelView("edit");
  };

  const closeFieldPreset = async () => {
    const confirmed = window.confirm(
      `关闭“${field.label}”的字段预设吗？\n\n当前字段内容和常用资料库中的资料都会保留，仅隐藏该字段的预设操作。`,
    );
    if (!confirmed) return;
    await updateFieldEnabled(false);
  };

  const applyPreset = async (preset: CommonPreset, mode: "replace" | "append") => {
    if (mode === "replace" && value.trim() && value !== preset.content) {
      const ok = window.confirm("当前字段已有内容，确定用所选常用内容替换吗？");
      if (!ok) return;
    }
    const nextValue = mode === "append" ? appendContent(value, preset.content, kind) : preset.content;
    onApply(nextValue);
    try {
      const updated = await commonPresetService.markUsed(preset.id);
      setPresets(current => current.map(item => item.id === updated.id ? updated : item));
    } catch (err) {
      console.warn("Failed to update preset usage", err);
    }
    setPanelView(null);
  };

  const deletePreset = async (preset: CommonPreset) => {
    const confirmed = window.confirm(
      `确定删除常用内容“${preset.name}”吗？\n\n删除只影响常用资料库，当前字段内容不会被清空。`,
    );
    if (!confirmed) return;
    setError("");
    try {
      await commonPresetService.delete(preset.id);
      await loadPresets();
    } catch (err) {
      console.error("Failed to delete common preset", err);
      setError("删除常用内容失败。");
    }
  };

  const saveCurrent = async () => {
    const content = value.trim();
    if (!content) {
      setSaveError("当前字段为空，不能保存为常用内容。");
      return;
    }
    const name = saveName.trim() || field.label;
    const category = saveCategory.trim() || field.category;
    setSaveError("");
    try {
      await commonPresetService.save({
        scope: "workspace",
        kind,
        category,
        name,
        content,
        tags: splitTags(saveTags),
        applicableFieldKeys: [field.fieldKey],
        enabled: true,
      });
      setSaveName("");
      setSaveTags("");
      await loadPresets();
      setPanelView("select");
    } catch (err) {
      console.error("Failed to save common preset", err);
      setSaveError("保存常用内容失败，请确认字段允许使用预设且工作区已打开。");
    }
  };

  const saveEditedPreset = async () => {
    if (!editingPreset) return;
    const name = editName.trim();
    const content = editContent.trim();
    const category = editCategory.trim();
    if (!name || !content || !category) {
      setEditError("请填写预设名称、预设内容和分类。");
      return;
    }
    setEditError("");
    try {
      await commonPresetService.save({
        id: editingPreset.id,
        scope: editingPreset.scope,
        kind: editingPreset.kind,
        category,
        name,
        content,
        tags: splitTags(editTags),
        applicableFieldKeys: editingPreset.applicableFieldKeys,
        enabled: editingPreset.enabled,
      });
      await loadPresets();
      setEditingPreset(null);
      setPanelView("select");
    } catch (err) {
      console.error("Failed to update common preset", err);
      setEditError("更新常用内容失败，请检查名称、分类和内容。");
    }
  };

  if (!fieldEnabled) {
    return (
      <div ref={rootRef} className={`relative inline-flex items-center ${className}`}>
        <Button
          type="button"
          variant="ghost"
          size="sm"
          className="h-6 px-2 text-xs text-muted-foreground"
          disabled={fieldSettingLoading}
          onClick={() => void updateFieldEnabled(true)}
          title={`为“${field.label}”启用字段预设`}
        >
          + 预设
        </Button>
        {error ? (
          <div className="absolute right-0 top-8 z-40 w-64 rounded-lg bg-destructive-soft px-3 py-2 text-xs text-destructive shadow-md">
            {error}
          </div>
        ) : null}
      </div>
    );
  }

  return (
    <div ref={rootRef} className={`relative inline-flex items-center gap-1 ${className}`}>
      <Button
        type="button"
        variant="ghost"
        size="sm"
        className="h-6 px-2 text-xs text-secondary-foreground hover:text-foreground"
        onClick={handleToggle}
      >
        <AppIcon name="presetLibrary" size={13} />
        选择常用
      </Button>
      <Button
        type="button"
        variant="ghost"
        size="sm"
        className="h-6 px-2 text-xs"
        onClick={openSaveView}
      >
        <AppIcon name="save" size={13} />
        保存当前
      </Button>
      <Button
        type="button"
        variant="ghost"
        size="icon"
        className="h-6 w-6 text-muted-foreground"
        aria-label={`关闭“${field.label}”字段预设`}
        title="关闭预设"
        disabled={fieldSettingLoading}
        onClick={() => void closeFieldPreset()}
      >
        <AppIcon name="close" size={14} />
      </Button>

      {panelView && (
        <div className="absolute right-0 top-10 z-40 w-[min(32rem,calc(100vw-2rem))] rounded-xl bg-popover p-3 text-popover-foreground shadow-lg ring-1 ring-border/60">
          {panelView === "select" ? (
            <>
              <div className="mb-2 flex items-start justify-between gap-3 px-1 pb-2">
                <div className="min-w-0">
                  <div className="text-sm font-semibold text-foreground">{field.label}</div>
                  <div className="mt-1">
                    <FieldBusinessMeta field={field} compact />
                  </div>
                </div>
                <Button
                  type="button"
                  variant="ghost"
                  size="icon"
                  className="h-7 w-7 shrink-0"
                  onClick={() => setPanelView(null)}
                  aria-label="关闭常用内容面板"
                >
                  <AppIcon name="close" size={14} />
                </Button>
              </div>

              <div className="mb-2 flex items-center justify-between px-1">
                <div className="text-xs font-semibold text-secondary-foreground">选择已有常用内容</div>
                {!loading && presets.length > 0 ? (
                  <div className="text-[11px] tabular-nums text-muted-foreground">{presets.length} 条</div>
                ) : null}
              </div>

              <div className="max-h-[min(22rem,calc(100vh-18rem))] space-y-1.5 overflow-y-auto pr-1">
                {loading ? (
                  <div className="flex min-h-36 items-center justify-center rounded-lg bg-muted/35 px-3 py-4 text-sm text-secondary-foreground">
                    正在读取常用内容...
                  </div>
                ) : presets.length === 0 ? (
                  <div className="flex min-h-36 flex-col items-center justify-center rounded-lg bg-muted/30 px-4 py-6 text-center">
                    <div className="text-sm text-secondary-foreground">暂无可用于该字段的常用内容</div>
                    <Button
                      type="button"
                      size="sm"
                      className="mt-3"
                      disabled={!value.trim()}
                      onClick={openSaveView}
                    >
                      保存当前内容为第一条常用内容
                    </Button>
                    {!value.trim() ? (
                      <div className="mt-2 text-[11px] text-muted-foreground">当前字段为空，填写内容后即可保存</div>
                    ) : null}
                  </div>
                ) : (
                  presets.map(preset => (
                    <Card
                      key={preset.id}
                      role="button"
                      tabIndex={0}
                      className="cursor-pointer border-0 bg-muted/30 px-2.5 py-2 shadow-none transition-colors hover:bg-muted/55 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/20"
                      onClick={() => void applyPreset(preset, "replace")}
                      onKeyDown={event => {
                        if (event.target !== event.currentTarget) return;
                        if (event.key === "Enter" || event.key === " ") {
                          event.preventDefault();
                          void applyPreset(preset, "replace");
                        }
                      }}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <div className="min-w-0 flex-1">
                          <div className="flex min-w-0 items-center gap-2">
                            <div className="truncate text-sm font-semibold text-foreground">{preset.name}</div>
                            <div className="shrink-0 truncate text-[11px] text-muted-foreground">
                              {preset.category} · {field.label}
                            </div>
                          </div>
                          <div className={`mt-0.5 whitespace-pre-wrap text-xs leading-4 text-secondary-foreground ${
                            kind === "text_snippet" ? "line-clamp-2" : "truncate"
                          }`}>
                            {preset.content}
                          </div>
                          <div className="mt-0.5 truncate text-[11px] tabular-nums text-muted-foreground">
                            使用 {preset.usageCount} 次
                            {preset.lastUsedAt ? ` · 最近使用 ${new Date(preset.lastUsedAt).toLocaleString()}` : ""}
                          </div>
                        </div>
                        <div className="flex shrink-0 items-center gap-0.5">
                          <Button
                            type="button"
                            size="sm"
                            className="h-7 px-2.5 shadow-none"
                            onClick={event => {
                              event.stopPropagation();
                              void applyPreset(preset, "replace");
                            }}
                          >
                            替换
                          </Button>
                          <Button
                            type="button"
                            variant="outline"
                            size="sm"
                            className="h-7 px-2.5 shadow-none"
                            onClick={event => {
                              event.stopPropagation();
                              void applyPreset(preset, "append");
                            }}
                          >
                            追加
                          </Button>
                          <Button
                            type="button"
                            variant="outline"
                            size="sm"
                            className="h-7 px-2 shadow-none"
                            onClick={event => {
                              event.stopPropagation();
                              openEditView(preset);
                            }}
                          >
                            <AppIcon name="edit" size={13} />
                            编辑
                          </Button>
                          <Button
                            type="button"
                            variant="ghost"
                            size="sm"
                            className="h-7 bg-destructive-soft px-2 text-destructive hover:bg-destructive-soft/80 hover:text-destructive"
                            onClick={event => {
                              event.stopPropagation();
                              void deletePreset(preset);
                            }}
                          >
                            <AppIcon name="delete" size={13} />
                            删除
                          </Button>
                        </div>
                      </div>
                    </Card>
                  ))
                )}
              </div>

              <div className="mt-3 flex items-center justify-between rounded-lg bg-muted/20 px-3 py-2">
                <span className="text-xs text-muted-foreground">没有合适的常用内容？</span>
                <Button type="button" variant="ghost" size="sm" className="h-7 px-2 text-xs" onClick={openSaveView}>
                  保存当前内容为常用
                </Button>
              </div>

              {error ? <div className="mt-2 rounded-lg bg-destructive-soft px-3 py-2 text-xs text-destructive">{error}</div> : null}

              <div className="mt-2 flex items-center justify-between px-1 pt-1">
                <span className="text-xs text-muted-foreground">该设置保存在当前工作区</span>
                <Button
                  type="button"
                  variant="ghost"
                  size="sm"
                  className="h-7 text-xs text-muted-foreground hover:text-destructive"
                  disabled={fieldSettingLoading}
                  onClick={() => void closeFieldPreset()}
                >
                  关闭字段预设
                </Button>
              </div>
            </>
          ) : panelView === "save" ? (
            <>
              <div className="flex items-center justify-between gap-2 px-1 pb-2">
                <Button type="button" variant="ghost" size="sm" className="h-7 px-2" onClick={returnToSelect}>
                  返回
                </Button>
                <div className="text-sm font-semibold text-foreground">保存当前内容为常用</div>
                <Button
                  type="button"
                  variant="ghost"
                  size="icon"
                  className="h-7 w-7"
                  onClick={() => setPanelView(null)}
                  aria-label="关闭字段预设面板"
                >
                  <AppIcon name="close" size={14} />
                </Button>
              </div>

              <div className="max-h-[min(30rem,calc(100vh-10rem))] overflow-y-auto px-1 pr-2">
                <div className="rounded-lg bg-muted/30 px-3 py-2">
                  <div className="text-sm font-semibold text-foreground">{field.label}</div>
                  <div className="mt-1 text-[11px] leading-4 text-muted-foreground">
                    <div><span className="font-medium text-secondary-foreground">所属模板：</span>{field.templates.join("、")}</div>
                    <div><span className="font-medium text-secondary-foreground">所属分组：</span>{field.groups.join("、")}</div>
                  </div>
                </div>

                <div className="mt-3">
                  <div className="mb-1 text-xs font-medium text-secondary-foreground">将保存的内容</div>
                  <div className="max-h-24 overflow-y-auto rounded-md bg-muted/35 px-3 py-2 text-sm leading-5 text-foreground whitespace-pre-wrap">
                    {value.trim() || "当前字段为空"}
                  </div>
                </div>
                <div className="mt-3 grid gap-2.5">
                  <label className="grid gap-1">
                    <span className="text-xs font-medium text-secondary-foreground">常用名称</span>
                    <Input value={saveName} onChange={event => setSaveName(event.target.value)} placeholder={field.label} autoFocus />
                  </label>
                  <label className="grid gap-1">
                    <span className="text-xs font-medium text-secondary-foreground">分类</span>
                    <Input value={saveCategory} onChange={event => setSaveCategory(event.target.value)} placeholder={field.category || field.label} />
                  </label>
                  <label className="grid gap-1">
                    <span className="text-xs font-medium text-secondary-foreground">标签（可选）</span>
                    <Input value={saveTags} onChange={event => setSaveTags(event.target.value)} placeholder="多个标签可用空格或逗号分隔" />
                  </label>
                </div>
              </div>

              {saveError ? (
                <div className="mt-2 rounded-lg bg-destructive-soft px-3 py-2 text-xs text-destructive">{saveError}</div>
              ) : null}

              <div className="mt-3 flex justify-end gap-2">
                <Button type="button" variant="ghost" size="sm" onClick={returnToSelect}>
                  取消
                </Button>
                <Button type="button" size="sm" disabled={!value.trim()} onClick={() => void saveCurrent()}>
                  保存为常用内容
                </Button>
              </div>
            </>
          ) : panelView === "edit" && editingPreset ? (
            <>
              <div className="flex items-center justify-between gap-2 px-1 pb-2">
                <Button type="button" variant="ghost" size="sm" className="h-7 px-2" onClick={returnToSelect}>
                  返回
                </Button>
                <div className="text-sm font-semibold text-foreground">编辑常用内容</div>
                <Button
                  type="button"
                  variant="ghost"
                  size="icon"
                  className="h-7 w-7"
                  onClick={() => setPanelView(null)}
                  aria-label="关闭字段预设面板"
                >
                  <AppIcon name="close" size={14} />
                </Button>
              </div>

              <div className="max-h-[min(30rem,calc(100vh-10rem))] overflow-y-auto px-1 pr-2">
                <div className="mb-3 rounded-lg bg-muted/25 px-3 py-2 text-xs text-muted-foreground">
                  修改常用资料本身，不会改写当前项目字段。
                </div>
                <div className="grid gap-2.5">
                  <label className="grid gap-1">
                    <span className="text-xs font-medium text-secondary-foreground">常用名称</span>
                    <Input value={editName} onChange={event => setEditName(event.target.value)} autoFocus />
                  </label>
                  <label className="grid gap-1">
                    <span className="text-xs font-medium text-secondary-foreground">预设内容</span>
                    <textarea
                      value={editContent}
                      onChange={event => setEditContent(event.target.value)}
                      rows={editingPreset.kind === "text_snippet" ? 6 : 3}
                      className="min-h-20 w-full resize-y rounded-md border border-input bg-card px-3 py-2 text-sm text-foreground shadow-sm transition-colors placeholder:text-muted-foreground focus-visible:border-ring focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/20"
                    />
                  </label>
                  <label className="grid gap-1">
                    <span className="text-xs font-medium text-secondary-foreground">分类</span>
                    <Input value={editCategory} onChange={event => setEditCategory(event.target.value)} />
                  </label>
                  <label className="grid gap-1">
                    <span className="text-xs font-medium text-secondary-foreground">标签（可选）</span>
                    <Input value={editTags} onChange={event => setEditTags(event.target.value)} placeholder="多个标签可用空格或逗号分隔" />
                  </label>
                </div>
              </div>

              {editError ? (
                <div className="mt-2 rounded-lg bg-destructive-soft px-3 py-2 text-xs text-destructive">{editError}</div>
              ) : null}

              <div className="mt-3 flex justify-end gap-2">
                <Button type="button" variant="ghost" size="sm" onClick={returnToSelect}>
                  取消
                </Button>
                <Button type="button" size="sm" onClick={() => void saveEditedPreset()}>
                  保存修改
                </Button>
              </div>
            </>
          ) : null}
        </div>
      )}
    </div>
  );
}

export function CommonPresetFieldHeader({
  children,
  fieldKey,
  kind,
  value,
  onApply,
  className = "",
  labelClassName = "text-sm font-semibold text-foreground",
  actionsClassName = "",
}: CommonPresetFieldHeaderProps) {
  return (
    <div className={`flex min-h-6 flex-wrap items-center justify-between gap-x-3 gap-y-1 ${className}`}>
      <div className={`min-w-0 ${labelClassName}`}>{children}</div>
      <CommonPresetQuickFill
        fieldKey={fieldKey}
        kind={kind}
        value={value}
        onApply={onApply}
        className={`shrink-0 ${actionsClassName}`}
      />
    </div>
  );
}

export function CommonPresetLabelHeader({
  children,
  className = "",
  labelClassName = "text-sm font-semibold text-foreground",
}: CommonPresetLabelHeaderProps) {
  return (
    <div className={`flex min-h-6 items-center ${className}`}>
      <div className={`min-w-0 ${labelClassName}`}>{children}</div>
    </div>
  );
}
