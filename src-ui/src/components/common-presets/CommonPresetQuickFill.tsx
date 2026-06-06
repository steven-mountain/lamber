import { type ReactNode, useEffect, useMemo, useState } from "react";
import AppIcon from "../icons/AppIcon";
import { Button } from "../ui/button";
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

const fieldSettingsCache = new Map<string, Map<string, PresetFieldSetting>>();
const fieldSettingsRequests = new Map<string, Promise<Map<string, PresetFieldSetting>>>();
const fieldSettingListeners = new Set<(workspaceId: string, setting: PresetFieldSetting) => void>();

function splitTags(raw: string): string[] {
  return raw
    .split(/[,\uFF0C\s]+/)
    .map(item => item.trim())
    .filter(Boolean);
}

function appendContent(current: string, incoming: string) {
  if (!current.trim()) return incoming;
  return `${current.trimEnd()}\n${incoming}`;
}

function listBusinessFields(fieldKeys: string[]) {
  if (fieldKeys.length === 0) return [];
  return fieldKeys.map(getPresetFieldDisplay);
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
  const workspaceId = useWorkspaceStore(state => state.workspaceId);
  const field = useMemo(() => getPresetFieldDisplay(fieldKey), [fieldKey]);
  const [fieldEnabled, setFieldEnabled] = useState(Boolean(field.defaultEnabled));
  const [fieldSettingLoading, setFieldSettingLoading] = useState(field.presetEligible);
  const [open, setOpen] = useState(false);
  const [presets, setPresets] = useState<CommonPreset[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [saveOpen, setSaveOpen] = useState(false);
  const [moreOpen, setMoreOpen] = useState(false);
  const [saveName, setSaveName] = useState("");
  const [saveCategory, setSaveCategory] = useState(field.category);
  const [saveTags, setSaveTags] = useState("");

  useEffect(() => {
    setSaveCategory(field.category);
  }, [field.category]);

  useEffect(() => {
    let cancelled = false;
    setOpen(false);
    setMoreOpen(false);
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
          setOpen(false);
          setMoreOpen(false);
        }
      }
    };
    fieldSettingListeners.add(listener);
    return () => {
      fieldSettingListeners.delete(listener);
    };
  }, [field.fieldKey, workspaceId]);

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
        setOpen(false);
        setMoreOpen(false);
      }
    } catch (err) {
      console.error("Failed to update preset field setting", err);
      setError("字段预设状态更新失败。");
    } finally {
      setFieldSettingLoading(false);
    }
  };

  const handleToggle = async () => {
    const nextOpen = !open;
    setOpen(nextOpen);
    setMoreOpen(false);
    setSaveOpen(false);
    if (nextOpen) {
      await loadPresets();
    }
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
    const nextValue = mode === "append" ? appendContent(value, preset.content) : preset.content;
    onApply(nextValue);
    try {
      const updated = await commonPresetService.markUsed(preset.id);
      setPresets(current => current.map(item => item.id === updated.id ? updated : item));
    } catch (err) {
      console.warn("Failed to update preset usage", err);
    }
    setOpen(false);
  };

  const saveCurrent = async () => {
    const content = value.trim();
    if (!content) {
      setError("当前字段为空，不能保存为常用内容。");
      return;
    }
    const name = saveName.trim() || field.label;
    const category = saveCategory.trim() || field.category;
    setError("");
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
      setSaveOpen(false);
      await loadPresets();
    } catch (err) {
      console.error("Failed to save common preset", err);
      setError("保存常用内容失败，请确认字段允许使用预设且工作区已打开。");
    }
  };

  if (!fieldEnabled) {
    return (
      <div className={`relative inline-flex items-center ${className}`}>
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
    <div className={`relative inline-flex items-center gap-1 ${className}`}>
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
        onClick={() => {
          setOpen(true);
          setSaveOpen(true);
          void loadPresets();
        }}
      >
        <AppIcon name="save" size={13} />
        保存当前
      </Button>
      <Button
        type="button"
        variant="ghost"
        size="icon"
        className="h-6 w-6 text-muted-foreground"
        aria-label={`管理“${field.label}”字段预设`}
        aria-expanded={moreOpen}
        onClick={() => {
          setMoreOpen(current => !current);
          setOpen(false);
        }}
      >
        <AppIcon name="more" size={14} />
      </Button>

      {moreOpen ? (
        <div className="absolute right-0 top-8 z-40 w-52 rounded-lg bg-popover p-1.5 text-popover-foreground shadow-lg ring-1 ring-border/60">
          <button
            type="button"
            className="flex w-full items-center gap-2 rounded-md px-3 py-2 text-left text-xs text-secondary-foreground transition-colors hover:bg-muted hover:text-foreground"
            disabled={fieldSettingLoading}
            onClick={() => void closeFieldPreset()}
          >
            <AppIcon name="close" size={13} />
            <span>
              <span className="block font-medium">关闭预设</span>
              <span className="mt-0.5 block text-[11px] text-muted-foreground">保留字段内容和资料库</span>
            </span>
          </button>
        </div>
      ) : null}

      {open && (
        <div className="absolute right-0 top-10 z-40 w-[min(30rem,calc(100vw-2rem))] rounded-xl bg-popover p-3 text-popover-foreground shadow-lg ring-1 ring-border/60">
          <div className="mb-3 flex items-start justify-between gap-3 rounded-lg bg-muted/35 p-3">
            <div className="min-w-0">
              <div className="text-sm font-semibold text-foreground">{field.label}</div>
              {field.description ? <div className="mt-1 text-xs text-secondary-foreground">{field.description}</div> : null}
              <div className="mt-2">
                <FieldBusinessMeta field={field} compact />
              </div>
            </div>
            <Button type="button" variant="ghost" size="icon" onClick={() => setOpen(false)} aria-label="关闭常用内容面板">
              <AppIcon name="close" size={14} />
            </Button>
          </div>

          <div className="max-h-72 space-y-2 overflow-y-auto">
            {loading ? (
              <div className="rounded-lg bg-muted/50 px-3 py-4 text-sm text-secondary-foreground">正在读取常用内容...</div>
            ) : presets.length === 0 ? (
              <div className="rounded-lg bg-muted/50 px-3 py-4 text-sm text-secondary-foreground">暂无可用于该字段的常用内容。</div>
            ) : (
              presets.map(preset => {
                const businessFields = listBusinessFields(preset.applicableFieldKeys);
                return (
                  <div key={preset.id} className="rounded-lg bg-muted/40 p-3">
                    <div className="flex items-start justify-between gap-3">
                      <div className="min-w-0">
                        <div className="truncate text-sm font-semibold text-foreground">{preset.name}</div>
                        <div className="mt-1 max-h-16 overflow-hidden whitespace-pre-wrap text-xs leading-5 text-secondary-foreground">
                          {preset.content}
                        </div>
                        <div className="mt-2 text-[11px] leading-4 text-muted-foreground">
                          {businessFields.length > 0
                            ? `适用字段：${businessFields.map(item => item.label).join("、")}`
                            : "适用字段：通用"}
                        </div>
                        <div className="mt-1 text-[11px] text-muted-foreground">
                          使用 {preset.usageCount} 次
                          {preset.lastUsedAt ? ` · 最近 ${new Date(preset.lastUsedAt).toLocaleString()}` : ""}
                        </div>
                      </div>
                      <div className="flex shrink-0 flex-col gap-1">
                        <Button type="button" size="sm" onClick={() => void applyPreset(preset, "replace")}>
                          替换
                        </Button>
                        {kind === "text_snippet" && (
                          <Button type="button" variant="outline" size="sm" onClick={() => void applyPreset(preset, "append")}>
                            追加
                          </Button>
                        )}
                      </div>
                    </div>
                  </div>
                );
              })
            )}
          </div>

          <div className="mt-3 rounded-lg bg-muted/30 p-3">
            <button
              type="button"
              className="flex w-full items-center justify-between text-left text-sm font-semibold text-foreground"
              onClick={() => setSaveOpen(current => !current)}
            >
              <span>保存当前字段为常用内容</span>
              <AppIcon name={saveOpen ? "chevronUp" : "chevronDown"} size={14} />
            </button>
            {saveOpen && (
              <div className="mt-3 grid gap-2">
                <div className="rounded-md bg-card/70 px-3 py-2">
                  <div className="text-xs font-semibold text-secondary-foreground">绑定字段：{field.label}</div>
                  <div className="mt-1 text-[11px] text-muted-foreground">适用模板：{field.templates.join("、")}</div>
                  <div className="text-[11px] text-muted-foreground">所属分组：{field.groups.join("、")}</div>
                </div>
                <Input value={saveName} onChange={event => setSaveName(event.target.value)} placeholder={field.label} />
                <Input value={saveCategory} onChange={event => setSaveCategory(event.target.value)} placeholder="分类" />
                <Input value={saveTags} onChange={event => setSaveTags(event.target.value)} placeholder="标签，可选" />
                <Button type="button" onClick={() => void saveCurrent()}>
                  保存为常用
                </Button>
              </div>
            )}
          </div>

          <div className="mt-3 flex items-center justify-between rounded-lg bg-muted/25 px-3 py-2">
            <span className="text-xs text-muted-foreground">该设置保存在当前工作区</span>
            <Button
              type="button"
              variant="ghost"
              size="sm"
              className="h-7 text-xs"
              disabled={fieldSettingLoading}
              onClick={() => void closeFieldPreset()}
            >
              关闭字段预设
            </Button>
          </div>

          {error ? <div className="mt-3 rounded-lg bg-destructive-soft px-3 py-2 text-xs text-destructive">{error}</div> : null}
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
