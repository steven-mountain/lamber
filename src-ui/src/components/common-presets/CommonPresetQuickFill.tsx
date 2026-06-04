import { type ReactNode, useEffect, useMemo, useState } from "react";
import AppIcon from "../icons/AppIcon";
import { Button } from "../ui/button";
import { Input } from "../ui/input";
import {
  getPresetFieldDefinition,
  type CommonPresetKind,
} from "../../lib/presetFieldKeys";
import {
  commonPresetService,
  type CommonPreset,
} from "../../services/commonPresetService";

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

export default function CommonPresetQuickFill({
  fieldKey,
  kind,
  value,
  onApply,
  className = "",
}: CommonPresetQuickFillProps) {
  const field = useMemo(() => getPresetFieldDefinition(fieldKey), [fieldKey]);
  const [open, setOpen] = useState(false);
  const [presets, setPresets] = useState<CommonPreset[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [saveOpen, setSaveOpen] = useState(false);
  const [saveName, setSaveName] = useState("");
  const [saveCategory, setSaveCategory] = useState(field?.category || "");
  const [saveTags, setSaveTags] = useState("");

  useEffect(() => {
    setSaveCategory(field?.category || "");
  }, [field?.category]);

  const loadPresets = async () => {
    setLoading(true);
    setError("");
    try {
      const items = await commonPresetService.list({
        kind,
        fieldKey,
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

  const handleToggle = async () => {
    const nextOpen = !open;
    setOpen(nextOpen);
    if (nextOpen) {
      await loadPresets();
    }
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
    const name = saveName.trim() || field?.label || "常用内容";
    const category = saveCategory.trim() || field?.category || "未分类";
    setError("");
    try {
      await commonPresetService.save({
        scope: "workspace",
        kind,
        category,
        name,
        content,
        tags: splitTags(saveTags),
        applicableFieldKeys: [fieldKey],
        enabled: true,
      });
      setSaveName("");
      setSaveTags("");
      setSaveOpen(false);
      await loadPresets();
    } catch (err) {
      console.error("Failed to save common preset", err);
      setError("保存常用内容失败，请确认已打开工作区。");
    }
  };

  return (
    <div className={`relative inline-flex items-center gap-1 ${className}`}>
      <Button type="button" variant="secondary" size="sm" className="h-6 px-2 text-xs" onClick={handleToggle}>
        <AppIcon name="quickAction" size={13} />
        常用
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
        存为常用
      </Button>

      {open && (
        <div className="absolute right-0 top-10 z-40 w-[min(28rem,calc(100vw-2rem))] rounded-xl bg-popover p-3 text-popover-foreground shadow-lg ring-1 ring-border/60">
          <div className="mb-3 flex items-center justify-between gap-3">
            <div>
              <div className="text-sm font-semibold text-foreground">{field?.label || "常用内容"}</div>
              <div className="text-xs text-secondary-foreground">{fieldKey}</div>
            </div>
            <Button type="button" variant="ghost" size="icon" onClick={() => setOpen(false)} aria-label="关闭常用内容面板">
              <AppIcon name="close" size={14} />
            </Button>
          </div>

          <div className="space-y-2">
            {loading ? (
              <div className="rounded-lg bg-muted/50 px-3 py-4 text-sm text-secondary-foreground">正在读取常用内容...</div>
            ) : presets.length === 0 ? (
              <div className="rounded-lg bg-muted/50 px-3 py-4 text-sm text-secondary-foreground">暂无可用于该字段的常用内容。</div>
            ) : (
              presets.map(preset => (
                <div key={preset.id} className="rounded-lg bg-muted/40 p-3">
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0">
                      <div className="truncate text-sm font-semibold text-foreground">{preset.name}</div>
                      <div className="mt-1 max-h-16 overflow-hidden whitespace-pre-wrap text-xs leading-5 text-secondary-foreground">
                        {preset.content}
                      </div>
                      <div className="mt-2 text-[11px] text-muted-foreground">
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
              ))
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
                <Input value={saveName} onChange={event => setSaveName(event.target.value)} placeholder={field?.label || "名称"} />
                <Input value={saveCategory} onChange={event => setSaveCategory(event.target.value)} placeholder="分类" />
                <Input value={saveTags} onChange={event => setSaveTags(event.target.value)} placeholder="标签，可选" />
                <Button type="button" onClick={() => void saveCurrent()}>
                  保存为常用
                </Button>
              </div>
            )}
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
