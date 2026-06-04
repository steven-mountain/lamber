import { useEffect, useMemo, useState } from "react";
import AppIcon from "../components/icons/AppIcon";
import { Button } from "../components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "../components/ui/card";
import { Input } from "../components/ui/input";
import {
  getPresetFieldCategories,
  PRESET_FIELD_DEFINITIONS,
  type CommonPresetKind,
} from "../lib/presetFieldKeys";
import {
  commonPresetService,
  type CommonPreset,
  type CommonPresetInput,
} from "../services/commonPresetService";

interface PresetCenterViewProps {
  onBack: () => void;
}

type SortMode = "recent" | "usage";

const EMPTY_FORM: CommonPresetInput = {
  scope: "workspace",
  kind: "short_value",
  category: "审核人员",
  name: "",
  content: "",
  tags: [],
  applicableFieldKeys: [],
  enabled: true,
};

function parseTags(raw: string): string[] {
  return raw
    .split(/[,\uFF0C\s]+/)
    .map(item => item.trim())
    .filter(Boolean);
}

function formatTime(value?: string | null) {
  if (!value) return "尚未使用";
  return new Date(value).toLocaleString();
}

export default function PresetCenterView({ onBack }: PresetCenterViewProps) {
  const [kind, setKind] = useState<CommonPresetKind>("short_value");
  const [category, setCategory] = useState("");
  const [sortBy, setSortBy] = useState<SortMode>("recent");
  const [items, setItems] = useState<CommonPreset[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [editing, setEditing] = useState<CommonPresetInput>({
    ...EMPTY_FORM,
  });
  const [tagText, setTagText] = useState("");
  const [fieldKeyDraft, setFieldKeyDraft] = useState<string[]>([]);

  const categories = useMemo(() => {
    const fromCatalog = getPresetFieldCategories(kind);
    const fromItems = items.filter(item => item.kind === kind).map(item => item.category);
    return Array.from(new Set([...fromCatalog, ...fromItems])).sort((a, b) => a.localeCompare(b, "zh-CN"));
  }, [items, kind]);

  const compatibleFields = PRESET_FIELD_DEFINITIONS.filter(field => field.kind === kind);

  const loadItems = async () => {
    setLoading(true);
    setError("");
    try {
      const result = await commonPresetService.list({
        kind,
        category: category || null,
        includeDisabled: true,
        sortBy,
      });
      setItems(result);
    } catch (err) {
      console.error("Failed to load common presets", err);
      setError("无法读取常用资料，请确认已打开工作区。");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void loadItems();
  }, [kind, category, sortBy]);

  const resetForm = (nextKind: CommonPresetKind = kind) => {
    const defaultCategory = getPresetFieldCategories(nextKind)[0] || "未分类";
    setEditing({
      ...EMPTY_FORM,
      kind: nextKind,
      category: defaultCategory,
      content: "",
      name: "",
    });
    setTagText("");
    setFieldKeyDraft([]);
  };

  const startEdit = (item: CommonPreset) => {
    setKind(item.kind);
    setEditing({
      id: item.id,
      scope: "workspace",
      kind: item.kind,
      category: item.category,
      name: item.name,
      content: item.content,
      tags: item.tags,
      applicableFieldKeys: item.applicableFieldKeys,
      enabled: item.enabled,
    });
    setTagText(item.tags.join(" "));
    setFieldKeyDraft(item.applicableFieldKeys);
  };

  const saveItem = async () => {
    setError("");
    try {
      await commonPresetService.save({
        ...editing,
        kind,
        category: editing.category.trim(),
        name: editing.name.trim(),
        content: editing.content.trim(),
        tags: parseTags(tagText),
        applicableFieldKeys: fieldKeyDraft,
        enabled: editing.enabled ?? true,
      });
      resetForm(kind);
      await loadItems();
    } catch (err) {
      console.error("Failed to save common preset", err);
      setError("保存失败，请检查名称、分类和正文是否已填写。");
    }
  };

  const toggleEnabled = async (item: CommonPreset) => {
    setError("");
    try {
      await commonPresetService.setEnabled(item.id, !item.enabled);
      await loadItems();
    } catch (err) {
      console.error("Failed to toggle preset", err);
      setError("更新启用状态失败。");
    }
  };

  const deleteItem = async (item: CommonPreset) => {
    const ok = window.confirm(`确定删除“${item.name}”吗？删除后不会再出现在填充列表中。`);
    if (!ok) return;
    setError("");
    try {
      await commonPresetService.delete(item.id);
      if (editing.id === item.id) resetForm(kind);
      await loadItems();
    } catch (err) {
      console.error("Failed to delete preset", err);
      setError("删除失败。");
    }
  };

  const handleKindChange = (nextKind: CommonPresetKind) => {
    setKind(nextKind);
    setCategory("");
    resetForm(nextKind);
  };

  const toggleFieldKey = (fieldKey: string) => {
    setFieldKeyDraft(current =>
      current.includes(fieldKey)
        ? current.filter(item => item !== fieldKey)
        : [...current, fieldKey],
    );
  };

  return (
    <div className="flex h-full min-h-0 flex-col overflow-hidden bg-background text-foreground">
      <header className="flex shrink-0 items-center justify-between bg-card px-6 py-4 shadow-sm">
        <div className="flex items-center gap-3">
          <Button type="button" variant="ghost" onClick={onBack}>
            <span>←</span>
            返回集市
          </Button>
          <div>
            <h1 className="text-page-title font-bold tracking-tight">常用资料与项目预设</h1>
            <p className="mt-0.5 text-caption text-secondary-foreground">管理工作区内可复用的短字段与长文本片段</p>
          </div>
        </div>
      </header>

      <main className="min-h-0 flex-1 overflow-y-auto p-6">
        <div className="mx-auto grid max-w-7xl gap-5 xl:grid-cols-[1fr_24rem]">
          <section className="space-y-4">
            <Card>
              <CardHeader className="pb-4">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div className="inline-flex rounded-lg bg-muted p-1">
                    <Button
                      type="button"
                      variant={kind === "short_value" ? "default" : "ghost"}
                      size="sm"
                      onClick={() => handleKindChange("short_value")}
                    >
                      常用字段
                    </Button>
                    <Button
                      type="button"
                      variant={kind === "text_snippet" ? "default" : "ghost"}
                      size="sm"
                      onClick={() => handleKindChange("text_snippet")}
                    >
                      常用文本
                    </Button>
                  </div>
                  <div className="flex flex-wrap items-center gap-2">
                    <select
                      value={category}
                      onChange={event => setCategory(event.target.value)}
                      className="h-[var(--density-input-height,2.25rem)] rounded-md border border-input bg-card px-3 text-sm shadow-sm"
                    >
                      <option value="">全部分类</option>
                      {categories.map(item => (
                        <option key={item} value={item}>{item}</option>
                      ))}
                    </select>
                    <select
                      value={sortBy}
                      onChange={event => setSortBy(event.target.value as SortMode)}
                      className="h-[var(--density-input-height,2.25rem)] rounded-md border border-input bg-card px-3 text-sm shadow-sm"
                    >
                      <option value="recent">最近使用优先</option>
                      <option value="usage">使用次数优先</option>
                    </select>
                  </div>
                </div>
              </CardHeader>
              <CardContent className="space-y-3">
                {loading ? (
                  <div className="rounded-xl bg-muted/50 p-6 text-sm text-secondary-foreground">正在读取常用资料...</div>
                ) : items.length === 0 ? (
                  <div className="rounded-xl bg-muted/50 p-6 text-sm text-secondary-foreground">当前分类下暂无资料，可在右侧新增。</div>
                ) : (
                  items.map(item => (
                    <div key={item.id} className={`rounded-xl bg-muted/40 p-4 ${item.enabled ? "" : "opacity-60"}`}>
                      <div className="flex flex-wrap items-start justify-between gap-3">
                        <div className="min-w-0 flex-1">
                          <div className="flex flex-wrap items-center gap-2">
                            <span className="text-body-strong font-semibold text-foreground">{item.name}</span>
                            <span className="rounded-md bg-card px-2 py-0.5 text-xs text-secondary-foreground">{item.category}</span>
                            <span className={`rounded-md px-2 py-0.5 text-xs ${item.enabled ? "bg-success-soft text-success" : "bg-muted text-muted-foreground"}`}>
                              {item.enabled ? "启用" : "停用"}
                            </span>
                          </div>
                          <div className="mt-2 max-h-24 overflow-hidden whitespace-pre-wrap text-sm leading-6 text-secondary-foreground">
                            {item.content}
                          </div>
                          <div className="mt-3 flex flex-wrap gap-3 text-xs text-muted-foreground">
                            <span>使用 {item.usageCount} 次</span>
                            <span>最近使用：{formatTime(item.lastUsedAt)}</span>
                            {item.tags.length > 0 ? <span>标签：{item.tags.join("、")}</span> : null}
                          </div>
                        </div>
                        <div className="flex shrink-0 flex-wrap gap-2">
                          <Button type="button" variant="outline" size="sm" onClick={() => startEdit(item)}>
                            <AppIcon name="edit" size={14} />
                            编辑
                          </Button>
                          <Button type="button" variant="secondary" size="sm" onClick={() => void toggleEnabled(item)}>
                            {item.enabled ? "停用" : "启用"}
                          </Button>
                          <Button type="button" variant="ghost" size="sm" onClick={() => void deleteItem(item)}>
                            <AppIcon name="delete" size={14} />
                            删除
                          </Button>
                        </div>
                      </div>
                    </div>
                  ))
                )}
              </CardContent>
            </Card>
          </section>

          <aside className="space-y-4">
            <Card>
              <CardHeader>
                <CardTitle className="text-section-title">{editing.id ? "编辑常用资料" : "新增常用资料"}</CardTitle>
              </CardHeader>
              <CardContent className="space-y-3">
                <div className="space-y-1">
                  <label className="text-label font-semibold text-secondary-foreground">名称</label>
                  <Input value={editing.name} onChange={event => setEditing(current => ({ ...current, name: event.target.value }))} />
                </div>
                <div className="space-y-1">
                  <label className="text-label font-semibold text-secondary-foreground">分类</label>
                  <Input value={editing.category} onChange={event => setEditing(current => ({ ...current, category: event.target.value }))} />
                </div>
                <div className="space-y-1">
                  <label className="text-label font-semibold text-secondary-foreground">正文内容</label>
                  <textarea
                    value={editing.content}
                    onChange={event => setEditing(current => ({ ...current, content: event.target.value }))}
                    rows={kind === "text_snippet" ? 8 : 3}
                    className="w-full rounded-md border border-input bg-card px-3 py-2 text-sm shadow-sm outline-none transition-colors hover:border-ring/50 focus:border-ring focus:ring-2 focus:ring-ring/20"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-label font-semibold text-secondary-foreground">标签</label>
                  <Input value={tagText} onChange={event => setTagText(event.target.value)} placeholder="可用空格或逗号分隔" />
                </div>
                <div className="space-y-2 rounded-lg bg-muted/40 p-3">
                  <div className="text-label font-semibold text-secondary-foreground">适用字段</div>
                  <div className="text-xs text-muted-foreground">不选择时仅按分类和类型管理；选择后可在表单快捷填充中精准匹配。</div>
                  <div className="grid gap-2">
                    {compatibleFields.map(field => (
                      <label key={field.fieldKey} className="flex items-center gap-2 text-sm text-foreground">
                        <input
                          type="checkbox"
                          checked={fieldKeyDraft.includes(field.fieldKey)}
                          onChange={() => toggleFieldKey(field.fieldKey)}
                        />
                        <span>{field.label}</span>
                        <span className="text-xs text-muted-foreground">{field.fieldKey}</span>
                      </label>
                    ))}
                  </div>
                </div>
                <label className="flex items-center gap-2 text-sm font-semibold text-secondary-foreground">
                  <input
                    type="checkbox"
                    checked={editing.enabled ?? true}
                    onChange={event => setEditing(current => ({ ...current, enabled: event.target.checked }))}
                  />
                  启用
                </label>
                <div className="flex gap-2">
                  <Button type="button" onClick={() => void saveItem()}>
                    <AppIcon name="save" size={14} />
                    保存
                  </Button>
                  <Button type="button" variant="outline" onClick={() => resetForm(kind)}>
                    新增空白
                  </Button>
                </div>
                {error ? <div className="rounded-lg bg-destructive-soft px-3 py-2 text-sm text-destructive">{error}</div> : null}
              </CardContent>
            </Card>
          </aside>
        </div>
      </main>
    </div>
  );
}
