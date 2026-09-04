import { useCallback, useEffect, useMemo, useRef, useState, type MouseEvent } from "react";
import AppIcon from "../components/icons/AppIcon";
import { Button } from "../components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "../components/ui/card";
import { Input } from "../components/ui/input";
import BusinessDictionaryManager from "../components/business-dictionaries/BusinessDictionaryManager";
import ProjectPresetManager from "../components/project-presets/ProjectPresetManager";
import {
  getPresetFieldDisplay,
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
type PanelMode = "create" | "edit" | null;
type CenterSection = "presets" | "dictionary" | "projectPresets";

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

function createEmptyForm(kind: CommonPresetKind): CommonPresetInput {
  return {
    ...EMPTY_FORM,
    kind,
    category: getPresetFieldCategories(kind)[0] || "未分类",
    name: "",
    content: "",
    tags: [],
    applicableFieldKeys: [],
    enabled: true,
  };
}

function serializeDraft(
  draft: CommonPresetInput,
  kind: CommonPresetKind,
  tagText: string,
  fieldKeys: string[],
) {
  return JSON.stringify({
    id: draft.id || null,
    kind,
    category: draft.category,
    name: draft.name,
    content: draft.content,
    tags: parseTags(tagText).sort(),
    applicableFieldKeys: [...fieldKeys].sort(),
    enabled: draft.enabled ?? true,
  });
}

function fieldSummary(fieldKeys: string[]) {
  if (fieldKeys.length === 0) return "未限定适用字段";
  const labels = fieldKeys.map(fieldKey => getPresetFieldDisplay(fieldKey).label);
  if (labels.length <= 2) return `适用：${labels.join("、")}`;
  return `适用 ${labels.length} 个字段`;
}

function matchesSearch(item: CommonPreset, query: string) {
  if (!query) return true;
  const fieldMeta = item.applicableFieldKeys.map(getPresetFieldDisplay);
  const haystack = [
    item.name,
    item.category,
    item.content,
    item.tags.join(" "),
    fieldMeta.map(field => field.label).join(" "),
    fieldMeta.flatMap(field => field.templates).join(" "),
    fieldMeta.flatMap(field => field.groups).join(" "),
  ].join(" ").toLocaleLowerCase();
  return haystack.includes(query);
}

export default function PresetCenterView({ onBack }: PresetCenterViewProps) {
  const panelRef = useRef<HTMLElement | null>(null);
  const [kind, setKind] = useState<CommonPresetKind>("short_value");
  const [section, setSection] = useState<CenterSection>("presets");
  const [category, setCategory] = useState("");
  const [sortBy, setSortBy] = useState<SortMode>("recent");
  const [searchText, setSearchText] = useState("");
  const [items, setItems] = useState<CommonPreset[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [panelMode, setPanelMode] = useState<PanelMode>(null);
  const [editing, setEditing] = useState<CommonPresetInput>(() => createEmptyForm("short_value"));
  const [tagText, setTagText] = useState("");
  const [fieldKeyDraft, setFieldKeyDraft] = useState<string[]>([]);
  const [draftBaseline, setDraftBaseline] = useState(() =>
    serializeDraft(createEmptyForm("short_value"), "short_value", "", []),
  );
  const [fieldSectionOpen, setFieldSectionOpen] = useState(false);

  const categories = useMemo(() => {
    const fromCatalog = getPresetFieldCategories(kind);
    const fromItems = items.filter(item => item.kind === kind).map(item => item.category);
    return Array.from(new Set([...fromCatalog, ...fromItems])).sort((a, b) => a.localeCompare(b, "zh-CN"));
  }, [items, kind]);

  const compatibleFields = useMemo(
    () => PRESET_FIELD_DEFINITIONS.filter(field => field.kind === kind),
    [kind],
  );

  const normalizedSearch = searchText.trim().toLocaleLowerCase();
  const filteredItems = useMemo(
    () => items.filter(item => matchesSearch(item, normalizedSearch)),
    [items, normalizedSearch],
  );

  const currentDraftSignature = useMemo(
    () => serializeDraft(editing, kind, tagText, fieldKeyDraft),
    [editing, fieldKeyDraft, kind, tagText],
  );

  const hasUnsavedChanges = panelMode !== null && currentDraftSignature !== draftBaseline;

  const confirmDiscard = () => {
    if (!hasUnsavedChanges) return true;
    return window.confirm("当前资料有未保存改动，确定放弃这些改动吗？");
  };
  const loadItems = useCallback(async () => {
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
  }, [category, kind, sortBy]);

  useEffect(() => {
    void loadItems();
  }, [loadItems]);

  const resetDraft = (nextKind: CommonPresetKind = kind) => {
    const nextDraft = createEmptyForm(nextKind);
    setEditing(nextDraft);
    setTagText("");
    setFieldKeyDraft([]);
    setDraftBaseline(serializeDraft(nextDraft, nextKind, "", []));
    setFieldSectionOpen(false);
  };

  const openCreate = () => {
    if (!confirmDiscard()) return;
    const nextDraft = createEmptyForm(kind);
    setEditing(nextDraft);
    setTagText("");
    setFieldKeyDraft([]);
    setPanelMode("create");
    setFieldSectionOpen(false);
    setDraftBaseline(serializeDraft(nextDraft, kind, "", []));
    setError("");
  };

  const closePanel = () => {
    if (!confirmDiscard()) return false;
    setPanelMode(null);
    resetDraft(kind);
    return true;
  };

  const handleMainMouseDownCapture = (event: MouseEvent<HTMLElement>) => {
    if (!panelMode || !panelRef.current) return;
    if (panelRef.current.contains(event.target as Node)) return;
    const closed = closePanel();
    if (!closed) {
      event.preventDefault();
      event.stopPropagation();
    }
  };

  const startEdit = (item: CommonPreset) => {
    if (!confirmDiscard()) return;
    const nextDraft: CommonPresetInput = {
      id: item.id,
      scope: "workspace",
      kind: item.kind,
      category: item.category,
      name: item.name,
      content: item.content,
      tags: item.tags,
      applicableFieldKeys: item.applicableFieldKeys,
      enabled: item.enabled,
    };
    setKind(item.kind);
    setEditing(nextDraft);
    setTagText(item.tags.join(" "));
    setFieldKeyDraft(item.applicableFieldKeys);
    setPanelMode("edit");
    setFieldSectionOpen(false);
    setDraftBaseline(serializeDraft(nextDraft, item.kind, item.tags.join(" "), item.applicableFieldKeys));
    setError("");
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
      setPanelMode(null);
      resetDraft(kind);
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
      if (editing.id === item.id) {
        setPanelMode(null);
        resetDraft(kind);
      }
      await loadItems();
    } catch (err) {
      console.error("Failed to delete preset", err);
      setError("删除失败。");
    }
  };

  const handleKindChange = (nextKind: CommonPresetKind) => {
    if (nextKind === kind) return;
    if (!confirmDiscard()) return;
    setKind(nextKind);
    setCategory("");
    setSearchText("");
    setPanelMode(null);
    resetDraft(nextKind);
  };

  const toggleFieldKey = (fieldKey: string) => {
    setFieldKeyDraft(current =>
      current.includes(fieldKey)
        ? current.filter(item => item !== fieldKey)
        : [...current, fieldKey],
    );
  };

  return (
    <div className="flex h-full min-h-0 flex-col overflow-hidden bg-background text-foreground" onMouseDownCapture={handleMainMouseDownCapture}>
      <header className="flex shrink-0 flex-wrap items-center justify-between gap-3 bg-card px-6 py-4 shadow-sm">
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
        <div className="inline-flex rounded-lg bg-muted p-1">
          <Button
            type="button"
            variant={section === "presets" && kind === "short_value" ? "default" : "ghost"}
            size="sm"
            onClick={() => {
              if (!confirmDiscard()) return;
              setSection("presets");
              handleKindChange("short_value");
            }}
          >
            常用字段
          </Button>
          <Button
            type="button"
            variant={section === "presets" && kind === "text_snippet" ? "default" : "ghost"}
            size="sm"
            onClick={() => {
              if (!confirmDiscard()) return;
              setSection("presets");
              handleKindChange("text_snippet");
            }}
          >
            常用文本
          </Button>
          <Button
            type="button"
            variant={section === "dictionary" ? "default" : "ghost"}
            size="sm"
            onClick={() => {
              if (!confirmDiscard()) return;
              setPanelMode(null);
              setSection("dictionary");
            }}
          >
            业务字典
          </Button>
          <Button
            type="button"
            variant={section === "projectPresets" ? "default" : "ghost"}
            size="sm"
            onClick={() => {
              if (!confirmDiscard()) return;
              setPanelMode(null);
              setSection("projectPresets");
            }}
          >
            项目预设模板
          </Button>
        </div>
      </header>

      <main className="min-h-0 flex-1 overflow-hidden p-6">
        {section === "dictionary" ? (
          <div className="mx-auto flex h-full min-h-0 max-w-7xl">
            <BusinessDictionaryManager />
          </div>
        ) : section === "projectPresets" ? (
          <div className="mx-auto flex h-full min-h-0 max-w-7xl">
            <ProjectPresetManager />
          </div>
        ) : (
        <div className={`mx-auto grid h-full min-h-0 max-w-7xl gap-5 ${panelMode ? "xl:grid-cols-[minmax(0,1fr)_minmax(360px,400px)]" : "grid-cols-1"}`}>
          <section className="min-h-0">
            <Card className="flex h-full min-h-0 flex-col overflow-hidden">
              <CardHeader className="shrink-0 pb-4">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div>
                    <CardTitle className="text-section-title">
                      {kind === "short_value" ? "常用字段" : "常用文本"}
                    </CardTitle>
                    <p className="mt-1 text-caption text-muted-foreground">
                      自由文本预设与受控业务选项分开管理
                    </p>
                  </div>
                  <div className="flex flex-1 flex-wrap items-center justify-end gap-2">
                    <div className="relative min-w-[220px] flex-1 sm:max-w-xs">
                      <AppIcon name="search" size={15} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground" />
                      <Input
                        value={searchText}
                        onChange={event => setSearchText(event.target.value)}
                        placeholder="搜索名称、分类、正文、标签"
                        className="pl-9 text-sm"
                      />
                    </div>
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
                    <Button type="button" onClick={openCreate}>
                      <span className="text-base leading-none">+</span>
                      新建资料
                    </Button>
                  </div>
                </div>
              </CardHeader>
              <CardContent className="min-h-0 flex-1 space-y-3 overflow-y-auto">
                {loading ? (
                  <div className="rounded-xl bg-muted/50 p-6 text-sm text-secondary-foreground">正在读取常用资料...</div>
                ) : items.length === 0 ? (
                  <div className="rounded-xl bg-muted/50 p-6 text-sm text-secondary-foreground">当前分类下暂无资料，可点击“新建资料”添加。</div>
                ) : filteredItems.length === 0 ? (
                  <div className="rounded-xl bg-muted/50 p-6 text-sm text-secondary-foreground">没有匹配搜索条件的资料。</div>
                ) : (
                  filteredItems.map(item => (
                    <div key={item.id} className={`rounded-xl bg-muted/40 p-4 shadow-sm transition-colors hover:bg-muted/60 ${item.enabled ? "" : "opacity-60"}`}>
                      <div className="flex flex-wrap items-start justify-between gap-4">
                        <div className="min-w-0 flex-1">
                          <div className="flex flex-wrap items-center gap-2">
                            <span className="text-body-strong font-semibold text-foreground">{item.name}</span>
                            <span className="rounded-md bg-card/90 px-2 py-0.5 text-xs text-secondary-foreground">{item.category}</span>
                            <span className={`rounded-md px-2 py-0.5 text-xs ${item.enabled ? "bg-success-soft text-success" : "bg-muted text-muted-foreground"}`}>
                              {item.enabled ? "启用" : "停用"}
                            </span>
                          </div>
                          <div
                            className="mt-2 overflow-hidden whitespace-pre-wrap text-sm leading-6 text-secondary-foreground"
                            style={{ display: "-webkit-box", WebkitBoxOrient: "vertical", WebkitLineClamp: 2 }}
                          >
                            {item.content}
                          </div>
                          {item.tags.length > 0 ? (
                            <div className="mt-3 flex flex-wrap gap-1.5">
                              {item.tags.map(tag => (
                                <span key={tag} className="rounded-md bg-card/85 px-2 py-0.5 text-xs text-muted-foreground">{tag}</span>
                              ))}
                            </div>
                          ) : null}
                          {item.applicableFieldKeys.length > 0 ? (
                            <div className="mt-3 grid gap-2 sm:grid-cols-2">
                              {item.applicableFieldKeys.map(fieldKey => {
                                const field = getPresetFieldDisplay(fieldKey);
                                return (
                                  <div key={fieldKey} className="rounded-lg bg-card/70 px-3 py-2 text-xs">
                                    <div className="font-semibold text-secondary-foreground">{field.label}</div>
                                    <div className="mt-1 text-muted-foreground">适用模板：{field.templates.join("、")}</div>
                                    <div className="text-muted-foreground">所属分组：{field.groups.join("、")}</div>
                                  </div>
                                );
                              })}
                            </div>
                          ) : (
                            <div className="mt-3 text-xs text-muted-foreground">适用字段：通用资料，未限定具体业务字段</div>
                          )}
                          <div className="mt-3 flex flex-wrap gap-x-4 gap-y-1 text-xs text-muted-foreground">
                            <span>使用 <span className="numeric-value">{item.usageCount}</span> 次</span>
                            <span>最近使用：{formatTime(item.lastUsedAt)}</span>
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

          {panelMode ? (
            <>
              <div className="fixed inset-0 z-20 bg-foreground/10 backdrop-blur-[1px] xl:hidden" onClick={closePanel} />
              <aside
                ref={panelRef}
                className="fixed inset-y-0 right-0 z-30 flex min-h-0 w-full max-w-[420px] p-3 sm:p-4 xl:static xl:z-auto xl:max-h-full xl:max-w-none xl:p-0"
              >
                <Card className="flex h-full max-h-full min-h-0 w-full flex-col overflow-hidden bg-card/95 shadow-md xl:max-h-[calc(100vh-9rem)]">
                  <CardHeader className="shrink-0 pb-4">
                    <div className="flex items-start justify-between gap-3">
                      <div>
                        <CardTitle className="text-section-title">
                          {panelMode === "edit" ? "编辑常用资料" : "新建常用资料"}
                        </CardTitle>
                        <p className="mt-1 text-caption text-muted-foreground">
                          {panelMode === "edit" ? "修改后保存会覆盖当前资料" : "保存后进入当前工作区资料库"}
                        </p>
                      </div>
                      <Button type="button" variant="ghost" size="icon" onClick={closePanel} aria-label="关闭面板">
                        <AppIcon name="close" size={16} />
                      </Button>
                    </div>
                  </CardHeader>

                  <CardContent className="min-h-0 flex-1 basis-0 space-y-4 overflow-y-auto pb-4">
                    <section className="space-y-3 rounded-lg bg-muted/35 p-3">
                      <div className="text-label font-semibold text-secondary-foreground">基础信息</div>
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
                          rows={kind === "text_snippet" ? 8 : 4}
                          className="w-full rounded-md border border-input bg-card px-3 py-2 text-sm shadow-sm outline-none transition-colors hover:border-ring/50 focus:border-ring focus:ring-2 focus:ring-ring/20"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-label font-semibold text-secondary-foreground">标签</label>
                        <Input value={tagText} onChange={event => setTagText(event.target.value)} placeholder="可用空格或逗号分隔" />
                      </div>
                    </section>

                    <section className="rounded-lg bg-muted/35 p-3">
                      <button
                        type="button"
                        className="flex w-full items-center justify-between gap-3 text-left"
                        onClick={() => setFieldSectionOpen(current => !current)}
                      >
                        <span>
                          <span className="block text-label font-semibold text-secondary-foreground">适用字段</span>
                          <span className="mt-0.5 block text-xs text-muted-foreground">{fieldSummary(fieldKeyDraft)}</span>
                        </span>
                        <AppIcon name={fieldSectionOpen ? "chevronUp" : "chevronDown"} size={16} className="text-muted-foreground" />
                      </button>
                      {fieldSectionOpen ? (
                        <div className="mt-3 space-y-2">
                          <div className="text-xs text-muted-foreground">不选择时仅按分类和类型管理；选择后可在表单快捷填充中精准匹配。</div>
                          <div className="grid gap-2">
                            {compatibleFields.map(field => (
                              <label key={field.fieldKey} className="flex items-start gap-2 rounded-md bg-card/70 px-2 py-2 text-sm text-foreground">
                                <input
                                  type="checkbox"
                                  className="mt-1"
                                  checked={fieldKeyDraft.includes(field.fieldKey)}
                                  onChange={() => toggleFieldKey(field.fieldKey)}
                                />
                                <span className="min-w-0">
                                  <span className="block font-medium">{field.label}</span>
                                  <span className="block text-xs text-muted-foreground">
                                    适用模板：{field.templates.join("、")}
                                  </span>
                                  <span className="block text-xs text-muted-foreground">
                                    所属分组：{field.groups.join("、")}
                                  </span>
                                </span>
                              </label>
                            ))}
                          </div>
                        </div>
                      ) : null}
                    </section>

                    <section className="rounded-lg bg-muted/35 p-3">
                      <div className="text-label font-semibold text-secondary-foreground">状态</div>
                      <label className="mt-2 flex items-center gap-2 text-sm font-semibold text-secondary-foreground">
                        <input
                          type="checkbox"
                          checked={editing.enabled ?? true}
                          onChange={event => setEditing(current => ({ ...current, enabled: event.target.checked }))}
                        />
                        启用
                      </label>
                    </section>

                    {error ? <div className="rounded-lg bg-destructive-soft px-3 py-2 text-sm text-destructive">{error}</div> : null}
                  </CardContent>

                  <div className="flex shrink-0 gap-2 bg-card/95 p-[var(--density-card-padding,1.5rem)] pt-3 shadow-[0_-8px_24px_hsl(var(--background)/0.85)]">
                    <Button type="button" onClick={() => void saveItem()}>
                      <AppIcon name="save" size={14} />
                      {panelMode === "edit" ? "保存修改" : "保存"}
                    </Button>
                    <Button type="button" variant="outline" onClick={closePanel}>
                      取消
                    </Button>
                  </div>
                </Card>
              </aside>
            </>
          ) : null}
        </div>
        )}
      </main>
    </div>
  );
}
