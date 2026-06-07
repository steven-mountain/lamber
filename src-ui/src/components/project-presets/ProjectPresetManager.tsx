import { useEffect, useMemo, useState } from "react";
import AppIcon from "../icons/AppIcon";
import { Button } from "../ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "../ui/card";
import { Input } from "../ui/input";
import { PRESET_FIELD_REGISTRY, getPresetFieldDefinition } from "../../lib/presetFieldKeys";
import {
  getProjectPresetValueType,
  isProjectPresetFieldAllowed,
  summarizeProjectPresetValue,
} from "../../lib/projectPresetFields";
import {
  projectPresetService,
  type ProjectPresetTemplate,
  type ProjectPresetTemplateEntryInput,
  type ProjectPresetTemplateInput,
} from "../../services/projectPresetService";

const eligibleFields = PRESET_FIELD_REGISTRY.filter(field =>
  isProjectPresetFieldAllowed(field.fieldKey),
);

const parseTags = (raw: string) =>
  raw.split(/[,\uFF0C\s]+/).map(item => item.trim()).filter(Boolean);

const emptyDraft = (): ProjectPresetTemplateInput => ({
  scope: "workspace",
  name: "",
  description: "",
  category: "",
  tags: [],
  enabled: true,
  entries: [],
});

export default function ProjectPresetManager() {
  const [templates, setTemplates] = useState<ProjectPresetTemplate[]>([]);
  const [draft, setDraft] = useState<ProjectPresetTemplateInput | null>(null);
  const [tagText, setTagText] = useState("");
  const [fieldToAdd, setFieldToAdd] = useState("");
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const load = async () => {
    setLoading(true);
    try {
      setTemplates(await projectPresetService.list(true));
      setError("");
    } catch (loadError) {
      console.error("Failed to load project presets", loadError);
      setError("读取项目预设模板失败。");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void load();
  }, []);

  const availableFields = useMemo(() => {
    const selected = new Set(draft?.entries.map(entry => entry.fieldKey) ?? []);
    return eligibleFields.filter(field => !selected.has(field.fieldKey));
  }, [draft?.entries]);

  const editTemplate = (template: ProjectPresetTemplate) => {
    setDraft({
      id: template.id,
      scope: "workspace",
      name: template.name,
      description: template.description || "",
      category: template.category,
      tags: template.tags,
      enabled: template.enabled,
      entries: template.entries.map(entry => ({
        id: entry.id,
        fieldKey: entry.fieldKey,
        value: entry.value,
        valueType: entry.valueType,
        sourceType: entry.sourceType,
        sortOrder: entry.sortOrder,
      })),
    });
    setTagText(template.tags.join(" "));
    setError("");
  };

  const addField = () => {
    if (!draft || !fieldToAdd) return;
    setDraft(current => current ? {
      ...current,
      entries: [
        ...current.entries,
        {
          fieldKey: fieldToAdd,
          value: "",
          valueType: getProjectPresetValueType(fieldToAdd),
          sourceType: getPresetFieldDefinition(fieldToAdd)?.dictionaryKey ? "dictionary" : "manual",
          sortOrder: (current.entries.length + 1) * 10,
        },
      ],
    } : current);
    setFieldToAdd("");
  };

  const updateEntry = (index: number, patch: Partial<ProjectPresetTemplateEntryInput>) => {
    setDraft(current => {
      if (!current) return current;
      const entries = [...current.entries];
      entries[index] = { ...entries[index], ...patch };
      return { ...current, entries };
    });
  };

  const removeEntry = (index: number) => {
    setDraft(current => current ? {
      ...current,
      entries: current.entries.filter((_, entryIndex) => entryIndex !== index),
    } : current);
  };

  const save = async () => {
    if (!draft) return;
    setSaving(true);
    try {
      await projectPresetService.save({
        ...draft,
        name: draft.name.trim(),
        description: draft.description?.trim() || null,
        category: draft.category?.trim() || "",
        tags: parseTags(tagText),
      });
      setDraft(null);
      await load();
    } catch (saveError) {
      console.error("Failed to save project preset", saveError);
      setError("保存失败。请确认名称已填写，且每个字段都有非空值。");
    } finally {
      setSaving(false);
    }
  };

  const toggleEnabled = async (template: ProjectPresetTemplate) => {
    await projectPresetService.setEnabled(template.id, !template.enabled);
    await load();
  };

  const deleteTemplate = async (template: ProjectPresetTemplate) => {
    if (!window.confirm(`确定删除“${template.name}”吗？`)) return;
    await projectPresetService.delete(template.id);
    if (draft?.id === template.id) setDraft(null);
    await load();
  };

  return (
    <div className={`grid h-full min-h-0 w-full gap-5 ${draft ? "xl:grid-cols-[minmax(0,1fr)_400px]" : "grid-cols-1"}`}>
      <Card className="flex min-h-0 flex-col overflow-hidden">
        <CardHeader className="shrink-0">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <CardTitle>项目预设模板</CardTitle>
              <p className="mt-1 text-caption text-muted-foreground">将多个安全业务字段组合为可复用的项目初始化方案</p>
            </div>
            <Button type="button" onClick={() => { setDraft(emptyDraft()); setTagText(""); }}>
              <span className="text-base">+</span>
              新建项目预设
            </Button>
          </div>
        </CardHeader>
        <CardContent className="min-h-0 flex-1 space-y-3 overflow-y-auto">
          {loading ? (
            <div className="rounded-lg bg-muted/40 p-5 text-sm text-muted-foreground">正在读取项目预设...</div>
          ) : templates.length === 0 ? (
            <div className="rounded-lg bg-muted/40 p-5 text-sm text-muted-foreground">暂无项目预设模板。</div>
          ) : templates.map(template => (
            <div key={template.id} className={`rounded-xl bg-muted/40 p-4 shadow-sm ${template.enabled ? "" : "opacity-60"}`}>
              <div className="flex flex-wrap items-start justify-between gap-3">
                <div className="min-w-0 flex-1">
                  <div className="flex flex-wrap items-center gap-2">
                    <strong>{template.name}</strong>
                    {template.category ? <span className="rounded-md bg-card px-2 py-0.5 text-xs text-muted-foreground">{template.category}</span> : null}
                    <span className="rounded-md bg-card px-2 py-0.5 text-xs text-muted-foreground">{template.enabled ? "启用" : "停用"}</span>
                  </div>
                  {template.description ? <p className="mt-2 text-sm text-secondary-foreground">{template.description}</p> : null}
                  <div className="mt-3 grid gap-2 sm:grid-cols-2">
                    {template.entries.slice(0, 6).map(entry => {
                      const field = getPresetFieldDefinition(entry.fieldKey);
                      return (
                        <div key={entry.id} className="rounded-lg bg-card/75 px-3 py-2">
                          <div className="text-sm font-semibold">{field?.label || "未命名字段"}</div>
                          <div className="mt-1 text-xs text-muted-foreground">
                            {field?.templates.join("、") || "暂未配置模板"} · {field?.groups.join("、") || "暂未配置分组"}
                          </div>
                          <div className="mt-1 truncate text-xs text-secondary-foreground">{summarizeProjectPresetValue(entry.value)}</div>
                        </div>
                      );
                    })}
                  </div>
                  {template.entries.length > 6 ? <div className="mt-2 text-xs text-muted-foreground">另有 {template.entries.length - 6} 个字段</div> : null}
                </div>
                <div className="flex gap-2">
                  <Button type="button" size="sm" variant="outline" onClick={() => editTemplate(template)}>
                    <AppIcon name="edit" size={14} />编辑
                  </Button>
                  <Button type="button" size="sm" variant="secondary" onClick={() => void toggleEnabled(template)}>
                    {template.enabled ? "停用" : "启用"}
                  </Button>
                  <Button type="button" size="sm" variant="ghost" onClick={() => void deleteTemplate(template)}>
                    <AppIcon name="delete" size={14} />删除
                  </Button>
                </div>
              </div>
            </div>
          ))}
          {error ? <div className="rounded-lg bg-destructive-soft px-3 py-2 text-sm text-destructive">{error}</div> : null}
        </CardContent>
      </Card>

      {draft ? (
        <Card className="flex min-h-0 flex-col overflow-hidden">
          <CardHeader className="shrink-0">
            <div className="flex items-start justify-between gap-3">
              <div>
                <CardTitle>{draft.id ? "编辑项目预设" : "新建项目预设"}</CardTitle>
                <p className="mt-1 text-caption text-muted-foreground">金额、税率、比例和自动计算字段不会出现在可选列表</p>
              </div>
              <Button type="button" variant="ghost" size="icon" onClick={() => setDraft(null)}>
                <AppIcon name="close" size={16} />
              </Button>
            </div>
          </CardHeader>
          <CardContent className="min-h-0 flex-1 space-y-4 overflow-y-auto">
            <section className="space-y-3 rounded-lg bg-muted/35 p-3">
              <label className="block text-label font-semibold">名称
                <Input className="mt-1" value={draft.name} onChange={event => setDraft(current => current ? { ...current, name: event.target.value } : current)} />
              </label>
              <label className="block text-label font-semibold">描述
                <textarea className="mt-1 min-h-20 w-full rounded-md border border-input bg-card px-3 py-2 text-sm" value={draft.description || ""} onChange={event => setDraft(current => current ? { ...current, description: event.target.value } : current)} />
              </label>
              <label className="block text-label font-semibold">分类
                <Input className="mt-1" value={draft.category || ""} onChange={event => setDraft(current => current ? { ...current, category: event.target.value } : current)} />
              </label>
              <label className="block text-label font-semibold">标签
                <Input className="mt-1" value={tagText} onChange={event => setTagText(event.target.value)} />
              </label>
            </section>

            <section className="space-y-3 rounded-lg bg-muted/35 p-3">
              <div className="text-label font-semibold">字段项</div>
              <div className="flex gap-2">
                <select value={fieldToAdd} onChange={event => setFieldToAdd(event.target.value)} className="min-w-0 flex-1 rounded-md border border-input bg-card px-3 text-sm">
                  <option value="">选择要添加的字段</option>
                  {availableFields.map(field => <option key={field.fieldKey} value={field.fieldKey}>{field.label} · {field.templates.join("、")}</option>)}
                </select>
                <Button type="button" variant="outline" onClick={addField} disabled={!fieldToAdd}>添加</Button>
              </div>
              {draft.entries.map((entry, index) => {
                const field = getPresetFieldDefinition(entry.fieldKey);
                return (
                  <div key={entry.fieldKey} className="rounded-lg bg-card/80 p-3">
                    <div className="flex items-start justify-between gap-3">
                      <div>
                        <div className="font-semibold">{field?.label || "未命名字段"}</div>
                        <div className="mt-1 text-xs text-muted-foreground">所属模板：{field?.templates.join("、")} · 所属分组：{field?.groups.join("、")}</div>
                        <div className="text-xs text-muted-foreground">字段类型：{entry.valueType} · 可应用</div>
                      </div>
                      <Button type="button" size="icon" variant="ghost" onClick={() => removeEntry(index)}><AppIcon name="delete" size={14} /></Button>
                    </div>
                    <textarea className="mt-3 min-h-16 w-full rounded-md border border-input bg-card px-3 py-2 text-sm" value={String(entry.value ?? "")} onChange={event => updateEntry(index, { value: event.target.value })} />
                  </div>
                );
              })}
            </section>
          </CardContent>
          <div className="flex shrink-0 gap-2 bg-card p-4 shadow-[0_-8px_24px_hsl(var(--background)/0.85)]">
            <Button type="button" onClick={() => void save()} disabled={saving}><AppIcon name="save" size={14} />{saving ? "保存中..." : "保存项目预设"}</Button>
            <Button type="button" variant="outline" onClick={() => setDraft(null)}>取消</Button>
          </div>
        </Card>
      ) : null}
    </div>
  );
}
