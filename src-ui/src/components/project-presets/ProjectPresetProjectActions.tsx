import { useEffect, useMemo, useState } from "react";
import AppIcon from "../icons/AppIcon";
import { Button } from "../ui/button";
import { Input } from "../ui/input";
import { getPresetFieldDefinition } from "../../lib/presetFieldKeys";
import {
  isProjectPresetValueEmpty,
  summarizeProjectPresetValue,
  type ProjectPresetFieldBinding,
} from "../../lib/projectPresetFields";
import {
  projectPresetService,
  type ProjectPresetTemplate,
} from "../../services/projectPresetService";
import { useSaveStore } from "../../store/useSaveStore";

type ApplyStrategy = "fill_empty_only" | "overwrite_all" | "selected_fields";
type DialogMode = "save" | "apply" | null;

interface Props {
  bindings: ProjectPresetFieldBinding[];
}

const waitForReact = () =>
  new Promise<void>(resolve => {
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
  });

export default function ProjectPresetProjectActions({ bindings }: Props) {
  const [mode, setMode] = useState<DialogMode>(null);
  const [templates, setTemplates] = useState<ProjectPresetTemplate[]>([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState("");
  const [strategy, setStrategy] = useState<ApplyStrategy>("fill_empty_only");
  const [selectedFields, setSelectedFields] = useState<Set<string>>(new Set());
  const [name, setName] = useState("");
  const [description, setDescription] = useState("");
  const [category, setCategory] = useState("");
  const [tags, setTags] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const saveCurrentProject = useSaveStore(state => state.saveCurrentProject);

  const usableBindings = useMemo(
    () => bindings.filter(binding => !isProjectPresetValueEmpty(binding.value)),
    [bindings],
  );
  const bindingMap = useMemo(
    () => new Map(bindings.map(binding => [binding.fieldKey, binding])),
    [bindings],
  );
  const selectedTemplate = templates.find(template => template.id === selectedTemplateId) || null;

  const loadTemplates = async () => {
    const items = await projectPresetService.list(false);
    setTemplates(items);
    setSelectedTemplateId(current =>
      current && items.some(item => item.id === current) ? current : items[0]?.id || "",
    );
  };

  useEffect(() => {
    if (mode === "apply") {
      void loadTemplates().catch(loadError => {
        console.error("Failed to load project presets", loadError);
        setError("读取项目预设失败。");
      });
    }
  }, [mode]);

  useEffect(() => {
    if (mode === "save") {
      setSelectedFields(new Set(usableBindings.map(binding => binding.fieldKey)));
    }
  }, [mode, usableBindings]);

  useEffect(() => {
    if (!selectedTemplate) return;
    setSelectedFields(new Set(selectedTemplate.entries.map(entry => entry.fieldKey)));
  }, [selectedTemplate, selectedTemplateId]);

  const previewRows = useMemo(() => {
    if (!selectedTemplate) return [];
    return selectedTemplate.entries.map(entry => {
      const binding = bindingMap.get(entry.fieldKey);
      const currentValue = binding?.value;
      const selected = selectedFields.has(entry.fieldKey);
      let action: "填充" | "覆盖" | "跳过" = "跳过";
      if (binding && selected) {
        if (strategy === "overwrite_all") {
          action = isProjectPresetValueEmpty(currentValue) ? "填充" : "覆盖";
        } else if (strategy === "selected_fields") {
          action = isProjectPresetValueEmpty(currentValue) ? "填充" : "覆盖";
        } else if (isProjectPresetValueEmpty(currentValue)) {
          action = "填充";
        }
      }
      return { entry, binding, currentValue, selected, action };
    });
  }, [bindingMap, selectedFields, selectedTemplate, strategy]);

  const saveAsPreset = async () => {
    const entries = usableBindings.filter(binding => selectedFields.has(binding.fieldKey));
    if (!name.trim() || entries.length === 0) {
      setError("请填写预设名称并至少选择一个非空字段。");
      return;
    }
    setBusy(true);
    try {
      await projectPresetService.save({
        name: name.trim(),
        description: description.trim() || null,
        category: category.trim(),
        tags: tags.split(/[,\uFF0C\s]+/).filter(Boolean),
        enabled: true,
        entries: entries.map((binding, index) => ({
          fieldKey: binding.fieldKey,
          value: binding.value,
          valueType: binding.valueType,
          sourceType: binding.sourceType || "from_project",
          sortOrder: (index + 1) * 10,
        })),
      });
      setMode(null);
      setName("");
      setDescription("");
      setCategory("");
      setTags("");
      alert("已保存为项目预设模板。");
    } catch (saveError) {
      console.error("Failed to save project preset from project", saveError);
      setError("保存项目预设失败。");
    } finally {
      setBusy(false);
    }
  };

  const applyPreset = async () => {
    const rows = previewRows.filter(row => row.action !== "跳过" && row.binding);
    if (rows.length === 0) {
      setError("当前没有可应用字段。请检查策略、逐项选择或切换到对应模板表单。");
      return;
    }
    if (
      rows.some(row => row.action === "覆盖") &&
      !window.confirm("本次应用会覆盖已有字段内容。确认继续吗？")
    ) {
      return;
    }
    setBusy(true);
    try {
      rows.forEach(row => row.binding?.apply(row.entry.value));
      await waitForReact();
      await saveCurrentProject();
      setMode(null);
      alert(`项目预设应用成功，已更新并保存 ${rows.length} 个字段。`);
    } catch (applyError) {
      console.error("Failed to apply project preset", applyError);
      setError(`应用失败：${applyError instanceof Error ? applyError.message : String(applyError)}`);
    } finally {
      setBusy(false);
    }
  };

  return (
    <>
      <Button type="button" size="sm" variant="outline" onClick={() => { setError(""); setMode("save"); }}>
        <AppIcon name="presets" size={14} />
        保存为项目预设
      </Button>
      <Button type="button" size="sm" variant="outline" onClick={() => { setError(""); setMode("apply"); }}>
        <AppIcon name="presetLibrary" size={14} />
        应用项目预设
      </Button>

      {mode ? (
        <div className="fixed inset-0 z-[80] flex items-center justify-center bg-background/80 p-4 backdrop-blur-sm">
          <div className="flex max-h-[88vh] w-full max-w-4xl flex-col overflow-hidden rounded-xl bg-card shadow-xl">
            <div className="flex shrink-0 items-start justify-between gap-3 bg-muted/35 px-5 py-4">
              <div>
                <h2 className="text-section-title font-semibold">
                  {mode === "save" ? "从当前项目保存为预设" : "应用项目预设"}
                </h2>
                <p className="mt-1 text-caption text-muted-foreground">
                  {mode === "save"
                    ? "仅列出当前已填写、已挂载且符合安全规则的字段"
                    : "确认前请核对当前值、预设值和实际动作"}
                </p>
              </div>
              <Button type="button" size="icon" variant="ghost" onClick={() => setMode(null)}>
                <AppIcon name="close" size={16} />
              </Button>
            </div>

            <div className="min-h-0 flex-1 space-y-4 overflow-y-auto p-5">
              {mode === "save" ? (
                <>
                  <div className="grid gap-3 md:grid-cols-2">
                    <label className="text-label font-semibold">预设名称
                      <Input className="mt-1" value={name} onChange={event => setName(event.target.value)} />
                    </label>
                    <label className="text-label font-semibold">分类
                      <Input className="mt-1" value={category} onChange={event => setCategory(event.target.value)} />
                    </label>
                    <label className="text-label font-semibold md:col-span-2">描述
                      <textarea className="mt-1 min-h-16 w-full rounded-md border border-input bg-card px-3 py-2 text-sm" value={description} onChange={event => setDescription(event.target.value)} />
                    </label>
                    <label className="text-label font-semibold md:col-span-2">标签
                      <Input className="mt-1" value={tags} onChange={event => setTags(event.target.value)} placeholder="空格或逗号分隔" />
                    </label>
                  </div>
                  <div className="space-y-2">
                    {usableBindings.map(binding => {
                      const field = getPresetFieldDefinition(binding.fieldKey);
                      return (
                        <label key={binding.fieldKey} className="flex items-start gap-3 rounded-lg bg-muted/35 p-3">
                          <input
                            type="checkbox"
                            className="mt-1"
                            checked={selectedFields.has(binding.fieldKey)}
                            onChange={event => {
                              setSelectedFields(current => {
                                const next = new Set(current);
                                if (event.target.checked) next.add(binding.fieldKey);
                                else next.delete(binding.fieldKey);
                                return next;
                              });
                            }}
                          />
                          <span className="min-w-0">
                            <span className="block font-semibold">{field?.label || "未命名字段"}</span>
                            <span className="block text-xs text-muted-foreground">
                              {field?.templates.join("、")} · {field?.groups.join("、")} · {binding.valueType}
                            </span>
                            <span className="mt-1 block truncate text-sm text-secondary-foreground">
                              {summarizeProjectPresetValue(binding.value)}
                            </span>
                          </span>
                        </label>
                      );
                    })}
                  </div>
                </>
              ) : (
                <>
                  <div className="grid gap-3 md:grid-cols-[1fr_auto]">
                    <select
                      value={selectedTemplateId}
                      onChange={event => setSelectedTemplateId(event.target.value)}
                      className="rounded-md border border-input bg-card px-3 py-2 text-sm"
                    >
                      {templates.length === 0 ? <option value="">暂无可用项目预设</option> : null}
                      {templates.map(template => <option key={template.id} value={template.id}>{template.name}</option>)}
                    </select>
                    <div className="inline-flex rounded-lg bg-muted p-1">
                      {([
                        ["fill_empty_only", "仅填充空字段"],
                        ["overwrite_all", "覆盖已有字段"],
                        ["selected_fields", "逐项选择"],
                      ] as const).map(([value, label]) => (
                        <Button key={value} type="button" size="sm" variant={strategy === value ? "default" : "ghost"} onClick={() => setStrategy(value)}>
                          {label}
                        </Button>
                      ))}
                    </div>
                  </div>
                  <div className="space-y-2">
                    {previewRows.map(row => {
                      const field = getPresetFieldDefinition(row.entry.fieldKey);
                      return (
                        <div key={row.entry.id} className="grid gap-3 rounded-lg bg-muted/35 p-3 md:grid-cols-[28px_1.1fr_1fr_1fr_70px]">
                          <input
                            type="checkbox"
                            className="mt-1"
                            disabled={!row.binding}
                            checked={row.selected && Boolean(row.binding)}
                            onChange={event => {
                              setSelectedFields(current => {
                                const next = new Set(current);
                                if (event.target.checked) next.add(row.entry.fieldKey);
                                else next.delete(row.entry.fieldKey);
                                return next;
                              });
                            }}
                          />
                          <div>
                            <div className="font-semibold">{field?.label || "未命名字段"}</div>
                            <div className="text-xs text-muted-foreground">{field?.templates.join("、")} · {field?.groups.join("、")}</div>
                          </div>
                          <div>
                            <div className="text-xs text-muted-foreground">当前值</div>
                            <div className="mt-1 text-sm">{row.binding ? summarizeProjectPresetValue(row.currentValue) || "空" : "当前页面未挂载"}</div>
                          </div>
                          <div>
                            <div className="text-xs text-muted-foreground">预设值</div>
                            <div className="mt-1 text-sm">{summarizeProjectPresetValue(row.entry.value)}</div>
                          </div>
                          <div className={`text-sm font-semibold ${row.action === "覆盖" ? "text-warning" : row.action === "填充" ? "text-success" : "text-muted-foreground"}`}>
                            {row.action}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </>
              )}
              {error ? <div className="rounded-lg bg-destructive-soft px-3 py-2 text-sm text-destructive">{error}</div> : null}
            </div>

            <div className="flex shrink-0 justify-end gap-2 bg-card px-5 py-4 shadow-[0_-8px_24px_hsl(var(--background)/0.85)]">
              <Button type="button" onClick={() => void (mode === "save" ? saveAsPreset() : applyPreset())} disabled={busy}>
                {busy ? "处理中..." : mode === "save" ? "保存项目预设" : "确认应用并保存"}
              </Button>
              <Button type="button" variant="outline" onClick={() => setMode(null)}>取消</Button>
            </div>
          </div>
        </div>
      ) : null}
    </>
  );
}
