import { useCallback, useEffect, useMemo, useState } from "react";
import AppIcon from "../icons/AppIcon";
import { Button } from "../ui/button";
import { Input } from "../ui/input";
import { getPresetFieldDisplay } from "../../lib/presetFieldKeys";
import {
  businessDictionaryService,
  type BusinessDictionary,
  type BusinessDictionaryItem,
} from "../../services/businessDictionaryService";

interface ItemDraft {
  id?: string;
  value: string;
  label: string;
  description: string;
}

const EMPTY_DRAFT: ItemDraft = {
  value: "",
  label: "",
  description: "",
};

export default function BusinessDictionaryManager() {
  const [dictionaries, setDictionaries] = useState<BusinessDictionary[]>([]);
  const [selectedKey, setSelectedKey] = useState("");
  const [draft, setDraft] = useState<ItemDraft>(EMPTY_DRAFT);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    setError("");
    try {
      const result = await businessDictionaryService.list(true);
      setDictionaries(result);
      setSelectedKey(current =>
        current && result.some(item => item.dictionaryKey === current)
          ? current
          : result[0]?.dictionaryKey ?? "",
      );
    } catch (loadError) {
      console.error("Failed to load business dictionaries", loadError);
      setError("无法读取业务字典，请确认已打开工作区。");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const selected = useMemo(
    () => dictionaries.find(item => item.dictionaryKey === selectedKey),
    [dictionaries, selectedKey],
  );

  const resetDraft = () => setDraft(EMPTY_DRAFT);

  const saveDraft = async () => {
    if (!selected) return;
    setError("");
    try {
      await businessDictionaryService.saveItem({
        id: draft.id,
        dictionaryId: selected.id,
        value: draft.value.trim(),
        label: draft.label.trim(),
        description: draft.description.trim() || null,
        enabled: true,
        sortOrder: draft.id
          ? selected.items.find(item => item.id === draft.id)?.sortOrder ?? 0
          : (selected.items.at(-1)?.sortOrder ?? 0) + 10,
      });
      resetDraft();
      await load();
    } catch (saveError) {
      console.error("Failed to save business dictionary item", saveError);
      setError("保存失败。请检查存储值和显示名称是否为空或重复。");
    }
  };

  const toggleItem = async (item: BusinessDictionaryItem) => {
    setError("");
    try {
      await businessDictionaryService.setItemEnabled(item.id, !item.enabled);
      await load();
    } catch (toggleError) {
      console.error("Failed to toggle business dictionary item", toggleError);
      setError("更新字典项状态失败。");
    }
  };

  const deleteItem = async (item: BusinessDictionaryItem) => {
    if (!window.confirm(`确定删除“${item.label}”吗？旧项目中已保存的值不会被改写。`)) return;
    setError("");
    try {
      await businessDictionaryService.deleteItem(item.id);
      if (draft.id === item.id) resetDraft();
      await load();
    } catch (deleteError) {
      console.error("Failed to delete business dictionary item", deleteError);
      setError("删除字典项失败。");
    }
  };

  const moveItem = async (item: BusinessDictionaryItem, direction: -1 | 1) => {
    if (!selected) return;
    const index = selected.items.findIndex(current => current.id === item.id);
    const target = index + direction;
    if (index < 0 || target < 0 || target >= selected.items.length) return;
    const ids = selected.items.map(current => current.id);
    [ids[index], ids[target]] = [ids[target], ids[index]];
    setError("");
    try {
      await businessDictionaryService.reorderItems(selected.id, ids);
      await load();
    } catch (moveError) {
      console.error("Failed to reorder business dictionary items", moveError);
      setError("调整排序失败。");
    }
  };

  return (
    <div className="grid min-h-0 flex-1 gap-4 lg:grid-cols-[280px_minmax(0,1fr)]">
      <aside className="min-h-0 overflow-y-auto rounded-xl bg-muted/35 p-3">
        <div className="mb-3">
          <h2 className="text-section-title font-semibold">业务字典</h2>
          <p className="mt-1 text-caption text-muted-foreground">
            管理下拉框、是否项等受控业务选项。
          </p>
        </div>
        <div className="space-y-2">
          {dictionaries.map(dictionary => (
            <button
              key={dictionary.id}
              type="button"
              onClick={() => {
                setSelectedKey(dictionary.dictionaryKey);
                resetDraft();
              }}
              className={`w-full rounded-lg px-3 py-2 text-left transition-colors ${
                dictionary.dictionaryKey === selectedKey
                  ? "bg-card text-foreground shadow-sm"
                  : "text-secondary-foreground hover:bg-card/70"
              }`}
            >
              <span className="block text-sm font-semibold">{dictionary.name}</span>
              <span className="mt-0.5 block text-xs text-muted-foreground">
                {dictionary.items.filter(item => item.enabled).length} 个启用选项
              </span>
            </button>
          ))}
        </div>
      </aside>

      <section className="min-h-0 overflow-y-auto rounded-xl bg-card p-[var(--density-card-padding,1.5rem)] shadow-sm">
        {loading && !selected ? (
          <div className="rounded-lg bg-muted/40 p-4 text-sm text-secondary-foreground">
            正在读取业务字典...
          </div>
        ) : selected ? (
          <div className="space-y-5">
            <div>
              <h2 className="text-section-title font-semibold">{selected.name}</h2>
              <p className="mt-1 text-sm text-secondary-foreground">{selected.description}</p>
              <div className="mt-3 flex flex-wrap gap-2">
                {selected.applicableFieldKeys.map(fieldKey => {
                  const field = getPresetFieldDisplay(fieldKey);
                  return (
                    <span key={fieldKey} className="rounded-md bg-muted px-2 py-1 text-xs text-secondary-foreground">
                      {field.label}
                    </span>
                  );
                })}
              </div>
            </div>

            <div className="rounded-lg bg-muted/30 p-3">
              <div className="mb-3 text-label font-semibold text-secondary-foreground">
                {draft.id ? "编辑字典项" : "新增字典项"}
              </div>
              <div className="grid gap-3 md:grid-cols-2">
                <Input
                  value={draft.value}
                  onChange={event => setDraft(current => ({ ...current, value: event.target.value }))}
                  placeholder="实际存储值"
                />
                <Input
                  value={draft.label}
                  onChange={event => setDraft(current => ({ ...current, label: event.target.value }))}
                  placeholder="用户显示名称"
                />
                <Input
                  value={draft.description}
                  onChange={event => setDraft(current => ({ ...current, description: event.target.value }))}
                  placeholder="说明（可选）"
                  className="md:col-span-2"
                />
              </div>
              <div className="mt-3 flex gap-2">
                <Button type="button" onClick={() => void saveDraft()}>
                  <AppIcon name="save" size={14} />
                  {draft.id ? "保存修改" : "新增选项"}
                </Button>
                {draft.id ? (
                  <Button type="button" variant="outline" onClick={resetDraft}>取消</Button>
                ) : null}
              </div>
            </div>

            <div className="space-y-2">
              {selected.items.map((item, index) => (
                <div
                  key={item.id}
                  className={`flex flex-wrap items-center gap-3 rounded-lg bg-muted/35 px-3 py-3 ${
                    item.enabled ? "" : "opacity-60"
                  }`}
                >
                  <div className="min-w-0 flex-1">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="font-semibold text-foreground">{item.label}</span>
                      <span className="rounded-md bg-card px-2 py-0.5 text-xs text-muted-foreground">
                        {item.value}
                      </span>
                      <span className={`rounded-md px-2 py-0.5 text-xs ${
                        item.enabled ? "bg-success-soft text-success" : "bg-muted text-muted-foreground"
                      }`}>
                        {item.enabled ? "启用" : "停用"}
                      </span>
                    </div>
                    {item.description ? (
                      <div className="mt-1 text-xs text-muted-foreground">{item.description}</div>
                    ) : null}
                  </div>
                  <div className="flex items-center gap-1">
                    <Button
                      type="button"
                      variant="ghost"
                      size="icon"
                      disabled={index === 0}
                      aria-label="上移"
                      onClick={() => void moveItem(item, -1)}
                    >
                      <AppIcon name="chevronUp" size={14} />
                    </Button>
                    <Button
                      type="button"
                      variant="ghost"
                      size="icon"
                      disabled={index === selected.items.length - 1}
                      aria-label="下移"
                      onClick={() => void moveItem(item, 1)}
                    >
                      <AppIcon name="chevronDown" size={14} />
                    </Button>
                    <Button
                      type="button"
                      variant="outline"
                      size="sm"
                      onClick={() => setDraft({
                        id: item.id,
                        value: item.value,
                        label: item.label,
                        description: item.description ?? "",
                      })}
                    >
                      编辑
                    </Button>
                    <Button type="button" variant="secondary" size="sm" onClick={() => void toggleItem(item)}>
                      {item.enabled ? "停用" : "启用"}
                    </Button>
                    <Button type="button" variant="ghost" size="sm" onClick={() => void deleteItem(item)}>
                      删除
                    </Button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        ) : (
          <div className="rounded-lg bg-muted/40 p-4 text-sm text-secondary-foreground">
            暂无业务字典。
          </div>
        )}
        {error ? (
          <div className="mt-4 rounded-lg bg-destructive-soft px-3 py-2 text-sm text-destructive">
            {error}
          </div>
        ) : null}
      </section>
    </div>
  );
}
