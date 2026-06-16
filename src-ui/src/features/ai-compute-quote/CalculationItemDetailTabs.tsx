import { useState } from "react";
import AppIcon from "../../components/icons/AppIcon";
import { Button } from "../../components/ui/button";
import { Input } from "../../components/ui/input";
import { ICT_SUBJECT_DEFINITIONS } from "../../lib/ictSubjectCatalog";
import AiComputeFundingPlanEditor from "./AiComputeFundingPlanEditor";
import QuoteFormulaCalculator from "./QuoteFormulaCalculator";
import type {
  AiComputeQuoteBlueprint,
  AiComputeQuoteExpressionFormula,
  AiComputeQuoteLineItem,
  AiComputeQuoteSubjectMapping,
} from "./types";

type DetailTab = "project" | "formula" | "funding";

type Props = {
  blueprint: AiComputeQuoteBlueprint;
  item: AiComputeQuoteLineItem;
  mapping?: AiComputeQuoteSubjectMapping;
  normalizedFormula: AiComputeQuoteExpressionFormula;
  projectCycleYears: number;
  onUpdateItem: (patch: Partial<AiComputeQuoteLineItem>) => void;
  onUpdateFormula: (formula: AiComputeQuoteExpressionFormula) => void;
  onUpdateMapping: (subjectCode: string) => void;
  onRemove: () => void;
};

const TABS: Array<{ id: DetailTab; label: string; icon: "edit" | "calculator" | "cashflow" }> = [
  { id: "project", label: "项目编辑", icon: "edit" },
  { id: "formula", label: "计算公式", icon: "calculator" },
  { id: "funding", label: "资金计划", icon: "cashflow" },
];

function ProjectEditTab({
  item,
  mapping,
  onUpdateItem,
  onUpdateMapping,
  onRemove,
}: Pick<Props, "item" | "mapping" | "onUpdateItem" | "onUpdateMapping" | "onRemove">) {
  const subjects = ICT_SUBJECT_DEFINITIONS.filter(subject => subject.side === item.side);

  return (
    <div className="grid gap-4">
      <div className="grid gap-3 lg:grid-cols-[minmax(220px,1fr)_140px_minmax(220px,1fr)]">
        <label className="grid gap-1.5 text-caption font-semibold text-secondary-foreground">
          项目名称
          <Input
            value={item.name}
            onChange={event => onUpdateItem({ name: event.target.value })}
          />
        </label>
        <label className="grid gap-1.5 text-caption font-semibold text-secondary-foreground">
          税率（%）
          <Input
            className="numeric-value"
            type="number"
            min="0"
            value={item.taxRate || 0}
            onChange={event => onUpdateItem({ taxRate: Number(event.target.value) || 0 })}
          />
        </label>
        <label className="grid gap-1.5 text-caption font-semibold text-secondary-foreground">
          ICT 同步科目
          <select
            className="h-[var(--density-input-height)] min-w-0 rounded-md bg-card px-3 text-sm shadow-sm ring-1 ring-input/70 outline-none focus:ring-2 focus:ring-ring/30"
            value={mapping?.ictSubjectCode || ""}
            onChange={event => onUpdateMapping(event.target.value)}
          >
            <option value="">未映射</option>
            {subjects.map(subject => (
              <option key={subject.subjectCode} value={subject.subjectCode}>
                {subject.standardSubjectName}
              </option>
            ))}
          </select>
        </label>
      </div>

      <div className="grid gap-2 md:grid-cols-2">
        <label className="flex items-start gap-2.5 rounded-lg bg-muted/55 px-3 py-3 text-sm font-semibold text-foreground">
          <input
            className="mt-0.5"
            type="checkbox"
            checked={item.enabled}
            onChange={event => onUpdateItem({ enabled: event.target.checked })}
          />
          <span>
            是否启用 / 参与测算
            <span className="mt-0.5 block text-caption font-normal text-secondary-foreground">
              关闭后该项目按 0 参与公式汇总。
            </span>
          </span>
        </label>
        <label className="flex items-start gap-2.5 rounded-lg bg-muted/55 px-3 py-3 text-sm font-semibold text-foreground">
          <input
            className="mt-0.5"
            type="checkbox"
            checked={item.outputEnabled !== false}
            onChange={event => onUpdateItem({ outputEnabled: event.target.checked })}
          />
          <span>
            是否参与 ICT 同步输出
            <span className="mt-0.5 block text-caption font-normal text-secondary-foreground">
              关闭后保留业务计算，但不写入映射科目。
            </span>
          </span>
        </label>
      </div>

      <div className="flex justify-end">
        <Button variant="ghost" size="sm" className="text-destructive" onClick={onRemove}>
          <AppIcon name="delete" />
          删除当前项目
        </Button>
      </div>
    </div>
  );
}

export default function CalculationItemDetailTabs({
  blueprint,
  item,
  mapping,
  normalizedFormula,
  projectCycleYears,
  onUpdateItem,
  onUpdateFormula,
  onUpdateMapping,
  onRemove,
}: Props) {
  const [activeTab, setActiveTab] = useState<DetailTab>("project");

  return (
    <div className="overflow-hidden rounded-lg bg-card shadow-sm">
      <div className="flex gap-1 overflow-x-auto bg-muted/45 p-1.5" role="tablist" aria-label={`${item.name} 详情编辑`}>
        {TABS.map(tab => (
          <button
            key={tab.id}
            type="button"
            role="tab"
            aria-selected={activeTab === tab.id}
            onClick={() => setActiveTab(tab.id)}
            className={`flex min-w-max flex-1 items-center justify-center gap-2 rounded-md px-3 py-2 text-sm font-bold transition-colors ${
              activeTab === tab.id
                ? "bg-card text-primary shadow-sm"
                : "text-secondary-foreground hover:bg-card/60 hover:text-foreground"
            }`}
          >
            <AppIcon name={tab.icon} size={16} />
            {tab.label}
          </button>
        ))}
      </div>

      <div className="p-4">
        {activeTab === "project" && (
          <ProjectEditTab
            item={item}
            mapping={mapping}
            onUpdateItem={onUpdateItem}
            onUpdateMapping={onUpdateMapping}
            onRemove={onRemove}
          />
        )}

        {activeTab === "formula" && (
          <div>
            <QuoteFormulaCalculator
              blueprint={blueprint}
              item={{ ...item, formula: normalizedFormula }}
              onChange={onUpdateFormula}
            />
          </div>
        )}

        {activeTab === "funding" && (
          <AiComputeFundingPlanEditor
            item={item}
            projectCycleYears={projectCycleYears}
            onChange={fundingPlan => onUpdateItem({ fundingPlan })}
          />
        )}
      </div>
    </div>
  );
}
