import type { ReactNode } from "react";
import AppIcon from "../icons/AppIcon";
import { IctMetricsCards } from "../IctMetricsDashboard";

export interface TemplateLayoutTab<T extends string = string> {
  id: T;
  label: string;
}

export interface TemplateCompletionItem {
  label: string;
  filled: boolean;
}

export interface TemplateCompletion {
  completedCount: number;
  totalCount: number;
  percent: number;
}

export function getTemplateCompletion(items: TemplateCompletionItem[]): TemplateCompletion {
  const totalCount = items.length;
  const completedCount = items.filter(item => item.filled).length;
  return {
    completedCount,
    totalCount,
    percent: totalCount > 0 ? Math.round((completedCount / totalCount) * 100) : 0,
  };
}

interface TemplateDocumentLayoutProps<T extends string> {
  templateName: string;
  title: string;
  tabs: Array<TemplateLayoutTab<T>>;
  activeTab: T;
  onTabChange: (tabId: T) => void;
  completion: TemplateCompletion;
  metrics: any;
  onGenerate: () => void;
  children: ReactNode;
}

export function TemplateDocumentLayout<T extends string>({
  templateName,
  title,
  tabs,
  activeTab,
  onTabChange,
  completion,
  metrics,
  onGenerate,
  children,
}: TemplateDocumentLayoutProps<T>) {
  return (
    <div className="flex flex-col gap-4 pb-6">
      <div className="sticky top-0 z-20 rounded-xl bg-card/95 p-3 shadow-sm ring-1 ring-border/60 backdrop-blur">
        <div className="mb-3 flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
          <div>
            <h4 className="font-bold text-primary">{title}</h4>
            <div className="mt-0.5 text-xs font-semibold text-secondary-foreground">
              {templateName.replace(".docx", "").replace(".xlsx", "")}
            </div>
          </div>
          <TemplateCompletionStatus completion={completion} />
        </div>
        <div className="flex gap-1 overflow-x-auto rounded-lg bg-muted/50 p-1">
          {tabs.map(tab => (
            <button
              key={tab.id}
              type="button"
              onClick={() => onTabChange(tab.id)}
              className={`shrink-0 rounded-md px-3 py-2 text-xs font-bold transition-colors ${
                activeTab === tab.id
                  ? "bg-card text-primary shadow-sm"
                  : "text-secondary-foreground hover:bg-card/70 hover:text-foreground"
              }`}
            >
              {tab.label}
            </button>
          ))}
        </div>
      </div>

      {children}

      <TemplateBottomMetricsBar metrics={metrics} onGenerate={onGenerate} />
    </div>
  );
}

export function TemplateTabSection({ children }: { children: ReactNode }) {
  return (
    <div className="rounded-xl bg-card p-5 shadow-sm ring-1 ring-border/60">
      {children}
    </div>
  );
}

export function TemplateSubmoduleCard({ children, className = "" }: { children: ReactNode; className?: string }) {
  return (
    <div className={`rounded-xl bg-muted/40 p-4 shadow-sm ${className}`}>
      {children}
    </div>
  );
}

export function TemplateCompletionStatus({ completion }: { completion: TemplateCompletion }) {
  return (
    <div className="rounded-lg bg-muted/60 px-3 py-1.5 text-xs font-extrabold text-secondary-foreground numeric-value">
      已填写 {completion.completedCount} / {completion.totalCount} 项&nbsp;&nbsp;{completion.percent}%
    </div>
  );
}

interface TemplateConfirmationPanelProps {
  templateName: string;
  currentSchemeLabel?: string;
  projectName?: string;
  customerName?: string;
  projectYears?: number | string;
  completion: TemplateCompletion;
  onGenerate: () => void;
}

export function TemplateConfirmationPanel({
  templateName,
  currentSchemeLabel,
  projectName,
  customerName,
  projectYears,
  completion,
  onGenerate,
}: TemplateConfirmationPanelProps) {
  return (
    <TemplateTabSection>
      <div className="grid grid-cols-1 gap-4 xl:grid-cols-2">
        <div className="rounded-xl bg-muted/40 p-4 shadow-sm">
          <div className="text-xs font-extrabold text-secondary-foreground">当前模板信息</div>
          <div className="mt-2 text-sm font-bold text-foreground">{templateName.replace(".docx", "").replace(".xlsx", "")}</div>
          <div className="mt-1 text-xs text-secondary-foreground">{templateName.endsWith(".xlsx") ? "Excel 模板" : "Word 模板"}</div>
        </div>
        <div className="rounded-xl bg-muted/40 p-4 shadow-sm">
          <div className="text-xs font-extrabold text-secondary-foreground">当前方案信息</div>
          <div className="mt-2 text-sm font-bold text-foreground">{currentSchemeLabel || projectName || "默认方案"}</div>
          <div className="mt-1 text-xs text-secondary-foreground">
            {customerName || "未填写客户"} · {projectYears || 1} 年
          </div>
        </div>
        <div className="rounded-xl bg-muted/40 p-4 shadow-sm xl:col-span-2">
          <div className="mb-3 flex items-center justify-between gap-3">
            <div className="text-xs font-extrabold text-secondary-foreground">完成度提示</div>
            <div className="text-sm font-extrabold text-primary numeric-value">
              {completion.completedCount} / {completion.totalCount} · {completion.percent}%
            </div>
          </div>
          <div className="h-2 overflow-hidden rounded-full bg-background">
            <div className="h-full rounded-full bg-primary transition-all" style={{ width: `${completion.percent}%` }} />
          </div>
        </div>
        <button
          type="button"
          className="inline-flex items-center justify-center gap-2 bg-primary text-primary-foreground font-bold py-3 px-6 rounded-lg shadow-sm hover:opacity-90 transition-opacity xl:col-span-2"
          onClick={onGenerate}
        >
          <AppIcon name="generate" size={18} /> 立即生成此文件
        </button>
      </div>
    </TemplateTabSection>
  );
}

function TemplateBottomMetricsBar({ metrics, onGenerate }: { metrics: any; onGenerate: () => void }) {
  return (
    <div className="rounded-xl bg-card p-5 shadow-sm ring-1 ring-border/60">
      <div className="mb-4 flex items-center justify-between gap-3">
        <h3 className="text-sm font-bold text-secondary-foreground">实时效益评估结果</h3>
        <button
          type="button"
          className="inline-flex items-center justify-center gap-2 rounded-lg bg-primary px-5 py-2.5 text-sm font-bold text-primary-foreground shadow-sm transition-opacity hover:opacity-90"
          onClick={onGenerate}
        >
          <AppIcon name="generate" size={18} /> 立即生成文件
        </button>
      </div>
      <IctMetricsCards metrics={metrics} />
    </div>
  );
}
