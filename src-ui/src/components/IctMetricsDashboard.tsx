import React from "react";

export interface IctMetricsLike {
  npv: number;
  npv_rate: number;
  margin_rate: number;
  dynamic_payback: string;
  irr: string;
  it_npv?: number;
  it_npv_rate?: number;
  it_margin_rate?: number;
}

export interface IctMetricCard {
  id: string;
  label: string;
  value: string;
  tone: string;
  accent: boolean;
}

const formatCurrency = (value: unknown) => {
  const numeric = Number(value);
  return Number.isFinite(numeric)
    ? new Intl.NumberFormat('zh-CN', { style: 'currency', currency: 'CNY' }).format(numeric)
    : "--";
};

const formatPercent = (value: unknown) => {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? `${(numeric * 100).toFixed(2)}%` : "--";
};

// Single source of truth for the "实时效益评估结果" metric cards. Used by both the
// calculation dashboard and the document-template flow so the two stay identical.
export function buildIctMetricCards(metrics: any): IctMetricCard[] {
  return [
    { id: "ict-metric-npv", label: "项目净现值 (NPV)", value: formatCurrency(metrics?.npv), tone: "", accent: false },
    { id: "ict-metric-npv-rate", label: "净现值率", value: formatPercent(metrics?.npv_rate), tone: "text-success", accent: false },
    { id: "ict-metric-margin", label: "毛利润率", value: formatPercent(metrics?.margin_rate), tone: "text-success", accent: false },
    { id: "ict-metric-payback", label: "动态回收期 (年)", value: String(metrics?.dynamic_payback ?? "--"), tone: "", accent: false },
    { id: "ict-metric-irr", label: "内部收益率 (IRR)", value: String(metrics?.irr ?? "--"), tone: "", accent: false },
    { id: "ict-metric-it-npv", label: "IT 净现值", value: formatCurrency(metrics?.it_npv ?? 0), tone: "", accent: true },
    { id: "ict-metric-it-npv-rate", label: "IT 净现值率", value: formatPercent(metrics?.it_npv_rate ?? 0), tone: "text-success", accent: true },
    { id: "ict-metric-it-margin", label: "IT 毛利率", value: formatPercent(metrics?.it_margin_rate ?? 0), tone: "text-success", accent: true },
  ];
}

export const IctMetricsCards: React.FC<{ metrics: any; className?: string }> = ({ metrics, className = "" }) => (
  <div className={`grid grid-cols-2 md:grid-cols-3 xl:grid-cols-5 gap-4 ${className}`}>
    {buildIctMetricCards(metrics).map(card => (
      <div
        key={card.id}
        className={`p-4 rounded-lg flex flex-col gap-1 border ${card.accent ? "bg-primary/5 border-primary/30" : "bg-muted border-border"}`}
      >
        <span className="text-xs font-semibold text-secondary-foreground">{card.label}</span>
        <span id={card.id} className={`text-lg font-bold numeric-value ${card.tone}`}>{card.value}</span>
      </div>
    ))}
  </div>
);

export const IctMetricsDashboard: React.FC<{ metrics: IctMetricsLike }> = ({ metrics }) => {
  return (
    <div className="mt-auto pt-6 border-t border-border mt-8">
      <h3 className="text-sm font-bold text-secondary-foreground mb-4">实时效益评估结果</h3>
      <IctMetricsCards metrics={metrics} />
    </div>
  );
};
export default IctMetricsDashboard;
