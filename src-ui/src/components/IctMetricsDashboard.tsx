import React from "react";
import { buildIctMetricCards } from "../lib/ictMetrics";

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
