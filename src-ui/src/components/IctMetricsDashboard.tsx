import React from "react";

interface IctMetricsDashboardProps {
  metrics: {
    npv: number;
    npv_rate: number;
    margin_rate: number;
    dynamic_payback: string;
    irr: string;
    it_npv?: number;
    it_npv_rate?: number;
    it_margin_rate?: number;
  };
}

const formatCurrency = (v: number) => new Intl.NumberFormat('zh-CN', { style: 'currency', currency: 'CNY' }).format(v);
const formatPercent = (v: number) => (v * 100).toFixed(2) + "%";

export const IctMetricsDashboard: React.FC<IctMetricsDashboardProps> = ({ metrics }) => {
  return (
    <div className="mt-auto pt-6 border-t border-border mt-8">
      <h3 className="text-sm font-bold text-secondary-foreground mb-4">实时效益评估结果</h3>
      <div className="grid grid-cols-5 gap-4">
        <div className="bg-muted p-4 rounded-lg flex flex-col gap-1 border border-border">
          <span className="text-xs font-semibold text-secondary-foreground">项目净现值 (NPV)</span>
          <span id="ict-metric-npv" className="text-lg font-bold numeric-value">{formatCurrency(metrics.npv)}</span>
        </div>
        <div className="bg-muted p-4 rounded-lg flex flex-col gap-1 border border-border">
          <span className="text-xs font-semibold text-secondary-foreground">净现值率</span>
          <span id="ict-metric-npv-rate" className="text-lg font-bold text-success numeric-value">{formatPercent(metrics.npv_rate)}</span>
        </div>
        <div className="bg-muted p-4 rounded-lg flex flex-col gap-1 border border-border">
          <span className="text-xs font-semibold text-secondary-foreground">毛利润率</span>
          <span id="ict-metric-margin" className="text-lg font-bold text-success numeric-value">{formatPercent(metrics.margin_rate)}</span>
        </div>
        <div className="bg-muted p-4 rounded-lg flex flex-col gap-1 border border-border">
          <span className="text-xs font-semibold text-secondary-foreground">动态回收期 (年)</span>
          <span id="ict-metric-payback" className="text-lg font-bold numeric-value">{metrics.dynamic_payback}</span>
        </div>
        <div className="bg-muted p-4 rounded-lg flex flex-col gap-1 border border-border">
          <span className="text-xs font-semibold text-secondary-foreground">内部收益率 (IRR)</span>
          <span id="ict-metric-irr" className="text-lg font-bold numeric-value">{metrics.irr}</span>
        </div>
      </div>
    </div>
  );
};
export default IctMetricsDashboard;
