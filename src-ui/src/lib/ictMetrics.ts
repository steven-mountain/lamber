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
    ? new Intl.NumberFormat("zh-CN", { style: "currency", currency: "CNY" }).format(numeric)
    : "--";
};

const formatPercent = (value: unknown) => {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? `${(numeric * 100).toFixed(2)}%` : "--";
};

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
