import React from "react";

interface IctCashflowTableProps {
  state: any;
  calculations: any;
}

const cashflowModelLabels: Record<string, string> = {
  model_a: "模型 A: 100% 在第一年收付",
  model_b: "模型 B: 按周期等额收付 (每年 1/n)",
  model_c: "模型 C: 尾款质保金 (首年95%，末年5%)",
  model_d: "模型 D: 高级自定义分配",
  model_e: "模型 E: 分板块资金计划",
};

const formatDistribution = (arr: number[]) => {
  return "[" + arr.map(v => (v * 100).toFixed(1) + "%").join(", ") + "]";
};

const formatCurrency = (v: number) => new Intl.NumberFormat('zh-CN', { style: 'currency', currency: 'CNY' }).format(v);
const formatCashflowSeries = (values: number[]) => {
  const parts = values
    .map((value, idx) => ({ year: idx + 1, value }))
    .filter(item => Math.abs(item.value) > 0.005)
    .map(item => `第${item.year}年 ${formatCurrency(item.value)}`);

  return parts.length > 0 ? parts.join("，") : "暂无有效金额";
};

export const IctCashflowTable: React.FC<IctCashflowTableProps> = ({ state, calculations }) => {
  const {
    cashflowModel,
    segmentValueMode,
  } = state;

  const {
    cashflowTable,
    directSegmentCashflow,
  } = calculations;

  const effectiveDistRev = state.distRev;
  const effectiveDistCost = state.distCost;

  const hasDirectSegmentCashflow = cashflowModel === "model_e" && segmentValueMode === "amount";

  const distributionPreview = (
    <div className="mt-6 mb-4 border-t border-border pt-4 text-sm">
      <div className="font-bold text-foreground mb-2">资金分布预览</div>
      <div className="grid grid-cols-1 gap-1 text-secondary-foreground">
        <div>当前资金收付模型：{cashflowModelLabels[cashflowModel]}</div>
        <div>收入分布：{formatDistribution(effectiveDistRev)}</div>
        <div>成本分布：{formatDistribution(effectiveDistCost)}</div>
        {hasDirectSegmentCashflow && (
          <>
            <div>收入现金流(不含税)：{formatCashflowSeries(directSegmentCashflow.rev)}</div>
            <div>成本现金流(不含税)：{formatCashflowSeries(directSegmentCashflow.cost)}</div>
          </>
        )}
      </div>
    </div>
  );

  return (
    <div className="bg-card border border-border rounded-xl p-8 shadow-sm">
      <h3 className="text-lg font-bold text-foreground mb-6">1-10年项目现金流推演</h3>
      {distributionPreview}
      <div className="overflow-x-auto">
        <table className="w-full text-sm text-right border-separate border-spacing-y-2">
          <thead>
            <tr className="text-secondary-foreground font-bold text-xs uppercase">
              <th className="text-left pb-2">年份</th>
              <th className="pb-2">现金流入</th>
              <th className="pb-2">现金流出</th>
              <th className="pb-2">净流量</th>
              <th className="pb-2">累计净流量</th>
              <th className="pb-2">净流量现值</th>
              <th className="pb-2">累计现值</th>
            </tr>
          </thead>
          <tbody>
            {cashflowTable.map((row: any, i: number) => (
              <tr key={i} className="bg-muted">
                <td className="text-left p-3 rounded-l-md font-semibold">第{row.year}年</td>
                <td className="p-3 numeric-value">{formatCurrency(row.cash_in)}</td>
                <td className="p-3 numeric-value">{formatCurrency(row.cash_out)}</td>
                <td className="p-3 font-bold text-primary numeric-value">{formatCurrency(row.net_cash)}</td>
                <td className="p-3 numeric-value">{formatCurrency(row.cum_net_cash)}</td>
                <td className="p-3 numeric-value">{formatCurrency(row.pv)}</td>
                <td className="p-3 rounded-r-md numeric-value">{formatCurrency(row.cum_pv)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};
export default IctCashflowTable;
