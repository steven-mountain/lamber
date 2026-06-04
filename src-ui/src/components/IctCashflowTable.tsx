import React from "react";
import { buildAnnualCashflowSubjectContributions } from "../lib/ictSubjectFundingPlan";

interface IctCashflowTableProps {
  state: any;
  calculations: any;
}

const toFiniteNumber = (value: unknown) => {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : 0;
};

const formatCurrency = (value: unknown) =>
  new Intl.NumberFormat('zh-CN', { style: 'currency', currency: 'CNY' }).format(toFiniteNumber(value));
const formatCashflowSeries = (values: number[]) => {
  const parts = values
    .map((value, idx) => ({ year: idx + 1, value }))
    .filter(item => Math.abs(item.value) > 0.005)
    .map(item => `第${item.year}年 ${formatCurrency(item.value)}`);

  return parts.length > 0 ? parts.join("，") : "暂无有效金额";
};

export const IctCashflowTable: React.FC<IctCashflowTableProps> = ({ state, calculations }) => {
  const [expandedYear, setExpandedYear] = React.useState<number | null>(null);
  const [cashflowView, setCashflowView] = React.useState<"overall" | "it">("overall");
  const {
    cashflowTable,
    subjectFundingAnnualCashflow,
    subjectFundingCalculationBlocked,
    subjectFundingCoverage,
  } = calculations;

  const contributions = React.useMemo(() => {
    if (!subjectFundingCoverage) return [];
    return buildAnnualCashflowSubjectContributions(state.subjectFundingPlans, [
      ...subjectFundingCoverage.revenueSubjects,
      ...subjectFundingCoverage.costSubjects,
    ]);
  }, [state.subjectFundingPlans, subjectFundingCoverage]);

  const displayCashflowRows = React.useMemo(() => {
    let cumItNetCash = 0;
    let cumItPv = 0;
    return cashflowTable.map((row: any) => {
      const netItCash = toFiniteNumber(row.net_it_cash);
      const itPv = toFiniteNumber(row.it_pv);
      cumItNetCash += netItCash;
      cumItPv += itPv;
      return {
        ...row,
        cum_it_net_cash: cumItNetCash,
        cum_it_pv: cumItPv,
      };
    });
  }, [cashflowTable]);

  const distributionPreview = (
    <div className="mt-6 mb-4 border-t border-border pt-4 text-sm">
      <div className="font-bold text-foreground mb-2">资金分布预览</div>
      <div className="grid grid-cols-1 gap-1 text-secondary-foreground">
        <div>现金流依据：科目收付款计划</div>
        <div>科目计划现金流入(不含税)：{formatCashflowSeries(subjectFundingAnnualCashflow?.annualRevenueExcl || [])}</div>
        <div>科目计划现金流出(不含税)：{formatCashflowSeries(subjectFundingAnnualCashflow?.annualCostExcl || [])}</div>
        {subjectFundingCalculationBlocked && (
          <div className="text-warning-foreground font-semibold">
            当前覆盖校验未通过，表格保留上一次有效计算结果。
          </div>
        )}
      </div>
    </div>
  );

  return (
    <div className="bg-card border border-border rounded-xl p-8 shadow-sm">
      <div className="flex items-center justify-between mb-6">
        <h3 className="text-lg font-bold text-foreground">1-10年项目现金流推演</h3>
        <div className="flex bg-muted rounded-lg p-1">
          <button
            onClick={() => setCashflowView("overall")}
            className={`px-3 py-1.5 text-xs font-bold rounded-md transition-colors ${
              cashflowView === "overall" ? "bg-background shadow-sm text-foreground" : "text-secondary-foreground hover:text-foreground"
            }`}
          >
            项目整体
          </button>
          <button
            onClick={() => setCashflowView("it")}
            className={`px-3 py-1.5 text-xs font-bold rounded-md transition-colors ${
              cashflowView === "it" ? "bg-background shadow-sm text-foreground" : "text-secondary-foreground hover:text-foreground"
            }`}
          >
            IT 部分
          </button>
        </div>
      </div>
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
            {displayCashflowRows.map((row: any, i: number) => {
              const isExpanded = expandedYear === i;
              const yearContributions = contributions[i] || [];
              const revContribs = yearContributions.filter(c => c.side === "revenue");
              const costContribs = yearContributions.filter(c => c.side === "cost");
              const hasDrillDown = !subjectFundingCalculationBlocked && yearContributions.length > 0;

              return (
                <React.Fragment key={i}>
                  <tr className="bg-muted transition-colors hover:bg-muted/80">
                    <td className="text-left p-3 rounded-l-md font-semibold">
                      第{row.year}年
                    </td>
                    <td className="p-3 numeric-value">
                      {hasDrillDown && revContribs.length > 0 && cashflowView === "overall" ? (
                        <button
                          className="hover:underline text-primary"
                          onClick={() => setExpandedYear(isExpanded ? null : i)}
                          title="点击查看年度收入明细"
                        >
                          {formatCurrency(row.cash_in)}
                        </button>
                      ) : (
                        formatCurrency(cashflowView === "it" ? row.it_cash_in : row.cash_in)
                      )}
                    </td>
                    <td className="p-3 numeric-value">
                      {hasDrillDown && costContribs.length > 0 && cashflowView === "overall" ? (
                        <button
                          className="hover:underline text-primary"
                          onClick={() => setExpandedYear(isExpanded ? null : i)}
                          title="点击查看年度投入明细"
                        >
                          {formatCurrency(row.cash_out)}
                        </button>
                      ) : (
                        formatCurrency(cashflowView === "it" ? row.it_cash_out : row.cash_out)
                      )}
                    </td>
                    <td className="p-3 font-bold text-primary numeric-value">{formatCurrency(cashflowView === "it" ? row.net_it_cash : row.net_cash)}</td>
                    <td className="p-3 numeric-value">{formatCurrency(cashflowView === "it" ? row.cum_it_net_cash : row.cum_net_cash)}</td>
                    <td className="p-3 numeric-value">{formatCurrency(cashflowView === "it" ? row.it_pv : row.pv)}</td>
                    <td className="p-3 rounded-r-md numeric-value">{formatCurrency(cashflowView === "it" ? row.cum_it_pv : row.cum_pv)}</td>
                  </tr>
                  {isExpanded && hasDrillDown && cashflowView === "overall" && (
                    <tr>
                      <td colSpan={7} className="p-3 bg-card border border-border rounded-md shadow-inner text-left">
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          {revContribs.length > 0 && (
                            <div>
                              <div className="font-bold text-xs mb-2 text-foreground">第 {row.year} 年收入收款组成</div>
                              <table className="w-full text-[11px] text-right">
                                <thead>
                                  <tr className="text-secondary-foreground border-b border-border">
                                    <th className="text-left py-1">科目 / 业务名称</th>
                                    <th className="py-1">含税收款</th>
                                    <th className="py-1">不含税现金流</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {revContribs.map((c, idx) => (
                                    <tr key={idx} className="border-b border-border/50">
                                      <td className="text-left py-1 font-semibold text-foreground">{c.subjectDisplayName}</td>
                                      <td className="py-1 text-secondary-foreground numeric-value">{formatCurrency(c.annualInclAmount)}</td>
                                      <td className="py-1 text-foreground numeric-value">{formatCurrency(c.annualExclAmount)}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          )}
                          {costContribs.length > 0 && (
                            <div>
                              <div className="font-bold text-xs mb-2 text-foreground">第 {row.year} 年投入付款组成</div>
                              <table className="w-full text-[11px] text-right">
                                <thead>
                                  <tr className="text-secondary-foreground border-b border-border">
                                    <th className="text-left py-1">科目 / 业务名称</th>
                                    <th className="py-1">含税付款</th>
                                    <th className="py-1">不含税现金流</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {costContribs.map((c, idx) => (
                                    <tr key={idx} className="border-b border-border/50">
                                      <td className="text-left py-1 font-semibold text-foreground">{c.subjectDisplayName}</td>
                                      <td className="py-1 text-secondary-foreground numeric-value">{formatCurrency(c.annualInclAmount)}</td>
                                      <td className="py-1 text-foreground numeric-value">{formatCurrency(c.annualExclAmount)}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          )}
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
};
export default IctCashflowTable;
