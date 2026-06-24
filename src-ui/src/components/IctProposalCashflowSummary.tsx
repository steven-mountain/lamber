import React from "react";
import { buildAnnualCashflowSubjectContributions } from "../lib/ictSubjectFundingPlan";

interface IctProposalCashflowSummaryProps {
  state: any;
  calculations: any;
}

const toFiniteNumber = (value: unknown) => {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : 0;
};

const formatCurrency = (value: unknown) =>
  new Intl.NumberFormat('zh-CN', {
    style: 'decimal',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(toFiniteNumber(value));

interface SummaryRow {
  year: number;
  revExcl: number;
  revIncl: number;
  costExcl: number;
  costIncl: number;
}

const hasAmount = (row: SummaryRow) =>
  Math.abs(row.revExcl) > 0.005 ||
  Math.abs(row.revIncl) > 0.005 ||
  Math.abs(row.costExcl) > 0.005 ||
  Math.abs(row.costIncl) > 0.005;

const buildTotals = (rows: SummaryRow[]) =>
  rows.reduce(
    (acc, row) => ({
      revExcl: acc.revExcl + row.revExcl,
      revIncl: acc.revIncl + row.revIncl,
      costExcl: acc.costExcl + row.costExcl,
      costIncl: acc.costIncl + row.costIncl,
    }),
    { revExcl: 0, revIncl: 0, costExcl: 0, costIncl: 0 }
  );

const SummaryTable: React.FC<{
  title: string;
  rows: SummaryRow[];
  inclAvailable: boolean;
}> = ({ title, rows, inclAvailable }) => {
  const totals = buildTotals(rows);
  const inclCell = (value: number) => (inclAvailable ? formatCurrency(value) : "—");

  return (
    <div>
      <div className="font-bold text-foreground mb-2 text-sm">{title}</div>
      <div className="overflow-x-auto">
        <table className="w-full text-sm text-right border-separate border-spacing-y-2">
          <thead>
            <tr className="text-secondary-foreground font-bold text-xs uppercase">
              <th className="text-left pb-2">年限</th>
              <th className="pb-2">收入（不含税/元）</th>
              <th className="pb-2">收入（含税/元）</th>
              <th className="pb-2">支出（不含税/元）</th>
              <th className="pb-2">支出（含税/元）</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={row.year} className="bg-muted transition-colors hover:bg-muted/80">
                <td className="text-left p-3 rounded-l-md font-semibold">第{row.year}年</td>
                <td className="p-3 numeric-value">{formatCurrency(row.revExcl)}</td>
                <td className="p-3 numeric-value">{inclCell(row.revIncl)}</td>
                <td className="p-3 numeric-value">{formatCurrency(row.costExcl)}</td>
                <td className="p-3 rounded-r-md numeric-value">{inclCell(row.costIncl)}</td>
              </tr>
            ))}
            <tr className="bg-primary/5 font-bold text-foreground">
              <td className="text-left p-3 rounded-l-md">总计</td>
              <td className="p-3 numeric-value">{formatCurrency(totals.revExcl)}</td>
              <td className="p-3 numeric-value">{inclCell(totals.revIncl)}</td>
              <td className="p-3 numeric-value">{formatCurrency(totals.costExcl)}</td>
              <td className="p-3 rounded-r-md numeric-value">{inclCell(totals.costIncl)}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
};

export const IctProposalCashflowSummary: React.FC<IctProposalCashflowSummaryProps> = ({ state, calculations }) => {
  const {
    cashflowTable,
    subjectFundingCalculationBlocked,
    subjectFundingCoverage,
  } = calculations;

  // 含税年度金额按科目计划聚合（与现金流明细下钻同源）。
  const contributions = React.useMemo(() => {
    if (!subjectFundingCoverage) return [];
    return buildAnnualCashflowSubjectContributions(state.subjectFundingPlans, [
      ...subjectFundingCoverage.revenueSubjects,
      ...subjectFundingCoverage.costSubjects,
    ]);
  }, [state.subjectFundingPlans, subjectFundingCoverage]);

  const inclAvailable = !subjectFundingCalculationBlocked && contributions.length > 0;

  const { overallRows, itRows } = React.useMemo(() => {
    const overall: SummaryRow[] = [];
    const it: SummaryRow[] = [];

    const rows: any[] = cashflowTable || [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i] ?? {};
      const yearContribs = contributions[i] || [];

      let overallRevIncl = 0;
      let overallCostIncl = 0;
      let itRevIncl = 0;
      let itCostIncl = 0;
      for (const c of yearContribs) {
        const isIt = c.subjectRef?.groupId === "revIt" || c.subjectRef?.groupId === "costIt";
        if (c.side === "revenue") {
          overallRevIncl += c.annualInclAmount;
          if (isIt) itRevIncl += c.annualInclAmount;
        } else {
          overallCostIncl += c.annualInclAmount;
          if (isIt) itCostIncl += c.annualInclAmount;
        }
      }

      const overallRow: SummaryRow = {
        year: i + 1,
        revExcl: toFiniteNumber(row.cash_in),
        revIncl: overallRevIncl,
        costExcl: toFiniteNumber(row.cash_out),
        costIncl: overallCostIncl,
      };
      const itRow: SummaryRow = {
        year: i + 1,
        revExcl: toFiniteNumber(row.it_cash_in),
        revIncl: itRevIncl,
        costExcl: toFiniteNumber(row.it_cash_out),
        costIncl: itCostIncl,
      };

      // 只保留有金额的年份（任一口径非零），与立项材料表一致，跳过空年份。
      if (hasAmount(overallRow)) overall.push(overallRow);
      if (hasAmount(itRow)) it.push(itRow);
    }

    return { overallRows: overall, itRows: it };
  }, [cashflowTable, contributions]);

  return (
    <div className="bg-card border border-border rounded-xl p-8 shadow-sm">
      <div className="flex items-center justify-between mb-2">
        <h3 className="text-lg font-bold text-foreground">立项材料现金流汇总（含税 / 不含税）</h3>
      </div>
      <p className="text-xs text-secondary-foreground mb-6">
        按年汇总项目整体与 IT 部分的收入、支出金额，含税与不含税并列，便于直接填写立项材料。不含税金额与上方现金流推演一致；含税金额按各科目收付款计划税率还原。
      </p>

      {!inclAvailable && (
        <div className="mb-4 text-xs text-warning-foreground font-semibold">
          部分科目尚未维护收付款计划，含税金额暂无法按年拆分（显示为「—」）。请补全科目收付款计划后查看含税分布。
        </div>
      )}

      <div className="grid grid-cols-1 gap-8">
        <SummaryTable title="项目整体" rows={overallRows} inclAvailable={inclAvailable} />
        <SummaryTable title="IT 部分" rows={itRows} inclAvailable={inclAvailable} />
      </div>
    </div>
  );
};

export default IctProposalCashflowSummary;
