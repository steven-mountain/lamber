import React, { useMemo } from "react";
import AppIcon from "./icons/AppIcon";
import {
  type CashflowSegment,
  type SegmentValueMode,
  type SegmentFlowMode,
  type SegmentSideScope,
  clampCashflowYear,
  getSegmentEffectiveDurationValue,
  normalizeProjectYears
} from "../hooks/useIctState";

interface IctBasicInfoProps {
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

const segmentFlowModeLabels: Record<SegmentFlowMode, string> = {
  upfront: "第一年一次性",
  equal: "按年/月等额",
  custom: "自定义年度计划",
};

const segmentRevenueScopeLabels: Record<Exclude<SegmentSideScope, "mix">, string> = {
  project: "项目整体",
  it: "IT收入",
  ct: "CT收入",
  non_it_ct: "非IT/CT收入",
};

const segmentCostScopeLabels: Record<SegmentSideScope, string> = {
  project: "项目整体",
  it: "IT支出",
  ct: "CT支出",
  non_it_ct: "非IT/CT支出",
  mix: "综合成本",
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

export const IctBasicInfo: React.FC<IctBasicInfoProps> = ({ state, calculations }) => {
  const {
    projName, setProjName,
    customerName, setCustomerName,
    propertyRights, setPropertyRights,
    discountRate, setDiscountRate,
    projectYears, setProjectYears,
    cashflowModel, setCashflowModel,
    distRev, setDistRev,
    distCost, setDistCost,
    segmentValueMode, setSegmentValueMode,
    cashflowSegments,
    projectBackground, setProjectBackground,
    addCashflowSegment,
    updateCashflowSegment,
    updateCashflowSegmentAnnualValue,
    removeCashflowSegment,
  } = state;

  const {
    directSegmentCashflow,
    revenueInclTotal,
    costInclTotal,
    segmentRevenueInclTotal,
    segmentCostInclTotal,
  } = calculations;

  const activeDistributionYears = useMemo(() => normalizeProjectYears(projectYears), [projectYears]);

  const effectiveDistRev = state.distRev;
  const effectiveDistCost = state.distCost;

  const hasDirectSegmentCashflow = cashflowModel === "model_e" && segmentValueMode === "amount";
  const revenueSegmentMismatch = hasDirectSegmentCashflow && revenueInclTotal > 0 && segmentRevenueInclTotal > 0 && Math.abs(revenueInclTotal - segmentRevenueInclTotal) > 0.01;
  const costSegmentMismatch = hasDirectSegmentCashflow && costInclTotal > 0 && segmentCostInclTotal > 0 && Math.abs(costInclTotal - segmentCostInclTotal) > 0.01;

  const getSegmentAnnualValues = (segment: CashflowSegment, side: "revenue" | "cost") => {
    const values = side === "revenue" ? segment.revenueAnnualValues : segment.costAnnualValues;
    return Array.from({ length: getSegmentEffectiveDurationValue(segment) }, (_, idx) => Number(values?.[idx] || 0));
  };

  const getSegmentAnnualValueSum = (segment: CashflowSegment, side: "revenue" | "cost") => {
    return getSegmentAnnualValues(segment, side).reduce((sum, value) => sum + value, 0);
  };

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

  const formatYearRange = (start: number, end: number) => start + 1 === end
    ? `第 ${start + 1} 年`
    : `第 ${start + 1}-${end} 年`;

  const distributionSegments = activeDistributionYears <= 5
    ? [{ label: formatYearRange(0, activeDistributionYears), start: 0, end: activeDistributionYears }]
    : [
      { label: formatYearRange(0, 5), start: 0, end: 5 },
      { label: formatYearRange(5, activeDistributionYears), start: 5, end: activeDistributionYears },
    ];

  const segmentGridColumns = segmentValueMode === "amount"
    ? "1.2fr 1fr 0.75fr 1fr 0.75fr 0.9fr 0.9fr 0.7fr 0.8fr 1.05fr 1.05fr 48px"
    : "1.2fr 1fr 0.8fr 0.9fr 1.1fr 1.1fr 48px";

  return (
    <div className="bg-card border border-border rounded-xl p-8 shadow-sm">
      <h3 className="text-lg font-bold text-foreground mb-6">项目概况</h3>
      <div className="grid grid-cols-2 gap-6">
        <div className="flex flex-col gap-2">
          <label className="text-sm font-bold text-secondary-foreground">项目名称</label>
          <input id="ict-proj-name" type="text" value={projName} onChange={e => setProjName(e.target.value)} className="bg-card border border-input px-3.5 py-2.5 rounded-md outline-none focus:border-ring" />
        </div>
        <div className="flex flex-col gap-2">
          <label className="text-sm font-bold text-secondary-foreground">客户单位名称</label>
          <input id="ict-customer-name" type="text" value={customerName} onChange={e => setCustomerName(e.target.value)} className="bg-card border border-input px-3.5 py-2.5 rounded-md outline-none focus:border-ring" />
        </div>
        <div className="flex flex-col gap-2">
          <label className="text-sm font-bold text-secondary-foreground">产权归属</label>
          <input id="ict-property-rights" type="text" value={propertyRights} onChange={e => setPropertyRights(e.target.value)} className="bg-card border border-input px-3.5 py-2.5 rounded-md outline-none focus:border-ring" />
        </div>
        <div className="flex flex-col gap-2">
          <label className="text-sm font-bold text-secondary-foreground">项目建设/服务周期 (年)</label>
          <input id="ict-project-years" type="number" min={1} max={10} value={projectYears} onChange={e => setProjectYears(Number(e.target.value))} className="bg-card border border-input px-3.5 py-2.5 rounded-md outline-none focus:border-ring" />
        </div>
        <div className="flex flex-col gap-2">
          <label className="text-sm font-bold text-secondary-foreground">折现率</label>
          <input id="ict-discount-rate" type="number" step={0.001} value={discountRate} onChange={e => setDiscountRate(Number(e.target.value))} className="bg-card border border-input px-3.5 py-2.5 rounded-md outline-none focus:border-ring" />
        </div>
        <div className="flex flex-col gap-2">
          <label className="text-sm font-bold text-secondary-foreground">资金收付模型</label>
          <select id="ict-cashflow-model" value={cashflowModel} onChange={e => setCashflowModel(e.target.value as any)} className="bg-card border border-input px-3.5 py-2.5 rounded-md outline-none focus:border-ring">
            <option value="model_a">模型 A: 100% 在第一年收付</option>
            <option value="model_b">模型 B: 按周期等额收付 (每年 1/n)</option>
            <option value="model_c">模型 C: 尾款质保金 (首年95%，末年5%)</option>
            <option value="model_d">模型 D: 高级自定义分配</option>
            <option value="model_e">模型 E: 分板块资金计划</option>
          </select>
        </div>
        <div className="flex flex-col gap-2 col-span-2">
          <label className="text-sm font-bold text-secondary-foreground">项目背景</label>
          <textarea id="ict-project-bg" rows={3} value={projectBackground} onChange={e => setProjectBackground(e.target.value)} className="bg-card border border-input px-3.5 py-2.5 rounded-md outline-none focus:border-ring" />
        </div>
      </div>

      {cashflowModel === 'model_d' && (
        <div className="mt-6 pt-6 border-t border-border">
          <h4 className="text-sm font-bold text-secondary-foreground mb-4">高级自定义分配 (可输入任意非负比例，系统自动归一化)</h4>
          <p className="text-xs text-secondary-foreground mb-4">
            仅展示项目周期内的 {activeDistributionYears} 个年份；周期外年份按 0 处理，不参与现金流分摊。
          </p>
          <div className="grid grid-cols-1 xl:grid-cols-2 gap-5">
            {distributionSegments.map(segment => {
              const segmentYears = Array.from({ length: segment.end - segment.start }, (_, idx) => segment.start + idx);

              return (
                <div key={segment.label} className="min-w-0 rounded-lg border border-border bg-muted/30 p-3">
                  <div className="text-xs font-bold text-secondary-foreground mb-3">{segment.label}</div>
                  <div
                    className="grid gap-2 items-center text-center text-sm"
                    style={{ gridTemplateColumns: `72px repeat(${segmentYears.length}, minmax(0, 1fr))` }}
                  >
                    <div className="font-bold text-right pr-2 text-secondary-foreground">年份</div>
                    {segmentYears.map(i => (
                      <div key={`year-${i}`} className="font-bold">{i + 1}</div>
                    ))}
                    <div className="font-bold text-right pr-2 text-secondary-foreground">收入比例</div>
                    {segmentYears.map(i => (
                      <input key={`rev-${i}`} type="number" step="0.01" value={distRev[i]} onChange={e => { const newArr = [...distRev]; newArr[i] = Number(e.target.value); setDistRev(newArr); }} className="min-w-0 w-full bg-card border border-input rounded px-1 py-1 outline-none focus:border-ring text-center" />
                    ))}
                    <div className="font-bold text-right pr-2 text-secondary-foreground">支出比例</div>
                    {segmentYears.map(i => (
                      <input key={`cost-${i}`} type="number" step="0.01" value={distCost[i]} onChange={e => { const newArr = [...distCost]; newArr[i] = Number(e.target.value); setDistCost(newArr); }} className="min-w-0 w-full bg-card border border-input rounded px-1 py-1 outline-none focus:border-ring text-center" />
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {cashflowModel === 'model_e' && (
        <div className="mt-6 pt-6 border-t border-border">
          <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
            <div>
              <h4 className="text-sm font-bold text-secondary-foreground">分板块资金计划</h4>
              <p className="text-xs text-secondary-foreground mt-1">
                {segmentValueMode === "ratio"
                  ? "系统会按板块权重汇总 10 年收入和成本分布；支持一次性、等额和自定义年度计划。"
                  : "系统会按板块含税金额、税率 and 年度计划逐年价税分离，生成精确的 10 年不含税现金流。"}
              </p>
            </div>
            <div className="flex rounded-lg border border-border bg-muted p-1">
              {([
                { key: "ratio", label: "填比例" },
                { key: "amount", label: "填金额" },
              ] as Array<{ key: SegmentValueMode; label: string }>).map(option => (
                <button
                  key={option.key}
                  type="button"
                  onClick={() => setSegmentValueMode(option.key)}
                  className={`px-4 py-2 text-sm font-semibold rounded-md transition-colors ${segmentValueMode === option.key ? 'bg-primary text-primary-foreground shadow-sm' : 'text-secondary-foreground hover:bg-background'}`}
                >
                  {option.label}
                </button>
              ))}
            </div>
          </div>

          <div className="overflow-x-auto">
            <div className={segmentValueMode === "amount" ? "min-w-[1360px]" : "min-w-[920px]"}>
              <div className="grid gap-2 px-2 pb-2 text-xs font-bold text-secondary-foreground" style={{ gridTemplateColumns: segmentGridColumns }}>
                <div>板块名称</div>
                {segmentValueMode === "ratio" ? (
                  <div>板块占比(%)</div>
                ) : (
                  <>
                    <div>收入金额(含税)</div>
                    <div>收入税率(%)</div>
                    <div>支出金额(含税)</div>
                    <div>支出税率(%)</div>
                    <div>收入归属</div>
                    <div>支出归属</div>
                  </>
                )}
                <div>开始年份</div>
                <div>服务周期</div>
                <div>收款方式</div>
                <div>付款方式</div>
                <div />
              </div>
              <div className="flex flex-col gap-2">
                {cashflowSegments.map((segment: CashflowSegment) => {
                  const customSides = [
                    ...(segment.revenueMode === "custom" ? [{ side: "revenue" as const, label: "收款年度计划" }] : []),
                    ...(segment.costMode === "custom" ? [{ side: "cost" as const, label: "付款年度计划" }] : []),
                  ];

                  return (
                    <div key={segment.id} className="rounded-lg border border-border bg-muted/25 p-2">
                      <div className="grid gap-2 items-center" style={{ gridTemplateColumns: segmentGridColumns }}>
                        <input
                          type="text"
                          value={segment.name}
                          onChange={e => updateCashflowSegment(segment.id, "name", e.target.value)}
                          className="min-w-0 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm focus:border-ring"
                        />
                        {segmentValueMode === "ratio" ? (
                          <input
                            type="number"
                            min={0}
                            step={1}
                            value={segment.value === 0 ? "" : segment.value}
                            onChange={e => updateCashflowSegment(segment.id, "value", Number(e.target.value))}
                            className="min-w-0 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm focus:border-ring"
                          />
                        ) : (
                          <>
                            <input
                              type="number"
                              min={0}
                              step={0.01}
                              value={segment.revenueValue === 0 ? "" : segment.revenueValue}
                              onChange={e => updateCashflowSegment(segment.id, "revenueValue", Number(e.target.value))}
                              className="min-w-0 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm focus:border-ring"
                            />
                            <input
                              type="number"
                              min={0}
                              step={0.01}
                              value={segment.revenueTax}
                              onChange={e => updateCashflowSegment(segment.id, "revenueTax", Number(e.target.value))}
                              className="min-w-0 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm focus:border-ring"
                            />
                            <input
                              type="number"
                              min={0}
                              step={0.01}
                              value={segment.costValue === 0 ? "" : segment.costValue}
                              onChange={e => updateCashflowSegment(segment.id, "costValue", Number(e.target.value))}
                              className="min-w-0 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm focus:border-ring"
                            />
                            <input
                              type="number"
                              min={0}
                              step={0.01}
                              value={segment.costTax}
                              onChange={e => updateCashflowSegment(segment.id, "costTax", Number(e.target.value))}
                              className="min-w-0 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm focus:border-ring"
                            />
                            <select
                              value={segment.revenueScope}
                              onChange={e => updateCashflowSegment(segment.id, "revenueScope", e.target.value as any)}
                              className="min-w-0 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm focus:border-ring"
                            >
                              {Object.entries(segmentRevenueScopeLabels).map(([value, label]) => (
                                <option key={value} value={value}>{label}</option>
                              ))}
                            </select>
                            <select
                              value={segment.costScope}
                              onChange={e => updateCashflowSegment(segment.id, "costScope", e.target.value as any)}
                              className="min-w-0 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm focus:border-ring"
                            >
                              {Object.entries(segmentCostScopeLabels).map(([value, label]) => (
                                <option key={value} value={value}>{label}</option>
                              ))}
                            </select>
                          </>
                        )}
                        <input
                          type="number"
                          min={1}
                          max={10}
                          value={segment.startYear}
                          onChange={e => updateCashflowSegment(segment.id, "startYear", clampCashflowYear(Number(e.target.value)))}
                          className="min-w-0 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm focus:border-ring"
                        />
                        <input
                          type="number"
                          min={1}
                          max={10}
                          value={segment.serviceYears}
                          onChange={e => updateCashflowSegment(segment.id, "serviceYears", clampCashflowYear(Number(e.target.value)))}
                          className="min-w-0 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm focus:border-ring"
                        />
                        <select
                          value={segment.revenueMode}
                          onChange={e => updateCashflowSegment(segment.id, "revenueMode", e.target.value as any)}
                          className="min-w-0 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm focus:border-ring"
                        >
                          {Object.entries(segmentFlowModeLabels).map(([value, label]) => (
                            <option key={value} value={value}>{label}</option>
                          ))}
                        </select>
                        <select
                          value={segment.costMode}
                          onChange={e => updateCashflowSegment(segment.id, "costMode", e.target.value as any)}
                          className="min-w-0 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm focus:border-ring"
                        >
                          {Object.entries(segmentFlowModeLabels).map(([value, label]) => (
                            <option key={value} value={value}>{label}</option>
                          ))}
                        </select>
                        <button
                          type="button"
                          onClick={() => removeCashflowSegment(segment.id)}
                          disabled={cashflowSegments.length <= 1}
                          className="flex h-10 w-10 items-center justify-center rounded-md text-destructive transition-colors hover:bg-destructive/10 disabled:cursor-not-allowed disabled:opacity-40"
                          title="删除板块"
                        >
                          <AppIcon name="delete" size={16} />
                        </button>
                      </div>

                      {customSides.length > 0 && (
                        <div className="mt-3 grid grid-cols-1 gap-3 border-t border-border pt-3">
                          {customSides.map(({ side, label }) => {
                            const annualValues = getSegmentAnnualValues(segment, side);
                            const annualSum = getSegmentAnnualValueSum(segment, side);
                            const sideAmount = side === "revenue" ? segment.revenueValue : segment.costValue;
                            const hasMismatch = segmentValueMode === "amount" && sideAmount > 0 && annualSum > 0 && Math.abs(annualSum - sideAmount) > 0.01;

                            return (
                              <div key={`${segment.id}-${side}`} className="rounded-md bg-background p-3">
                                <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                                  <span className="text-xs font-bold text-secondary-foreground">{label}</span>
                                  <span className={`text-[11px] font-semibold ${hasMismatch ? 'text-amber-700' : 'text-secondary-foreground'}`}>
                                    合计：{annualSum.toFixed(segmentValueMode === "amount" ? 2 : 0)}{segmentValueMode === "amount" ? " 元" : ""}
                                  </span>
                                </div>
                                <div className="grid gap-2" style={{ gridTemplateColumns: `repeat(${annualValues.length}, minmax(96px, 1fr))` }}>
                                  {annualValues.map((value, idx) => (
                                    <div key={`${segment.id}-${side}-${idx}`} className="flex flex-col gap-1">
                                      <label className="text-[11px] font-semibold text-secondary-foreground">
                                        第 {clampCashflowYear(segment.startYear) + idx} 年
                                      </label>
                                      <input
                                        type="number"
                                        min={0}
                                        step={segmentValueMode === "amount" ? 0.01 : 1}
                                        value={value === 0 ? "" : value}
                                        placeholder={segmentValueMode === "amount" ? "金额" : "比例"}
                                        onChange={e => updateCashflowSegmentAnnualValue(segment.id, side, idx, Number(e.target.value))}
                                        className="min-w-0 bg-card border border-input px-3 py-2 rounded-md outline-none text-sm focus:border-ring"
                                      />
                                    </div>
                                  ))}
                                </div>
                                {hasMismatch && (
                                  <div className="mt-2 rounded-md border border-amber-200 bg-amber-50 px-3 py-2 text-xs leading-5 text-amber-700">
                                    年度计划合计与该侧含税金额不一致，系统将按你填写的年度金额直接计算现金流。
                                  </div>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          {hasDirectSegmentCashflow && (revenueSegmentMismatch || costSegmentMismatch) && (
            <div className="mt-3 rounded-md border border-amber-200 bg-amber-50 px-3 py-2 text-xs leading-5 text-amber-800">
              {revenueSegmentMismatch && (
                <div>收入板块含税合计 {formatCurrency(segmentRevenueInclTotal)} 与收入侧明细含税合计 {formatCurrency(revenueInclTotal)} 不一致。</div>
              )}
              {costSegmentMismatch && (
                <div>支出板块含税合计 {formatCurrency(segmentCostInclTotal)} 与支出侧明细含税合计 {formatCurrency(costInclTotal)} 不一致。</div>
              )}
            </div>
          )}

          <div className="mt-3 flex items-center justify-between gap-3">
            <p className="text-xs text-secondary-foreground">
              {segmentValueMode === "ratio"
                ? "比例合计不要求刚好等于 100%，系统会自动按合计归一化。"
                : "金额模式会按各板块收入/支出的含税金额、税率和年度计划生成精确的不含税现金流，不会反写收入/成本明细。"}
            </p>
            <button
              type="button"
              onClick={addCashflowSegment}
              className="inline-flex items-center gap-2 rounded-md bg-primary/10 px-3 py-2 text-xs font-bold text-primary transition-colors hover:bg-blue-50"
            >
              + 新增板块
            </button>
          </div>
        </div>
      )}

      {distributionPreview}
    </div>
  );
};
