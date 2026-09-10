import { trustedSessionId } from './projectScope.js';
/**
 * `run_benefit_calculation` — replay one lamber project's saved ICT benefit
 * inputs through the Rust calculation engine and return the resulting metrics.
 *
 * The tool owns no math. It resolves a project (and optionally a scheme) on the
 * lamber side and reports what `benefit::calculator::calculate_ict_benefit`
 * produced, so the agent always quotes the same numbers the desktop UI shows.
 */
import { defineTool } from '@deepseek-ai/dsh-tools';
import { postBridge } from './bridge.js';

/** Bridge route backing this tool. */
export const CALCULATE_ROUTE = '/lamber-bridge/calculate';

/**
 * Response contract of `POST /lamber-bridge/calculate`.
 * Mirrors `agent_bridge::CalculateResponse` on the Rust side.
 */
export interface CalculateResponse {
  basis: string;
  irrNotice: string;
  projectId: string;
  projectName: string;
  customerName: string;
  schemeId: string;
  schemeName: string;
  stage: string;
  snapshotVersion: number;
  calculatedAt: string;
  metrics: {
    npv: string;
    npvRate: string;
    marginRate: string;
    dynamicPayback: string;
    irr: string;
    itNpv: string;
    itNpvRate: string;
    itMarginRate: string;
  };
  cashflow: Array<{
    year: number;
    cashIn: string;
    cashOut: string;
    netCash: string;
    cumNetCash: string;
    pv: string;
    cumPv: string;
  }>;
}

/** Money/rate strings are already rounded by the Rust engine; never reformat them. */
const moneyField = (description: string) =>
  ({ type: 'string', required: true, description }) as const;

export const calculationOutputSchema = {
      type: 'object',
      additionalProperties: false,
      properties: {
        basis: moneyField('口径：saved_snapshot，重放已保存快照。'),
        irrNotice: moneyField('本系统未计算 IRR。'),
        projectId: { type: 'string', required: true, description: '项目 id。' },
        projectName: { type: 'string', required: true, description: '项目名称。' },
        customerName: { type: 'string', required: true, description: '客户名称。' },
        schemeId: { type: 'string', required: true, description: '实际使用的测算方案 id。' },
        schemeName: { type: 'string', required: true, description: '实际使用的测算方案名称。' },
        stage: {
          type: 'string',
          required: true,
          description: '方案甄选阶段：`pre_selection` / `post_selection` / `unlabeled`。',
        },
        snapshotVersion: { type: 'integer', required: true, description: '所用输入快照的版本号。' },
        calculatedAt: { type: 'string', required: true, description: '本次测算的执行时间（ISO 8601）。' },
        metrics: {
          type: 'object',
          required: true,
          additionalProperties: false,
          properties: {
            npv: moneyField('项目整体净现值（元）。'),
            npvRate: moneyField('项目整体 NPV 率。'),
            marginRate: moneyField('项目整体利润率。'),
            dynamicPayback: moneyField('动态投资回收期（年），无法回收时为 `--`。'),
            irr: moneyField('本系统未计算 IRR，固定为 `--`，不表示该项目无法求解。'),
            itNpv: moneyField('IT 部分净现值（元）。'),
            itNpvRate: moneyField('IT 部分 NPV 率。'),
            itMarginRate: moneyField('IT 部分利润率。'),
          },
        },
        cashflow: {
          type: 'array',
          required: true,
          description: '逐年现金流。',
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              year: { type: 'integer', required: true, description: '年份序号，从 1 开始。' },
              cashIn: moneyField('当年现金流入（元）。'),
              cashOut: moneyField('当年现金流出（元）。'),
              netCash: moneyField('当年净现金流（元）。'),
              cumNetCash: moneyField('累计净现金流（元）。'),
              pv: moneyField('当年现值（元）。'),
              cumPv: moneyField('累计现值（元）。'),
            },
          },
        },
      },
    } as const;

// Saved and hypothetical outputs use the same per-year rendering; never round or group years.
export function renderCalculationResult(value: CalculateResponse): string {
      const { metrics } = value;
      const header = [
        `项目：${value.projectName}（${value.projectId}）`,
        `客户：${value.customerName || '未填写'}`,
        `方案：${value.schemeName}（${value.schemeId}，阶段 ${value.stage}，快照 v${value.snapshotVersion}）`,
      ].join('\n');
      const summary = [
        `NPV：${metrics.npv}`,
        `NPV 率：${metrics.npvRate}`,
        `利润率：${metrics.marginRate}`,
        `动态回收期：${metrics.dynamicPayback}`,
        `IRR：${metrics.irr}`,
        `IT NPV：${metrics.itNpv}`,
        `IT NPV 率：${metrics.itNpvRate}`,
        `IT 利润率：${metrics.itMarginRate}`,
      ].join('\n');
      const rows = value.cashflow
        .map(
          (row) =>
            `第 ${row.year} 年 | 流入 ${row.cashIn} | 流出 ${row.cashOut} | 净 ${row.netCash} | 累计净 ${row.cumNetCash} | 现值 ${row.pv} | 累计现值 ${row.cumPv}`,
        )
        .join('\n');
  return `${value.irrNotice}\n${header}\n\n【效益指标】\n${summary}\n\n【现金流 · 按年逐分】\n${rows}\n转述规则：引用年度金额时逐年保留原始小数，不得把有尾差的年份合并成“各相同金额”或“×年数”；篇幅有限时可不展开现金流，但不得改写金额。金额引用必须能在本次工具原始字段中逐字找到；不要自行计算差额、均值、尾差大小，也不要在表后追加含新金额的解释；只可说明“第4年与前三年不同，已逐年列示”等不带新增数值的事实。`;
}

export const runBenefitCalculation = defineTool({
  name: 'run_benefit_calculation',
  description: [
    '重新运行 lamber 中某个 ICT 项目的经济效益测算，返回已保存快照的 NPV、NPV 率、利润率、动态回收期；本系统未计算 IRR',
    '以及逐年现金流。计算完全由 lamber 的 Rust 测算引擎执行，数值与桌面端测算表一致。',
    '只读操作，不会修改任何项目数据。',
    '`projectId` 传 lamber 项目 id；`scenario` 可选，用于选择测算方案：',
    '`pre_selection`（甄选前）、`post_selection`（甄选后）、方案 id 或方案名称；',
    '缺省时使用该项目的默认方案。',
  ].join(' '),
  parameters: {
    projectId: {
      type: 'string',
      required: true,
      description: 'lamber 项目 id（不是项目名称）。',
    },
    scenario: {
      type: 'string',
      description:
        '可选的测算方案选择器：`pre_selection`、`post_selection`、方案 id 或方案名称。缺省使用项目默认方案。',
    },
  },
  output: {
    schema: calculationOutputSchema,
    render(_args, value) {
      return [
        {
          type: 'text',
          text: `【已保存快照重算】\n${renderCalculationResult(value)}`,
        },
      ];
    },
  },
  timeoutMs: 30_000,
  isConcurrencySafe: () => true,
  async execute(args, exec) {
    return postBridge<CalculateResponse>(
      CALCULATE_ROUTE,
      { sessionId: trustedSessionId(exec), projectId: args.projectId, scenario: args.scenario ?? null },
      exec.signal,
    );
  },
});
