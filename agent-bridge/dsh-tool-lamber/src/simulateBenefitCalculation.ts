import { defineTool, type InferValue } from '@deepseek-ai/dsh-tools';
import { postBridge } from './bridge.js';
import { trustedSessionId } from './projectScope.js';
import { scenarioParameter } from './readBenefitInputs.js';
import { calculationOutputSchema, renderCalculationResult } from './runBenefitCalculation.js';
const text = { type: 'string', required: true } as const;
const changes = { type: 'array', required: true, items: { type: 'object', additionalProperties: false, properties: { kind: text, subject: text, before: text, after: text } } } as const;
const schema = { type: 'object', additionalProperties: false, properties: { basis: text, notice: text, result: { ...calculationOutputSchema, required: true }, explicitChanges: changes, linkedChanges: changes, taxInclAutoFix: { oneOf: [{ type: 'boolean' }, { type: 'null' }], required: true, description: '本次覆盖使用的桌面自动修正设置；未提供覆盖时为null' } } } as const;
type Change = {kind: string; subject: string; before: string; after: string};
function annualValues(serialized: string): unknown[] | null {
    try {
        const value: unknown = JSON.parse(serialized);
        if (Array.isArray(value)) return value;
    } catch { /* Non-JSON scalar changes retain the original text. */ }
    return null;
}
export function renderSimulationChanges(changes: Change[]): string {
    return changes.map(change => {
        const before = annualValues(change.before), after = annualValues(change.after);
        const heading = `${change.subject}（${change.kind}）`;
        if (change.kind === 'annual_cashflow' && before && after && before.length === after.length) {
            return `${heading}\n年度 | 覆盖前 | 假设覆盖后\n${before.map((value, index) => `第 ${index + 1} 年 | ${String(value)} | ${String(after[index])}`).join('\n')}`;
        }
        return `${heading}\n覆盖前：${change.before}\n假设覆盖后：${change.after}`;
    }).join('\n\n') || '无';
}
export const simulateBenefitCalculation = defineTool({
    name: 'simulate_benefit_calculation',
    description: '在绑定项目的已保存快照副本上做假设试算，复用桌面金额编辑、税额自动修正偏好、CT收入联动成本、清拆分、资金计划同步和现金流重建，再调用原Rust引擎。先read_benefit_inputs确认科目和方案；只传具名覆盖，不可传整套输入或projectId，不落库、不回填、不生成方案。结果必须称“假设试算”，列明显式覆盖和联动变更，不得称已保存值。overrides=[]重放原快照；未知/重复科目、同时覆盖CT联动的一对科目、总额锁定方案、缺少或不自洽资金计划等会明确拒绝，应如实转述并引导桌面核对，不能自己补零或保留旧现金流计算。本系统未计算IRR，不得说该项目无法求解。无论联动变更还是结果现金流，引用年度金额必须逐年完整保留原始小数与末年尾差，不得用省略号、仅末尾小数、“各相同金额”或“×4”代替不同金额；联动不展开时只说明调整方式，不引用不完整数字。用户追问已保存值时重新run_benefit_calculation，不能沿用假设值。',
    parameters: { scenario: scenarioParameter, overrides: { type: 'array', required: true, items: { type: 'object', additionalProperties: false, properties: { subject: { ...text, description: 'read_benefit_inputs返回的科目key' }, inclTax: { ...text, description: '含税金额，无单位十进制字符串，最多两位小数' }, taxRate: { type: 'string', description: '可选税率百分数，如6表示6%；不传沿用快照税率' } } } } },
    output: { schema, render(_args, value) { return [{ type: 'text', text: `【假设试算 · 未保存】\n${value.notice}\n【显式覆盖】\n${renderSimulationChanges(value.explicitChanges)}\n【联动变更】\n${renderSimulationChanges(value.linkedChanges)}\n【假设结果】\n${renderCalculationResult(value.result)}` }]; } },
    timeoutMs: 30000, isConcurrencySafe: () => true,
    async execute(args, exec) { return postBridge<InferValue<typeof schema>>('/lamber-bridge/simulate-benefit-calculation', { ...args, sessionId: trustedSessionId(exec) }, exec.signal); },
});
