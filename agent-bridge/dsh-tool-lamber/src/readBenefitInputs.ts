import { defineTool, type InferValue } from '@deepseek-ai/dsh-tools';
import { postBridge } from './bridge.js';
import { trustedSessionId } from './projectScope.js';
const text = { type: 'string', required: true } as const;
export const scenarioParameter = { type: 'string', description: 'pre_selection（甄选前）、post_selection（甄选后）、方案 id 或名称；省略使用默认方案。无法匹配时报错。' } as const;
const schema = { type: 'object', additionalProperties: false, properties: {
        basis: text, projectId: text, projectName: text, customerName: text, schemeId: text, schemeName: text, stage: text, snapshotVersion: { type: 'integer', required: true }, irrNotice: text,
        subjects: { type: 'array', required: true, items: { type: 'object', additionalProperties: false, properties: {
                    key: text, name: text, displayName: text, groupId: text, side: text, inclTax: text, taxRate: text, custom_subject_name: text, billing_subject_name: text,
                    split: { type: 'boolean', required: true }, split_parts: { type: 'array', required: true, items: { type: 'object', additionalProperties: false, properties: { incl_tax: text, excl_tax: text } } },
                } } },
        groups: { type: 'array', required: true, items: { type: 'object', additionalProperties: false, properties: { groupId: text, side: text, inclTax: text } } },
        totals: { type: 'object', required: true, additionalProperties: false, properties: { revenueInclTax: text, costInclTax: text } },
        assumptions: { type: 'object', required: true, additionalProperties: true },
    } } as const;
export const readBenefitInputs = defineTool({
    name: 'read_benefit_inputs',
    description: '读取会话绑定项目指定方案的最新已保存输入：28个科目（9收入、19支出）、中文名称、自定义/开票名称、含税金额、税率、拆分明细、分组合计、资金计划与现金流假设。无需打开测算页，不接受projectId或整套输入，通用聊天不能读。先读此工具再回答科目金额，不要拿项目汇总或其他项目替代。结果仅代表已保存快照，不含编辑器未保存修改。本系统未计算IRR。',
    parameters: { scenario: scenarioParameter },
    output: { schema, render(_args, value) { return [{ type: 'text', text: `【已保存测算输入】\n${value.projectName} / ${value.schemeName} / ${value.stage} / 快照 v${value.snapshotVersion}\n${JSON.stringify(value)}` }]; } },
    timeoutMs: 30000, isConcurrencySafe: () => true,
    async execute(args, exec) { return postBridge<InferValue<typeof schema>>('/lamber-bridge/read-benefit-inputs', { ...args, sessionId: trustedSessionId(exec) }, exec.signal); },
});
