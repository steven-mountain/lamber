import { defineTool, type InferValue } from '@deepseek-ai/dsh-tools';
import { postBridge } from './bridge.js';
import { trustedSessionId } from './projectScope.js';
const text = { type: 'string', required: true } as const;
const schema = { type: 'object', additionalProperties: false, properties: { selection_fee_excl: text, selection_fee_incl: text, quote_excl: text, quote_candidates: { type: 'array', required: true, items: { type: 'string' } }, actual_cost: text, final_limit: text, quote: text } } as const;
type Result = InferValue<typeof schema>;
const description = '报价/限价/浮动均按含税处理，按报价除以1.06后的不含税基数查档；quote_excl四位小数仅供展示，不得用展示值另算。全部金额由原Rust甄选费引擎计算。纯计算，不读取或写入项目，项目会话和通用聊天均可用，不接受projectId。后端错误须原样转述：坏输入不得当0，无解不得估算或给绕过办法。quote_candidates多于一个时必须列出全部精确解、说明当前采用较低报价，可用正算指定报价。目标投入科目必须是6%税率，非6%会被产品拦截，不要建议绕过；本工具不能写科目。';
const output = { schema, render(_args: unknown, value: Result) {
        return [{ type: 'text' as const, text: [
                    '【甄选费纯计算 · 未写入项目】',
                    `含税报价：${value.quote}；不含税计费基数（展示4位小数）：${value.quote_excl}`,
                    `不含税服务费：${value.selection_fee_excl}；含税服务费：${value.selection_fee_incl}`,
                    `含税实际成本：${value.actual_cost}；含税限价：${value.final_limit}`,
                    '报价按含税处理；档位按除以1.06后的不含税基数查找。',
                    value.quote_candidates.length > 1 ? `存在多个精确解：${value.quote_candidates.join('、')}。当前采用较低报价，可改用报价正算指定。` : `精确报价候选：${value.quote_candidates.join('、') || '无'}`,
                    '目标投入科目须为6%税率；结果未写入任何科目。',
                ].join('\n') }];
    } };
const markup = { ...text, description: '含税浮动金额，无单位数字字符串；可以有符号。请显式传入，零浮动为0。' };
async function forward(route: string, args: Record<string, unknown>, exec: Parameters<typeof trustedSessionId>[0] & {
    signal: AbortSignal;
}) {
    const { selection_fee: _legacy, ...value } = await postBridge<Result & {
        selection_fee: string;
    }>(route, { ...args, sessionId: trustedSessionId(exec) }, exec.signal);
    return value;
}
export const calculateSelectionFee = defineTool({ name: 'calculate_selection_fee', description: '输入含税报价计算甄选服务费及含税限价。' + description,
    parameters: { quote: { ...text, description: '含税报价，最多两位小数，无单位数字字符串。' }, markup }, output,
    timeoutMs: 30000, isConcurrencySafe: () => true, execute: (args, exec) => forward('/lamber-bridge/calculate-selection-fee', args, exec) });
export const reverseCalculateSelectionFee = defineTool({ name: 'reverse_calculate_selection_fee', description: '输入含税限价精确反算报价并返回全部候选。' + description,
    parameters: { limit: { ...text, description: '含税限价，最多两位小数，无单位数字字符串。' }, markup }, output,
    timeoutMs: 30000, isConcurrencySafe: () => true, execute: (args, exec) => forward('/lamber-bridge/reverse-calculate-selection-fee', args, exec) });
