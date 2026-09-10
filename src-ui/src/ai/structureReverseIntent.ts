import type { AiChatMessage } from './types';
import type { ReverseMetric } from '../lib/structureReverseResult';
export function structureReverseIntent(messages: readonly Pick<AiChatMessage, 'role' | 'content'>[]) {
  const index = messages.reduce((last, item, i) => item.role === 'user' ? i : last, -1);
  const raw = messages[index]?.content ?? '';
  const text = raw.replace(/```[\s\S]*?(?:```|$)|^\s*>.*$/gm, '').replace(/[“「][\s\S]*?[”」]|"[^"\n]*"/g, '');
  const metricType: ReverseMetric = /(?:NPV\s*率|净现值率)/i.test(text) ? 'npv_rate' : 'margin';
  const affirmativeClauses = text.split(/[，。！？;；\n]/).filter(clause =>
    !/(?:不要|不用|无需|暂不|先不|取消|别)/.test(clause)
    && !/(?:刚才|上次|上一轮|此前|之前|昨天)/.test(clause));
  const requested = affirmativeClauses.some(clause =>
    !/(?:甄选费|限价).{0,6}反算/.test(clause)
    && (/(?:结构反算|智能反算)/.test(clause) || (/(?:利润率|净现值率|NPV\s*率)/i.test(clause) && /(?:做到|达到|调整为|调整到|设为|提高到|降到|反算)/.test(clause))));
  const targetText = affirmativeClauses.filter(clause =>
    /(?:目标|利润率|净现值率|NPV\s*率|做到|达到|调整为|调整到|设为|提高到|降到)/i.test(clause)).join('，');
  const values = [...targetText.matchAll(/(-?\d+(?:\.\d+)?)\s*[%％]/g)];
  return { requested, key: `${index}:${raw}`, metricType,
    targetPercent: requested && values.length === 1 ? values[0][1] : '',
    scenario: /甄选后|post_selection/.test(text) ? 'post_selection' : /甄选前|pre_selection/.test(text) ? 'pre_selection' : '' };
}
export function structureReversePrompt(requested: boolean) {
  return `结构反算只能由用户点击聊天卡片执行，没有AI反算写入工具。${requested ? '本轮已邀请绑定项目的结构反算卡片。用户必须自己选择目标科目，先查看该科目的可达范围，再确认目标指标和目标值；没有用户明确说出的目标数字时留空。' : '本轮没有结构反算卡片邀请，不要声称卡片已显示。'}
你只能转述用户自己明确说出的目标指标数字，不得发明目标值、代选科目、建议具体科目金额、自动降低目标或绕过总额锁定/承接/财务核验。用户确认才写编辑器，保存仍是用户的下一步操作。
结构反算只支持毛利润率和净现值率，不支持NPV金额或IRR作为目标。6%税率限制只属于采购甄选费回填，不适用于本结构反算卡片，不能混用两种反算规则。用户未询问税率时，不主动介绍甄选费规则或税率数字。介绍操作时按“用户选科目及指标类型→读取范围→用户确认目标值→确认写入”的顺序，不附加未请求的金额示例。
应用回执的目标值与实际达成值必须同时逐字引用，保留回执给出的四位百分数，不能只说已完成，也不能把范围内的任意值说成必然可达。
失败回执中的原错误必须完整放入一个text代码块，逐字复制，保留原有换行；不得在原文内部加粗数字、拆行、改标点或改写。包括目标、最接近值、范围和最终复算值，不能缩成无法达到。
原错误之后只可引导用户重新选择科目、查看范围、自行确认目标或补齐原文指出的前置条件。禁止自行追加数值关系、原因推测或解读，尤其禁止把“没有稳定解”解释成“目标落在范围外”——落在采样范围内同样可能不可达。
成功后逐科目金额和收付款计划前后数值、每年尾差来自应用回执，不自行算数或省略。旧model_e分板块不是本次联动计划。`;
}
