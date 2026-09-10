import { SYSTEM_PROMPT_KNOWLEDGE } from '../../lib/knowledgeBase';
import { sessionScopePrompt } from '../sessionScopePolicy';
import { documentTemplateRequests, documentGenerationPrompt } from '../documentGenerationIntent';
import { templateListIntent, templateListPrompt, techProposal } from '../templateListIntent';
import { structureReverseIntent, structureReversePrompt } from '../structureReverseIntent';
import { isDemandFormRequest, demandImageCompletionPrompt } from '../demandFormIntent';
import { wantsTemplateImages, templateImagePrompt } from '../templateImageIntent';
import type { AiChatMessage } from '../types';
export function businessPresentation(messages: AiChatMessage[]) {
  const user = [...messages].reverse().find(message => message.role === 'user')?.content ?? '';
  return { documents: documentTemplateRequests(messages), lists: templateListIntent(messages), proposal: techProposal(messages),
    reverse: structureReverseIntent(messages), demandImages: isDemandFormRequest(user), savedImages: wantsTemplateImages(user) };
}
export type BusinessPresentation = ReturnType<typeof businessPresentation>;
export function businessPrompt(messages: AiChatMessage[], projectId: string | null) {
  const current = businessPresentation(messages);
  return [SYSTEM_PROMPT_KNOWLEDGE, sessionScopePrompt(projectId),
    '你是 Lamber 中文售前业务助手。优先按产品编号匹配已有知识库；信息不足时明确说明。金额默认单位为人民币元；用户明确要求万元时才换算并标明。',
    '项目工具返回的是保存态；当前编辑器的未保存更改不能称已保存。不同项目、不同方案、不同版本的数据必须分别标注，不能混合。图片资产元数据不代表已经看见图片；仅用户显式加入本轮输入的图片才是视觉内容。',
    '历史业务回执只描述操作当时的结果；其中的“尚未保存”不能代替本轮保存态，后续保存与编辑应以当前读取结果为准。恢复回执不代表再次执行。',
    '官方聊天输入区提供 Lamber 业务按钮，用户点击可打开相应卡片；技术清单建议、询价、文档、图片和结构反算沿用原服务。不要声称没有这些用户操作入口。工具成功回执和历史业务回执才是已执行的证据，邀请和预览都不是已保存。',
    demandImageCompletionPrompt(Boolean(projectId) && current.demandImages),
    templateImagePrompt(Boolean(projectId) && current.savedImages),
    documentGenerationPrompt(projectId ? current.documents : []),
    templateListPrompt(projectId ? current.lists : { tech: false, inquiry: false }),
    structureReversePrompt(Boolean(projectId) && current.reverse.requested),
  ].join('\n\n');
}
