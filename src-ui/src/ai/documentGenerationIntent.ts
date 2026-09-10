import type { AiChatMessage } from './types';
import { templateCatalog } from '../lib/templateCompletion/catalog';

/** Invitations expire with the latest user turn; quoted/model text never starts an action. */
export function documentTemplateRequests(messages: readonly AiChatMessage[]): string[] {
  const content = [...messages].reverse().find(message => message.role === 'user')?.content ?? '';
  const text = content.replace(/```[\s\S]*?(?:```|$)/g, '').replace(/^\s*>.*$/gm, '')
    .replace(/[“「][\s\S]*?[”」]|"[^"\n]*"/g, '').replace(/[《》\s]/g, '');
  const clauses = text.split(/[，,。.!！?？;；\n]/).filter(clause =>
    !/(?:不要|不用|无需|别|暂不|先不|如何|怎么|为什么|只是|只问|取消)/.test(clause)
    && /(?:生成|导出|制作|出一份)/.test(clause));
  return templateCatalog.filter(template => !template.excludedReason && clauses.some(clause => clause.includes(template.name))).map(template => template.id);
}
export function documentGenerationPrompt(ids: readonly string[]): string {
  return ids.length
    ? 'The user requested document generation. The chat checks the bound project and offers a 生成文档 card for the requested supported template, independent of the open page. Use read_template_fields to report saved completion/missing fields; unknownCount > 0 is NOT complete. Direct the user to click the card: that click runs the existing product generation path and returns the actual output directory or error to chat. Missing fields do not prohibit clicking; existing financial/product checks still apply. You cannot generate files yourself or call generate_lifecycle_docs; never claim a file exists before the UI receipt. No AI approval dialog is needed for this user click.'
    : 'Document generation is available through a user-clicked chat card when the user explicitly asks to generate a catalog-supported template by name (需求导入表、立项签批表、甄选结果签批表). Do not claim a card is visible or a document was generated on this turn. You do not generate files directly.';
}
