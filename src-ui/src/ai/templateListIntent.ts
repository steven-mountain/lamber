import type { AiChatMessage } from './types';
import type { TechItem } from '../services/templateListTypes';
export function templateListIntent(messages: readonly Pick<AiChatMessage, 'role'|'content'>[]) {
  const last = [...messages].reverse().find(message => message.role === 'user')?.content ?? '';
  const text = last.replace(/```[\s\S]*?(?:```|$)|^\s*>.*$/gm, '').replace(/[“「][\s\S]*?[”」]|"[^"\n]*"/g, '');
  if (/(?:不要|不用|无需|暂不|先不|别).{0,15}(?:清单|询价|报价|卡片)/.test(text)) return {tech:false,inquiry:false};
  return {tech:/(?:技术方案可行性清单|设备需求清单|技术清单)/.test(text) && /(?:建议|提议|拟|列|填|补|完善|编辑|修改|增加|删除|生成|做)/.test(text),
    inquiry:/(?:询价|三家报价)/.test(text) || (/会审纪要/.test(text) && /(?:填|补|完善|生成)/.test(text))};
}
/** Suggestions are untrusted display data. Only the user's adopt/save buttons can use these rows. */
export function techProposal(messages: readonly Pick<AiChatMessage,'role'|'content'>[]): TechItem[] {
  const lastUser = messages.reduce((last, message, index) => message.role === 'user' ? index : last, -1);
  const response = messages.slice(lastUser + 1).filter(message => message.role === 'assistant').map(message => message.content).join('\n');
  const lines = response.split('\n');
  const header = lines.findIndex(line => /^\s*\|\s*服务名称\s*\|\s*服务描述\s*\|\s*数量\s*\|\s*单位\s*\|\s*$/.test(line));
  if (header < 0) return [];
  const rows: TechItem[] = [];
  for (const line of lines.slice(header + 2)) {
    if (!/^\s*\|.*\|\s*$/.test(line)) break;
    const cells = line.trim().slice(1,-1).split('|').map(value => value.trim());
    if (cells.length !== 4 || !Number.isFinite(Number(cells[2])) || Number(cells[2]) < 0) return [];
    rows.push({serviceName:cells[0],serviceDesc:cells[1],amount:Number(cells[2]),unit:cells[3]});
    if (rows.length > 100) return [];
  }
  return rows;
}
export function templateListPrompt(intent: {tech:boolean;inquiry:boolean}) {
  return `清单规则：你不能调用工具写techItems或inqVendors。技术清单中的amount是数量，询价中的amount才是金额。
${intent.tech ? '本轮已邀请技术清单卡片。请根据已读取的项目背景和服务内容提出建议，用Markdown四列表格，表头严格为「服务名称 | 服务描述 | 数量 | 单位」，数量列只写数字，单元格不含竖线。用户点击「采用本轮建议」后可增删改并保存；此清单同时用于需求导入表与会审纪要。尚未保存时只能称建议。' : '本轮没有技术清单编辑邀请，不要声称技术清单卡片已出现。'}
${intent.inquiry ? '本轮可显示绑定项目的询价卡片。先读会审纪要保存态，缺询价时引导用户点击确认，按IT成本生成三家报价且最高价不超过含税总收入；不要给生成前数字预览。' : '本轮没有询价卡片邀请，不要声称询价卡片已出现。'}
无论何时，你都不得提议具体询价厂商名、报价金额、税率或修改生成后的询价内容。用户可自行在卡片改厂商名/金额和上传截图。询价生成校验失败时，原样转述错误并引导先完善IT成本或收入，绝不能建议手工加三行等绕过办法。用户点击确认才执行既有生成器，生成后展示实际结果，不能声称自己已生成。`;
}
