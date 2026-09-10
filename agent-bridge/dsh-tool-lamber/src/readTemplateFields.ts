import { defineTool, type InferValue } from '@deepseek-ai/dsh-tools';
import { postBridge } from './bridge.js';
import { trustedSessionId } from './projectScope.js';
import { getCatalogCompletion, type CompletionAsset } from './catalogCompletion.generated.js';
const text = { type: 'string', required: true } as const;
const number = { type: 'number', required: true } as const;
const bool = { type: 'boolean', required: true } as const;
const completionItem = {type:'object',additionalProperties:false,properties:{key:text,label:text,filled:bool,evaluated:bool,kind:text,templateName:text,usage:{type:'string'}}} as const;
const outputSchema = {type:'object',additionalProperties:false,properties:{
  projectId:text,projectName:text,templateId:text,source:text,hasSavedState:bool,templateVersion:number,
  fields:{type:'array',required:true,items:{type:'object',additionalProperties:false,properties:{key:text,label:text,value:text,valueSource:text,originalCharacters:number,truncated:bool}}},
  techItems:{type:'array',required:true,items:{type:'object',additionalProperties:false,properties:{serviceName:text,serviceDesc:text,amount:text,unit:text}}},
  hasPublicUrl:bool,hasSecurity:bool,
  attachments:{type:'array',required:true,items:{type:'object',additionalProperties:false,properties:{fieldKey:text,exists:bool}}},
  truncated:bool,textLimit:number,originalCharacters:number,returnedCharacters:number,techItemCount:number,returnedTechItemCount:number,notice:text,
  completion:{type:'array',required:true,items:completionItem},missingFields:{type:'array',required:true,items:completionItem},
  filledCount:number,totalCount:number,unknownCount:number,
  catalogId:{type:'string'},
  readOnlyFields:{type:'array',items:{type:'object',additionalProperties:true}},
  lists:{type:'object',additionalProperties:true},listCounts:{type:'object',additionalProperties:true},returnedListCounts:{type:'object',additionalProperties:true},
}} as const;
type Projection = Omit<InferValue<typeof outputSchema>, 'completion'|'missingFields'|'filledCount'|'totalCount'|'unknownCount'> & {completionState:Record<string,unknown>};
export const readTemplateFields = defineTool({
  name: 'read_template_fields',
  description: '主动读取当前会话绑定项目的目录支持模板已保存文本、只读字段分类及缺项，无需打开模板页，无需审批。不接受projectId，通用聊天拒绝。先读取现状再判断填空或修改。templateId可用目录模板别名定位本项目唯一已保存模板；有多张时按候选名称再查。返回文本原文、技术清单和附件存在状态，不返回本机路径。valueSource=default是界面默认值而非已保存填写；truncated=true表示超过总文本24000字/清单100行上限，不能声称读到了全文。techItems只允许模型提出服务名称/描述/数量/单位建议，由用户在聊天清单卡片编辑并保存，同时用于需求导入表与会审纪要。inqVendors只允许引导用户确认调用原生成器，模型不得提出具体厂商名/报价/税率，不得修改生成后的报价；失败须原样转述并引导补IT成本或收入，不给手工加三行等绕过办法。两张清单都没有AI写工具。甄选结果签批表的项目背景、甄选后方案、公共字段一致、批次字段差异已确认、续签成本归类已确认、立项金额低于50万元这6项，聊天侧为unknown（缺少页面运行期事实），请用户到该表页面查看；不得把unknown列为已完成或未完成。合并项目名称按保存的单项目/批量模式判断且不可由AI写入。derived/check不可写，业务校验及动态默认值未能判定时evaluated=false并计入unknownCount，不能说整表已完成；不应调用bash/glob读取空白docx或工作区文件。仅在用户本轮明确要求填写或生成需求导入表时，需求表缺失图片才会触发聊天上传卡片；绑定项目、普通聊天或只读查询不触发，也不要主动催上传。卡片激活时，用户可选择图片或聚焦后粘贴补齐，无需打开模板页；模型不能直接写图片，普通聊天附图不会自动入库。遵守本轮demand_image_invitation提示，不要声称未激活的卡片已经显示，也不要说图片只能去聊天外界面上传。已有需求表图片可通过聊天“项目图片”卡片展示；用户选中具体图片，选择或粘贴新图，查看新旧对照并点击确认后才替换。该卡片也可将选中图片加入输入框供视觉模型分析；资产元数据不等于已看见图片内容。普通上传是追加，不会自动覆盖旧图。模型不得用read_image/bash/glob猜路径，也没有裁剪或改字的图像编辑服务。用户明确请求生成目录模板时，聊天提供生成文档卡片，用户点击后走产品原有生成流程，结果回执实际目录或原始错误。缺项仍可点击，有unknownCount不能称已齐备。模型不能自行生成文件或调用generate_lifecycle_docs；应引导用户点击卡片，不要只回答无法生成。',
  parameters: {templateId:{type:'string',required:true,description:'目录中的完整模板名，或模板别名定位当前绑定项目唯一已保存模板。'}},
  output:{schema:outputSchema,render(_args,value){return [{type:'text',text:JSON.stringify(value)}];}},
  timeoutMs:30_000,
  isConcurrencySafe:()=>true,
  async execute(args,exec){
    // Keep unrecognized model keys in the envelope so Rust rejects them too.
    const result = await postBridge<Projection>('/lamber-bridge/read-template-fields',{...args,sessionId:trustedSessionId(exec)},exec.signal);
    const {completionState,...projection}=result;
    const completion=getCatalogCompletion(result.templateId,completionState,result.attachments as CompletionAsset[]);
    return {...projection,completion,missingFields:completion.filter(item=>item.evaluated && !item.filled),filledCount:completion.filter(item=>item.filled).length,totalCount:completion.length,unknownCount:completion.filter(item=>!item.evaluated).length};
  },
});
