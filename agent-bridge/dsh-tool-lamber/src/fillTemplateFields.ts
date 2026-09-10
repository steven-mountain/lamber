import { defineTool } from '@deepseek-ai/dsh-tools';
import { postBridge } from './bridge.js';
import { authorizeTool, trustedSessionId } from './projectScope.js';
import { templateTextFields } from './templateFields.generated.js';
import { getCatalogTemplate } from './catalogCompletion.generated.js';
export const FILL_TEMPLATE_FIELDS = 'fill_template_fields';
const text = { type:'string',required:true } as const;
export const fillTemplateFields = defineTool({
  name:FILL_TEMPLATE_FIELDS,
  description:'经人工审批填入绑定项目的目录文本字段。先用read_template_fields读取目标模板及可写字段；templateId使用上下文完整模板名称；fields使用目录中可写文本key（如gen_demand_env_require部署环境要求、gen_demand_service_content服务内容）。只提交用户要求的文本，不得写金额、税率、年限、折现率、测算结果、图片或清单。弹窗展示原值和新值，用户可修改；返回的是实际批准并保存的字段。通用聊天和其他项目一律拒绝。审批期间模板改变会拒绝，须读取最新状态并重新审批，不能盲目重试。',
  parameters:{projectId:{...text,description:'当前会话绑定项目id'},templateId:{...text,description:'目录支持的完整模板名称'},fields:{type:'object',required:true,additionalProperties:false,properties:templateTextFields}},
  output:{schema:{type:'object',additionalProperties:false,properties:{workspaceId:text,projectId:text,templateId:text,fields:{type:'object',required:true,properties:templateTextFields,additionalProperties:false},templateVersion:{type:'integer',required:true},updatedAt:text,message:text}},
    render(_args,value){return [{type:'text',text:JSON.stringify(value)}];}},
  timeoutMs:15000,
  isConcurrencySafe:()=>false,
  async execute(_args,exec){
    const target = getCatalogTemplate(_args.templateId);
    if (!target || Object.keys(_args.fields).some(key => !target.fields.some(f => f.key === key && f.kind === 'text'))) throw new Error('字段不在目标模板文本目录中');
    await authorizeTool(exec);
    if(!exec.callId) throw new Error('缺少可信工具调用身份，拒绝写入');
    // Rust consumes the reviewed grant and applies it atomically; model args never become the write payload.
    return postBridge('/lamber-bridge/fill-template-fields',{sessionId:trustedSessionId(exec),callId:exec.callId,originalArgs:exec.arguments},exec.signal);
  },
});
