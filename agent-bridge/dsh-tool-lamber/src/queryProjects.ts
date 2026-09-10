import { defineTool, type InferValue } from '@deepseek-ai/dsh-tools';
import { postBridge } from './bridge.js';
import { trustedSessionId } from './projectScope.js';

export const QUERY_PROJECTS = 'query_projects';
export const QUERY_ROUTE = '/lamber-bridge/query-projects';
const range = (description: string) => ({ type: 'object', description, additionalProperties: false,
  properties: { gte: { type: 'number', description: '大于等于' }, gt: { type: 'number', description: '严格大于' }, lte: { type: 'number', description: '小于等于' }, lt: { type: 'number', description: '严格小于' } },
}) as const;
const time = (description: string) => ({ type: 'string', description: `${description}。RFC3339 可携带时区；YYYY-MM-DD 按 UTC 零点解释。` }) as const;
const text = { type: 'string', required: true } as const;
const number = { type: 'number', required: true } as const;
const nullableNumber = { oneOf: [{ type: 'number' }, { type: 'null' }], required: true } as const;
const nullableText = { oneOf: [{type:'string'}, {type:'null'}], required: true } as const;
const totals = {type:'object', additionalProperties:false, properties:{totalRevenueIncl:text,totalCostIncl:text,revenueValueCount:number,costValueCount:number}} as const;
const metrics = { type: 'object', additionalProperties: false, properties: {
  marginRate: nullableNumber, npv: nullableNumber, npvRate: nullableNumber, irr: nullableNumber, dynamicPayback: nullableNumber, riskLevel: text,
} } as const;
const outputSchema = { type: 'object', additionalProperties: false, properties: {
  scope: text, source: text, notice: text, queriedAt: text, message: text,
  boundProjectId: { oneOf: [{type:'string'}, {type:'null'}], required: true },
  matchedCount: number, returnedCount: number, appliedLimit: number, truncated: {type:'boolean', required:true},
  totals: {oneOf:[totals,{type:'null'}],required:true},
  mixedStages:{type:'boolean',required:true}, comparisonNotice:text,
  stageTotals:{type:'array',required:true,items:{type:'object',additionalProperties:false,properties:{stage:text,stageLabel:text,matchedCount:number,totals:{...totals,required:true}}}},
  projects: {type:'array',required:true,items:{type:'object',additionalProperties:false,properties:{
    id:text,name:text,defaultSchemeId:nullableText,stage:text,stageLabel:text,schemeName:nullableText,schemeUpdatedAt:nullableText,customerName:text,status:text,benefitStatus:text,createdAt:text,updatedAt:text,
    progress:nullableNumber,projectYears:number,discountRate:nullableNumber,totalRevenueIncl:nullableNumber,totalCostIncl:nullableNumber,
    isBoundProject:{type:'boolean',required:true},summaryMetrics:{oneOf:[metrics,{type:'null'}],required:true},
  }}},
} } as const;
export const queryProjects = defineTool({
  name: QUERY_PROJECTS,
  description: '检索当前工作区跨项目已保存汇总。通用聊天和项目会话均可用，无需审批。只读取 projects 与 summary_metrics 及默认方案的阶段/名称/保存时间元数据，不读任何明细表，不重新测算。支持客户、状态、时间、金额、指标过滤；金额为含税元，比例为小数（20%=0.2）。返回全体命中数量及按阶段金额合计；mixedStages=true时totals为空，不可跨口径合计或横向比较，财务指标按阶段内排序。回答指标必须带stageLabel、方案名和保存时间，未标注不得猜测，项目列表最多50条，缺失指标为null且不参与指标过滤。结果是跨项目检索，不能代替当前绑定项目结论；回答“我这个项目”时仅使用 isBoundProject=true 的项目，未命中时不可拿其他项目代替。',
  parameters: {
    customerName: { type: 'string', description: '客户名不区分大小写的包含匹配（普通文本，不是SQL或正则）。' },
    status: { type: 'string', description: '项目状态精确匹配；不确定状态名称时先查询，不要猜。' },
    benefitStatus: { type: 'string', description: '测算摘要状态精确匹配，如 normal / outdated / not_started。' },
    createdFrom: time('创建时间起点，包含'), createdBefore: time('创建时间终点，不包含'),
    updatedFrom: time('更新时间起点，包含'), updatedBefore: time('更新时间终点，不包含'),
    totalRevenueIncl: range('项目含税收入区间，元'), totalCostIncl: range('项目含税成本区间，元'),
    marginRate: range('毛利率/利润率区间，小数；低于20%传 {lt:0.2}'), npv: range('NPV区间，元'),
    npvRate: range('NPV率区间，小数'), irr: range('本系统未计算IRR，此筛选会明确报错，不得解释为没有匹配项目'), dynamicPayback: range('动态回收期区间，年'),
    sortBy: { type: 'string', enum: ['name', 'createdAt', 'updatedAt', 'totalRevenueIncl', 'totalCostIncl', 'marginRate', 'npv', 'npvRate', 'irr', 'dynamicPayback'], description: '排序字段，默认updatedAt；IRR未计算，按irr排序会报错；缺失值始终最后，同值按id稳定排序。' },
    sortOrder: { type: 'string', enum: ['asc','desc'], description: '默认desc。' },
    limit: { type: 'integer', description: '返回项目数，默认20，必须大于0；服务端硬上限50，超过上限会明确报告截断。合计仍覆盖全部命中。' },
  },
  output: { schema: outputSchema, render(_args, value) { return [{ type: 'text', text: JSON.stringify(value) }]; } },
  timeoutMs: 30_000,
  isConcurrencySafe: () => true,
  async execute(args, exec) {
    return postBridge<InferValue<typeof outputSchema>>(QUERY_ROUTE, { sessionId: trustedSessionId(exec), query: args }, exec.signal);
  },
});
