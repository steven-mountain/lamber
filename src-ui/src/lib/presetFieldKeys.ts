export type CommonPresetKind = "short_value" | "text_snippet";

export type PresetFieldType =
  | "short_text"
  | "long_text"
  | "select"
  | "number"
  | "amount"
  | "percent"
  | "computed";

export interface PresetFieldDefinition {
  fieldKey: string;
  label: string;
  description?: string;
  templates: string[];
  groups: string[];
  fieldType: PresetFieldType;
  presetEligible: boolean;
  recommendedCategories: string[];
  aliases?: string[];
  kind: CommonPresetKind;
  category: string;
  defaultEnabled?: boolean;
}

export const PRESET_FIELD_KEYS = {
  projectCustomerName: "project_basic.customer_name",
  projectBackground: "project_basic.background",
  projectSolution: "project_basic.solution",
  projectPropertyRights: "project_basic.property_rights",

  approvalReviewers: "approval.reviewers",
  approvalDepartment: "approval.department",
  approvalProjectManager: "approval.project_manager",
  approvalItServiceContent: "approval.it_service_content",
  approvalCtServiceContent: "approval.ct_service_content",

  demandUnit: "demand.unit",
  demandServiceContent: "demand.service_content",
  demandCustomerConfirmation: "demand.customer_confirmation",
  demandDeploymentEnvironment: "demand.deployment_environment",

  meetingOnsiteSupport: "meeting.onsite_support",
  meetingItConstructionContent: "meeting.it_construction_content",
  meetingCtConstructionContent: "meeting.ct_construction_content",
  meetingTimeRequirement: "meeting.time_requirement",

  paymentRevenueCollectionMethod: "payment.revenue_collection_method",
  paymentExpenditurePaymentMethod: "payment.expenditure_payment_method",

  serviceDescription: "service.description",
  riskDescription: "risk.description",

  financeRevenueAmount: "finance.revenue_amount",
  financeExpenditureAmount: "finance.expenditure_amount",
  financeTaxRate: "finance.tax_rate",
  financeMarginRate: "finance.margin_rate",
  financeNpv: "finance.npv",
  financeAnnualCashflow: "finance.annual_cashflow",
  financeReverseTargetAmount: "finance.reverse_target_amount",
  financeBalanceAmount: "finance.balance_amount",
} as const;

function eligibleField(
  definition: Omit<
    PresetFieldDefinition,
    "presetEligible" | "category" | "recommendedCategories"
  > & {
    category: string;
    recommendedCategories?: string[];
  },
): PresetFieldDefinition {
  return {
    ...definition,
    presetEligible: true,
    recommendedCategories: definition.recommendedCategories ?? [definition.category],
  };
}

function excludedField(
  definition: Omit<
    PresetFieldDefinition,
    "presetEligible" | "recommendedCategories" | "defaultEnabled"
  >,
): PresetFieldDefinition {
  return {
    ...definition,
    presetEligible: false,
    recommendedCategories: [],
    defaultEnabled: false,
  };
}

export const PRESET_FIELD_REGISTRY: PresetFieldDefinition[] = [
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.projectCustomerName,
    label: "客户单位",
    description: "项目服务或建设对应的客户单位名称。",
    templates: ["ICT生命周期测算", "ICT项目需求导入表"],
    groups: ["项目概况"],
    fieldType: "short_text",
    category: "客户单位",
    kind: "short_value",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.projectBackground,
    label: "项目背景",
    description: "项目建设背景、客户现状与立项原因。",
    templates: ["ICT生命周期测算", "立项签批表", "会审纪要"],
    groups: ["项目概况"],
    fieldType: "long_text",
    category: "项目背景",
    kind: "text_snippet",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.projectSolution,
    label: "项目方案",
    description: "项目技术方案与总体实施思路。",
    templates: ["会审纪要"],
    groups: ["技术方案"],
    fieldType: "long_text",
    category: "项目方案",
    kind: "text_snippet",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.projectPropertyRights,
    label: "产权归属",
    templates: ["ICT生命周期测算"],
    groups: ["项目概况"],
    fieldType: "short_text",
    category: "产权归属",
    kind: "short_value",
    defaultEnabled: false,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.approvalReviewers,
    label: "审核人员",
    templates: ["会审纪要"],
    groups: ["参会与审核"],
    fieldType: "short_text",
    category: "审核人员",
    kind: "short_value",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.approvalDepartment,
    label: "部门名称",
    templates: ["会审纪要", "立项签批表"],
    groups: ["项目组织"],
    fieldType: "short_text",
    category: "部门名称",
    kind: "short_value",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.approvalProjectManager,
    label: "项目负责人",
    templates: ["立项签批表", "会审纪要"],
    groups: ["项目组织", "风险管理"],
    fieldType: "short_text",
    category: "项目负责人",
    kind: "short_value",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.approvalItServiceContent,
    label: "IT服务内容",
    templates: ["立项签批表"],
    groups: ["服务内容"],
    fieldType: "long_text",
    category: "IT服务内容",
    kind: "text_snippet",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.approvalCtServiceContent,
    label: "CT服务内容",
    templates: ["立项签批表"],
    groups: ["服务内容"],
    fieldType: "long_text",
    category: "CT服务内容",
    kind: "text_snippet",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.demandUnit,
    label: "项目需求单位",
    templates: ["ICT项目需求导入表"],
    groups: ["项目概况"],
    fieldType: "short_text",
    category: "项目需求单位",
    kind: "short_value",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.demandServiceContent,
    label: "服务说明",
    templates: ["ICT项目需求导入表"],
    groups: ["服务内容"],
    fieldType: "long_text",
    category: "服务说明",
    recommendedCategories: ["服务说明", "服务内容"],
    kind: "text_snippet",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.demandCustomerConfirmation,
    label: "客户确认方式",
    templates: ["ICT项目需求导入表"],
    groups: ["客户确认"],
    fieldType: "short_text",
    category: "客户确认",
    kind: "short_value",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.demandDeploymentEnvironment,
    label: "部署环境要求",
    templates: ["ICT项目需求导入表"],
    groups: ["技术条件"],
    fieldType: "long_text",
    category: "部署环境要求",
    kind: "text_snippet",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.meetingOnsiteSupport,
    label: "驻点支撑人员",
    templates: ["会审纪要"],
    groups: ["项目组织"],
    fieldType: "short_text",
    category: "驻点支撑人员",
    kind: "short_value",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.meetingItConstructionContent,
    label: "IT建设内容",
    templates: ["会审纪要"],
    groups: ["建设内容"],
    fieldType: "long_text",
    category: "IT建设内容",
    kind: "text_snippet",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.meetingCtConstructionContent,
    label: "CT建设内容",
    templates: ["会审纪要"],
    groups: ["建设内容"],
    fieldType: "long_text",
    category: "CT建设内容",
    kind: "text_snippet",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.meetingTimeRequirement,
    label: "时间要求",
    templates: ["会审纪要"],
    groups: ["实施要求"],
    fieldType: "long_text",
    category: "时间要求",
    kind: "text_snippet",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.paymentRevenueCollectionMethod,
    label: "收入条款",
    description: "项目收入侧收款节点、账期与方式。",
    templates: ["立项签批表", "会审纪要"],
    groups: ["商务条款"],
    fieldType: "long_text",
    category: "收付款方式",
    recommendedCategories: ["收入条款", "商务条款", "收付款方式"],
    aliases: ["contract.income_terms"],
    kind: "text_snippet",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.paymentExpenditurePaymentMethod,
    label: "支出条款",
    description: "项目支出侧付款节点、账期与方式。",
    templates: ["立项签批表", "会审纪要"],
    groups: ["商务条款"],
    fieldType: "long_text",
    category: "收付款方式",
    recommendedCategories: ["支出条款", "商务条款", "收付款方式"],
    aliases: ["contract.expenditure_terms"],
    kind: "text_snippet",
    defaultEnabled: true,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.serviceDescription,
    label: "服务说明",
    templates: ["ICT项目需求导入表"],
    groups: ["服务清单"],
    fieldType: "long_text",
    category: "服务说明",
    kind: "text_snippet",
    defaultEnabled: false,
  }),
  eligibleField({
    fieldKey: PRESET_FIELD_KEYS.riskDescription,
    label: "风险说明",
    templates: ["立项签批表", "会审纪要"],
    groups: ["风险管理"],
    fieldType: "long_text",
    category: "风险说明",
    kind: "text_snippet",
    defaultEnabled: false,
  }),
  excludedField({
    fieldKey: PRESET_FIELD_KEYS.financeRevenueAmount,
    label: "收入金额",
    templates: ["效益测算表"],
    groups: ["收入测算"],
    fieldType: "amount",
    category: "财务字段",
    kind: "short_value",
  }),
  excludedField({
    fieldKey: PRESET_FIELD_KEYS.financeExpenditureAmount,
    label: "支出金额",
    templates: ["效益测算表"],
    groups: ["支出测算"],
    fieldType: "amount",
    category: "财务字段",
    kind: "short_value",
  }),
  excludedField({
    fieldKey: PRESET_FIELD_KEYS.financeTaxRate,
    label: "税率",
    templates: ["效益测算表"],
    groups: ["税务参数"],
    fieldType: "percent",
    category: "财务字段",
    kind: "short_value",
  }),
  excludedField({
    fieldKey: PRESET_FIELD_KEYS.financeMarginRate,
    label: "利润率",
    templates: ["效益测算表"],
    groups: ["计算结果"],
    fieldType: "computed",
    category: "财务字段",
    kind: "short_value",
  }),
  excludedField({
    fieldKey: PRESET_FIELD_KEYS.financeNpv,
    label: "NPV",
    templates: ["效益测算表"],
    groups: ["计算结果"],
    fieldType: "computed",
    category: "财务字段",
    kind: "short_value",
  }),
  excludedField({
    fieldKey: PRESET_FIELD_KEYS.financeAnnualCashflow,
    label: "年度现金流",
    templates: ["效益测算表"],
    groups: ["现金流"],
    fieldType: "computed",
    category: "财务字段",
    kind: "short_value",
  }),
  excludedField({
    fieldKey: PRESET_FIELD_KEYS.financeReverseTargetAmount,
    label: "智能反算目标金额",
    templates: ["效益测算表"],
    groups: ["智能反算"],
    fieldType: "computed",
    category: "财务字段",
    kind: "short_value",
  }),
  excludedField({
    fieldKey: PRESET_FIELD_KEYS.financeBalanceAmount,
    label: "差额承接金额",
    templates: ["效益测算表"],
    groups: ["差额承接"],
    fieldType: "computed",
    category: "财务字段",
    kind: "short_value",
  }),
];

export const PRESET_FIELD_DEFINITIONS = PRESET_FIELD_REGISTRY.filter(
  field => field.presetEligible,
);

const FIELD_META_BY_KEY = new Map<string, PresetFieldDefinition>();
for (const field of PRESET_FIELD_REGISTRY) {
  FIELD_META_BY_KEY.set(field.fieldKey, field);
  for (const alias of field.aliases ?? []) {
    FIELD_META_BY_KEY.set(alias, field);
  }
}

export function getPresetFieldDefinition(fieldKey: string): PresetFieldDefinition | undefined {
  return FIELD_META_BY_KEY.get(fieldKey);
}

export function isPresetFieldEligible(fieldKey: string): boolean {
  return getPresetFieldDefinition(fieldKey)?.presetEligible === true;
}

export function getPresetFieldCategories(kind?: CommonPresetKind): string[] {
  const categories = PRESET_FIELD_DEFINITIONS
    .filter(field => !kind || field.kind === kind)
    .flatMap(field => field.recommendedCategories);
  return Array.from(new Set(categories)).sort((a, b) => a.localeCompare(b, "zh-CN"));
}

export function getPresetFieldDisplay(fieldKey: string) {
  const field = getPresetFieldDefinition(fieldKey);
  if (field) return field;
  console.warn(`Missing preset field metadata for "${fieldKey}"`);
  return {
    fieldKey,
    label: "未命名字段",
    description: "该字段尚未配置业务元信息。",
    templates: ["暂未配置"],
    groups: ["暂未配置"],
    fieldType: "short_text" as const,
    presetEligible: false,
    recommendedCategories: [],
    kind: "short_value" as const,
    category: "未分类",
    defaultEnabled: false,
  };
}

export function presetAppliesToField(
  applicableFieldKeys: string[] | undefined,
  fieldKey: string,
): boolean {
  if (!applicableFieldKeys || applicableFieldKeys.length === 0) return true;
  const target = getPresetFieldDefinition(fieldKey);
  const acceptedKeys = new Set([fieldKey, target?.fieldKey, ...(target?.aliases ?? [])].filter(Boolean));
  return applicableFieldKeys.some(key => acceptedKeys.has(key));
}
