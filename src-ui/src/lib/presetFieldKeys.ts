export type CommonPresetKind = "short_value" | "text_snippet";

export interface PresetFieldDefinition {
  fieldKey: string;
  label: string;
  category: string;
  kind: CommonPresetKind;
  description?: string;
}

export const PRESET_FIELD_KEYS = {
  projectCustomerName: "project_basic.customer_name",
  projectBackground: "project_basic.background",
  projectSolution: "project_basic.solution",

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
} as const;

export const PRESET_FIELD_DEFINITIONS: PresetFieldDefinition[] = [
  {
    fieldKey: PRESET_FIELD_KEYS.projectCustomerName,
    label: "客户单位",
    category: "客户单位",
    kind: "short_value",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.projectBackground,
    label: "项目背景",
    category: "项目背景",
    kind: "text_snippet",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.projectSolution,
    label: "项目方案",
    category: "项目方案",
    kind: "text_snippet",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.approvalReviewers,
    label: "审核人员",
    category: "审核人员",
    kind: "short_value",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.approvalDepartment,
    label: "部门名称",
    category: "部门名称",
    kind: "short_value",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.approvalProjectManager,
    label: "项目负责人",
    category: "项目负责人",
    kind: "short_value",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.approvalItServiceContent,
    label: "IT服务内容",
    category: "IT服务内容",
    kind: "text_snippet",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.approvalCtServiceContent,
    label: "CT服务内容",
    category: "CT服务内容",
    kind: "text_snippet",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.demandUnit,
    label: "项目需求单位",
    category: "项目需求单位",
    kind: "short_value",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.demandServiceContent,
    label: "服务内容",
    category: "服务内容",
    kind: "text_snippet",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.demandCustomerConfirmation,
    label: "客户确认",
    category: "客户确认",
    kind: "short_value",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.demandDeploymentEnvironment,
    label: "部署环境要求",
    category: "部署环境要求",
    kind: "text_snippet",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.meetingOnsiteSupport,
    label: "驻点支撑人员",
    category: "驻点支撑人员",
    kind: "short_value",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.meetingItConstructionContent,
    label: "IT建设内容",
    category: "IT建设内容",
    kind: "text_snippet",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.meetingCtConstructionContent,
    label: "CT建设内容",
    category: "CT建设内容",
    kind: "text_snippet",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.meetingTimeRequirement,
    label: "时间要求",
    category: "时间要求",
    kind: "text_snippet",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.paymentRevenueCollectionMethod,
    label: "收入侧收款方式",
    category: "收付款方式",
    kind: "text_snippet",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.paymentExpenditurePaymentMethod,
    label: "支出侧付款方式",
    category: "收付款方式",
    kind: "text_snippet",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.serviceDescription,
    label: "服务说明",
    category: "服务说明",
    kind: "text_snippet",
  },
  {
    fieldKey: PRESET_FIELD_KEYS.riskDescription,
    label: "风险说明",
    category: "风险说明",
    kind: "text_snippet",
  },
];

export function getPresetFieldDefinition(fieldKey: string): PresetFieldDefinition | undefined {
  return PRESET_FIELD_DEFINITIONS.find(field => field.fieldKey === fieldKey);
}

export function getPresetFieldCategories(kind?: CommonPresetKind): string[] {
  const categories = PRESET_FIELD_DEFINITIONS
    .filter(field => !kind || field.kind === kind)
    .map(field => field.category);
  return Array.from(new Set(categories)).sort((a, b) => a.localeCompare(b, "zh-CN"));
}

export function presetAppliesToField(
  applicableFieldKeys: string[] | undefined,
  fieldKey: string,
): boolean {
  if (!applicableFieldKeys || applicableFieldKeys.length === 0) return true;
  return applicableFieldKeys.includes(fieldKey);
}
