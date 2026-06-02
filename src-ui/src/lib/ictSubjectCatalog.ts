export type IctSubjectGroupId =
  | "revIt"
  | "revCt"
  | "revNonItCt"
  | "costIt"
  | "costCt"
  | "costMix";

export type IctSubjectSide = "revenue" | "cost";
export type IctDocumentPrefix = "IT" | "CT" | "非IT/CT" | "综合类";

export interface IctSubjectDefinition {
  subjectCode: string;
  groupId: IctSubjectGroupId;
  key: string;
  side: IctSubjectSide;
  standardSubjectName: string;
  documentPrefix: IctDocumentPrefix;
  excelVariablePrefix: string;
}

export interface IctTaxItemLike {
  incl?: number | string | null;
  tax?: number | string | null;
  excl?: number | string | null;
  customSubjectName?: string | null;
  custom_subject_name?: string | null;
  billingSubjectName?: string | null;
  billing_subject_name?: string | null;
}

export interface ResolvedBillingSubject {
  standardName: string;
  productOrBusinessName: string;
  billingSubjectName: string;
  excelDisplayName: string;
  documentBusinessName: string;
  documentDedupKey: string;
}

export const ICT_SUBJECT_DEFINITIONS: IctSubjectDefinition[] = [
  { subjectCode: "rev_it_integration", groupId: "revIt", key: "integration", side: "revenue", standardSubjectName: "系统集成服务收入", documentPrefix: "IT", excelVariablePrefix: "EXCEL_REV_IT_INTEGRATION" },
  { subjectCode: "rev_it_maintenance", groupId: "revIt", key: "maintenance", side: "revenue", standardSubjectName: "维保收入", documentPrefix: "IT", excelVariablePrefix: "EXCEL_REV_IT_MAINTENANCE" },
  { subjectCode: "rev_it_device_sales", groupId: "revIt", key: "device_sales", side: "revenue", standardSubjectName: "设备销售收入", documentPrefix: "IT", excelVariablePrefix: "EXCEL_REV_IT_DEVICE_SALES" },
  { subjectCode: "rev_it_device_lease", groupId: "revIt", key: "device_lease", side: "revenue", standardSubjectName: "设备租赁收入", documentPrefix: "IT", excelVariablePrefix: "EXCEL_REV_IT_DEVICE_LEASE" },
  { subjectCode: "rev_it_other", groupId: "revIt", key: "other", side: "revenue", standardSubjectName: "其他收入", documentPrefix: "IT", excelVariablePrefix: "EXCEL_REV_IT_OTHER" },
  { subjectCode: "rev_it_cloud", groupId: "revIt", key: "cloud", side: "revenue", standardSubjectName: "移动云-定制化收入", documentPrefix: "IT", excelVariablePrefix: "EXCEL_REV_IT_CLOUD" },
  { subjectCode: "rev_ct_line", groupId: "revCt", key: "line", side: "revenue", standardSubjectName: "专线收入", documentPrefix: "CT", excelVariablePrefix: "EXCEL_REV_CT_LINE" },
  { subjectCode: "rev_ct_product", groupId: "revCt", key: "product", side: "revenue", standardSubjectName: "产品收入", documentPrefix: "CT", excelVariablePrefix: "EXCEL_REV_CT_PRODUCT" },
  { subjectCode: "rev_non_it_ct", groupId: "revNonItCt", key: "item", side: "revenue", standardSubjectName: "工程施工收入等", documentPrefix: "非IT/CT", excelVariablePrefix: "EXCEL_REV_NON_IT_CT" },

  { subjectCode: "cost_it_device", groupId: "costIt", key: "device", side: "cost", standardSubjectName: "主要设备/甲供材料", documentPrefix: "IT", excelVariablePrefix: "EXCEL_COST_IT_DEVICE" },
  { subjectCode: "cost_it_construction", groupId: "costIt", key: "construction", side: "cost", standardSubjectName: "施工", documentPrefix: "IT", excelVariablePrefix: "EXCEL_COST_IT_CONSTRUCTION" },
  { subjectCode: "cost_it_survey", groupId: "costIt", key: "survey", side: "cost", standardSubjectName: "勘察设计/预备费", documentPrefix: "IT", excelVariablePrefix: "EXCEL_COST_IT_SURVEY" },
  { subjectCode: "cost_it_integration", groupId: "costIt", key: "integration", side: "cost", standardSubjectName: "集成服务", documentPrefix: "IT", excelVariablePrefix: "EXCEL_COST_IT_INTEGRATION" },
  { subjectCode: "cost_it_other", groupId: "costIt", key: "other", side: "cost", standardSubjectName: "其他投入", documentPrefix: "IT", excelVariablePrefix: "EXCEL_COST_IT_OTHER" },
  { subjectCode: "cost_it_maintenance", groupId: "costIt", key: "maintenance", side: "cost", standardSubjectName: "维护费用", documentPrefix: "IT", excelVariablePrefix: "EXCEL_COST_IT_MAINTENANCE" },
  { subjectCode: "cost_it_running", groupId: "costIt", key: "running", side: "cost", standardSubjectName: "其他运行支出（电费等）", documentPrefix: "IT", excelVariablePrefix: "EXCEL_COST_IT_RUNNING" },
  { subjectCode: "cost_it_bidding", groupId: "costIt", key: "bidding", side: "cost", standardSubjectName: "中标服务费", documentPrefix: "IT", excelVariablePrefix: "EXCEL_COST_IT_BIDDING" },
  { subjectCode: "cost_it_design_eval", groupId: "costIt", key: "design_eval", side: "cost", standardSubjectName: "设计院成本评估费", documentPrefix: "IT", excelVariablePrefix: "EXCEL_COST_IT_DESIGN_EVAL" },
  { subjectCode: "cost_it_audit", groupId: "costIt", key: "audit", side: "cost", standardSubjectName: "第三方审计评估费", documentPrefix: "IT", excelVariablePrefix: "EXCEL_COST_IT_AUDIT" },
  { subjectCode: "cost_ct_construction", groupId: "costCt", key: "construction", side: "cost", standardSubjectName: "专线建设", documentPrefix: "CT", excelVariablePrefix: "EXCEL_COST_CT_CONSTRUCTION" },
  { subjectCode: "cost_ct_maintenance", groupId: "costCt", key: "maintenance", side: "cost", standardSubjectName: "专线维护", documentPrefix: "CT", excelVariablePrefix: "EXCEL_COST_CT_MAINTENANCE" },
  { subjectCode: "cost_ct_other", groupId: "costCt", key: "other", side: "cost", standardSubjectName: "其他产品成本", documentPrefix: "CT", excelVariablePrefix: "EXCEL_COST_CT_OTHER" },
  { subjectCode: "cost_ct_bandwidth", groupId: "costCt", key: "bandwidth", side: "cost", standardSubjectName: "专线带宽成本", documentPrefix: "CT", excelVariablePrefix: "EXCEL_COST_CT_BANDWIDTH" },
  { subjectCode: "cost_ct_renewal", groupId: "costCt", key: "renewal", side: "cost", standardSubjectName: "专线/其他产品续签成本", documentPrefix: "CT", excelVariablePrefix: "EXCEL_COST_CT_RENEWAL" },
  { subjectCode: "cost_non_it_ct", groupId: "costMix", key: "non_it_ct", side: "cost", standardSubjectName: "工程施工投入等", documentPrefix: "非IT/CT", excelVariablePrefix: "EXCEL_COST_NON_IT_CT" },
  { subjectCode: "cost_mix_marketing", groupId: "costMix", key: "marketing", side: "cost", standardSubjectName: "融合营销成本", documentPrefix: "综合类", excelVariablePrefix: "EXCEL_COST_MIX_MARKETING" },
  { subjectCode: "cost_mix_channel", groupId: "costMix", key: "channel", side: "cost", standardSubjectName: "渠道酬金", documentPrefix: "综合类", excelVariablePrefix: "EXCEL_COST_MIX_CHANNEL" },
  { subjectCode: "cost_mix_other", groupId: "costMix", key: "other", side: "cost", standardSubjectName: "其他管理费用等", documentPrefix: "综合类", excelVariablePrefix: "EXCEL_COST_MIX_OTHER" },
];

export const ICT_SUBJECT_GROUPS = ICT_SUBJECT_DEFINITIONS.reduce((groups, subject) => {
  if (!groups[subject.groupId]) groups[subject.groupId] = [];
  groups[subject.groupId].push(subject);
  return groups;
}, {} as Record<IctSubjectGroupId, IctSubjectDefinition[]>);

export const normalizeCustomSubjectName = (value: unknown) => String(value ?? "").trim();

export const getSubjectCustomName = (item?: IctTaxItemLike | null) => {
  return normalizeCustomSubjectName(item?.customSubjectName ?? item?.custom_subject_name ?? "");
};

export const getSubjectBillingName = (item?: IctTaxItemLike | null) => {
  return normalizeCustomSubjectName(item?.billingSubjectName ?? item?.billing_subject_name ?? "");
};

export const resolveBillingSubjectPresentation = (
  subject: IctSubjectDefinition,
  item?: IctTaxItemLike | null,
  options: { fallbackDocumentBusinessName?: string | null; useStandardDocumentFallback?: boolean } = {},
): ResolvedBillingSubject => {
  const standardName = normalizeCustomSubjectName(subject.standardSubjectName);
  const productOrBusinessName = getSubjectCustomName(item);
  const billingSubjectName = getSubjectBillingName(item);
  const preferredDisplayName = billingSubjectName || productOrBusinessName;
  const excelDisplayName = preferredDisplayName ? `${standardName}（${preferredDisplayName}）` : standardName;
  const fallbackDocumentBusinessName = normalizeCustomSubjectName(options.fallbackDocumentBusinessName);
  const documentName = preferredDisplayName
    ? `${subject.documentPrefix}-${preferredDisplayName}`
    : fallbackDocumentBusinessName || (options.useStandardDocumentFallback ? `${subject.documentPrefix}-${standardName}` : "");

  return {
    standardName,
    productOrBusinessName,
    billingSubjectName,
    excelDisplayName,
    documentBusinessName: documentName,
    documentDedupKey: documentName,
  };
};

export const getSubjectExcelDisplayName = (subject: IctSubjectDefinition, item?: IctTaxItemLike | null) => {
  return resolveBillingSubjectPresentation(subject, item).excelDisplayName;
};

export const getSubjectDocumentBusinessName = (subject: IctSubjectDefinition, item?: IctTaxItemLike | null) => {
  return resolveBillingSubjectPresentation(subject, item).documentBusinessName;
};

export const hasSubjectAmount = (item?: IctTaxItemLike | null) => {
  const incl = Number(item?.incl ?? 0);
  const excl = Number(item?.excl ?? 0);
  return Math.abs(Number.isFinite(incl) ? incl : 0) > 0.005 || Math.abs(Number.isFinite(excl) ? excl : 0) > 0.005;
};

export const getProjectDataSubjectItem = (projectData: any, subject: IctSubjectDefinition): IctTaxItemLike | null => {
  if (subject.groupId === "revIt") return projectData.revenue?.it?.[subject.key] || null;
  if (subject.groupId === "revCt") return projectData.revenue?.ct?.[subject.key] || null;
  if (subject.groupId === "revNonItCt") return projectData.revenue?.non_it_ct || null;
  if (subject.groupId === "costIt") return projectData.cost?.it?.[subject.key] || null;
  if (subject.groupId === "costCt") return projectData.cost?.ct?.[subject.key] || null;
  if (subject.groupId === "costMix") return projectData.cost?.mix?.[subject.key] || null;
  return null;
};

export const buildExcelSubjectVariables = (projectData: any) => {
  return ICT_SUBJECT_DEFINITIONS.reduce<Record<string, string>>((variables, subject) => {
    const item = getProjectDataSubjectItem(projectData, subject);
    const shouldWriteAmount = hasSubjectAmount(item);
    const resolved = resolveBillingSubjectPresentation(subject, item);
    variables[`${subject.excelVariablePrefix}_NAME`] = resolved.excelDisplayName;
    variables[`${subject.excelVariablePrefix}_CUSTOM_NAME`] = resolved.productOrBusinessName;
    variables[`${subject.excelVariablePrefix}_BILLING_NAME`] = resolved.billingSubjectName;
    variables[`${subject.excelVariablePrefix}_DOCUMENT_NAME`] = resolved.documentBusinessName;
    variables[`${subject.excelVariablePrefix}_EXCL`] = shouldWriteAmount ? String(item?.excl ?? 0) : "";
    variables[`${subject.excelVariablePrefix}_INCL`] = shouldWriteAmount ? String(item?.incl ?? 0) : "";
    return variables;
  }, {});
};

export const collectDocumentBusinessNames = (
  projectData: any,
  options: {
    side?: IctSubjectSide;
    documentPrefix?: IctDocumentPrefix;
    groupId?: IctSubjectGroupId;
  } = {},
) => {
  const names: string[] = [];
  const seen = new Set<string>();

  ICT_SUBJECT_DEFINITIONS.forEach(subject => {
    if (options.side && subject.side !== options.side) return;
    if (options.documentPrefix && subject.documentPrefix !== options.documentPrefix) return;
    if (options.groupId && subject.groupId !== options.groupId) return;

    const item = getProjectDataSubjectItem(projectData, subject);
    const resolved = resolveBillingSubjectPresentation(subject, item);
    const name = resolved.documentBusinessName;
    const dedupKey = resolved.documentDedupKey;
    if (!name || !dedupKey || !hasSubjectAmount(item) || seen.has(dedupKey)) return;

    seen.add(dedupKey);
    names.push(name);
  });

  return names;
};
