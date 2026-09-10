import subjectDefinitions from "./ictSubjects.json";
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
  defaultTaxRate: number;
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

export const ICT_SUBJECT_DEFINITIONS: IctSubjectDefinition[] = subjectDefinitions as IctSubjectDefinition[];

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
