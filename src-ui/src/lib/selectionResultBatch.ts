import Decimal from "decimal.js"
import type {
  BenefitAnalysisScheme,
  Project,
} from "../utils/projectService"
import { validateFinancialData, type ValidationReport } from "./financeValidator"
import {
  getProjectDataSubjectItem,
  ICT_SUBJECT_DEFINITIONS,
  resolveBillingSubjectPresentation,
  type IctSubjectDefinition,
  type IctSubjectSide,
} from "./ictSubjectCatalog"
import { exclFromIncl, restoreTaxSplitParts, roundMoneyHalfUp } from "./taxAmount"

export type SelectionRenewalDecision = "include" | "exclude"

export type SelectionSharedFieldKey =
  | "winnerName"
  | "scope"
  | "industry"
  | "method"
  | "rule"
  | "standardPlan"
  | "revenueCollection"
  | "expenditurePayment"

export interface SelectionSharedFields {
  winnerName: string
  scope: string
  industry: string
  method: string
  rule: string
  standardPlan: string
  revenueCollection: string
  expenditurePayment: string
}

export interface SelectionResultCandidate {
  project: Project
  preScheme: BenefitAnalysisScheme | null
  postScheme: BenefitAnalysisScheme | null
  eligible: boolean
  reason: string
}

export interface SelectionResultBatchProject {
  projectId: string
  projectName: string
  customerName: string
  postSchemeId: string
  postSchemeName: string
  preSchemeId: string
  preSchemeName: string
  projectData: any
  preSelectionCostIt: Record<string, any>
  metrics: any
  projectBackground: string
  sharedFields: SelectionSharedFields
  validationErrors: ValidationReport[]
  missingMetrics: string[]
}

export interface SelectionSharedConflict {
  key: SelectionSharedFieldKey
  label: string
  blocking: boolean
  values: Array<{ projectName: string; value: string }>
}

export interface SelectionResultBatchModel {
  tableA: Record<string, string>[]
  tableB: Record<string, string>[]
  tableC: Record<string, string>[]
  tableD: Record<string, string>[]
  tableE: Record<string, string>[]
  totalLimitExcl: Decimal
  totalWinnerExcl: Decimal
  totalWinnerIncl: Decimal
  totalCostExcl: Decimal
  totalCostIncl: Decimal
  totalRevenueExcl: Decimal
  totalRevenueIncl: Decimal
  approvalAmountExcl: Decimal
  investmentSituation: string
  revenueSituation: string
  renewalProjects: Array<{ projectId: string; projectName: string; amountExcl: Decimal }>
}

const IT_COST_SUBJECTS = ICT_SUBJECT_DEFINITIONS.filter(
  subject => subject.side === "cost" && subject.groupId === "costIt",
)
const CT_DEDICATED_LINE_KEYS = new Set(["construction", "maintenance", "bandwidth"])

const SHARED_FIELD_CONFIG: Array<{
  key: SelectionSharedFieldKey
  label: string
  blocking: boolean
}> = [
  { key: "winnerName", label: "中选合作伙伴", blocking: true },
  { key: "method", label: "甄选方式", blocking: true },
  { key: "rule", label: "甄选规则", blocking: true },
  { key: "revenueCollection", label: "客户支付方式", blocking: true },
  { key: "expenditurePayment", label: "合作伙伴支付方式", blocking: true },
  { key: "scope", label: "甄选范围", blocking: false },
  { key: "industry", label: "行业/场景", blocking: false },
  { key: "standardPlan", label: "标准方案说明", blocking: false },
]

const asText = (value: unknown) => String(value ?? "").trim()

const readFiniteNumber = (...values: unknown[]) => {
  for (const value of values) {
    if (value === undefined || value === null || value === "") continue
    const numeric = Number(value)
    if (Number.isFinite(numeric)) return numeric
  }
  return null
}

const restoreTaxItem = (subject: IctSubjectDefinition, item: any) => {
  const incl = readFiniteNumber(item?.incl_tax, item?.incl) ?? 0
  const tax = readFiniteNumber(item?.tax_rate, item?.tax) ?? subject.defaultTaxRate
  const explicitExcl = readFiniteNumber(item?.excl_tax, item?.excl)
  const splitParts = restoreTaxSplitParts(item?.split_parts ?? item?.splitParts, incl, tax)
  const excl = splitParts
    ? roundMoneyHalfUp(splitParts.reduce((sum, part) => sum + part.excl, 0))
    : explicitExcl ?? exclFromIncl(incl, tax)
  const customSubjectName = asText(item?.customSubjectName ?? item?.custom_subject_name)
  const billingSubjectName = asText(item?.billingSubjectName ?? item?.billing_subject_name)

  return {
    incl,
    tax,
    excl,
    ...(customSubjectName ? { customSubjectName } : {}),
    ...(billingSubjectName ? { billingSubjectName } : {}),
    ...(splitParts ? { splitParts } : {}),
  }
}

const assignProjectSubject = (projectData: any, subject: IctSubjectDefinition, item: any) => {
  if (subject.groupId === "revIt") projectData.revenue.it[subject.key] = item
  else if (subject.groupId === "revCt") projectData.revenue.ct[subject.key] = item
  else if (subject.groupId === "revNonItCt") projectData.revenue.non_it_ct = item
  else if (subject.groupId === "costIt") projectData.cost.it[subject.key] = item
  else if (subject.groupId === "costCt") projectData.cost.ct[subject.key] = item
  else if (subject.groupId === "costMix") projectData.cost.mix[subject.key] = item
}

export const restoreSelectionProjectData = (
  project: Project,
  input: any,
  cashflowAssumptions?: any,
) => {
  const projectData: any = {
    basic: {
      proj_name: asText(input?.project_name) || project.name,
      customer_name: asText(input?.customer_name) || project.customer_name,
      project_years: Number(input?.project_years || cashflowAssumptions?.projectYears || project.project_years || 1),
    },
    revenue: { it: {}, ct: {}, non_it_ct: null },
    cost: { it: {}, ct: {}, mix: {} },
    selectionFee: {
      quote: input?.selection_fee_quote ?? "",
      markup: input?.selection_fee_markup ?? "",
      actualCost: input?.selection_fee_actual_cost ?? "",
      amount: input?.selection_fee_amount ?? "",
      limit: input?.selection_fee_limit ?? "",
      anchor: input?.selection_fee_anchor ?? "limit",
    },
  }

  const assumptionGroups: Record<string, any> = {
    revIt: cashflowAssumptions?.revIt,
    revCt: cashflowAssumptions?.revCt,
    revNonItCt: cashflowAssumptions?.revNonItCt,
    costIt: cashflowAssumptions?.costIt,
    costCt: cashflowAssumptions?.costCt,
    costMix: cashflowAssumptions?.costMix,
  }

  ICT_SUBJECT_DEFINITIONS.forEach(subject => {
    const assumptionGroup = assumptionGroups[subject.groupId]
    const assumptionItem = subject.groupId === "revNonItCt"
      ? assumptionGroup
      : assumptionGroup?.[subject.key]
    const rawItem = assumptionItem ?? input?.[subject.subjectCode]
    assignProjectSubject(projectData, subject, restoreTaxItem(subject, rawItem))
  })

  return projectData
}

export const readSelectionSharedFields = (templateState: any): SelectionSharedFields => {
  const filled = templateState?.filledDataJson || {}
  const form = filled.formData || {}
  return {
    winnerName: asText(form.gen_zx_winner_name),
    scope: asText(form.gen_zx_scope),
    industry: asText(form.gen_zx_industry),
    method: asText(form.gen_zx_method),
    rule: asText(form.gen_zx_rule),
    standardPlan: asText(form.gen_zx_std_plan),
    revenueCollection: asText(filled.revCollection),
    expenditurePayment: asText(filled.expPayment),
  }
}

export const validateSelectionProjectData = (projectData: any) => validateFinancialData(
  {
    it: projectData.revenue?.it || {},
    ct: projectData.revenue?.ct || {},
    non_it_ct: { item: projectData.revenue?.non_it_ct || {} },
  },
  {
    it: projectData.cost?.it || {},
    ct: projectData.cost?.ct || {},
    mix: projectData.cost?.mix || {},
  },
).errors

export const missingBenefitMetrics = (metrics: any) => [
  ["项目净现值率", metrics?.npv_rate],
  ["项目毛利率", metrics?.margin_rate],
  ["IT净现值率", metrics?.it_npv_rate],
].filter(([, value]) => value === undefined || value === null || value === "")
  .map(([label]) => String(label))

export const buildLiveSelectionResultProject = (options: {
  projectId: string
  projectData: any
  metrics: any
  projectBackground: string
  preSelectionCostIt: Record<string, any>
  preSchemeId: string
  preSchemeName: string
  postSchemeId: string
  postSchemeName: string
  sharedFields: SelectionSharedFields
}): SelectionResultBatchProject => ({
  projectId: options.projectId,
  projectName: options.projectData.basic?.proj_name || "",
  customerName: options.projectData.basic?.customer_name || "",
  postSchemeId: options.postSchemeId,
  postSchemeName: options.postSchemeName,
  preSchemeId: options.preSchemeId,
  preSchemeName: options.preSchemeName,
  projectData: options.projectData,
  preSelectionCostIt: options.preSelectionCostIt,
  metrics: options.metrics || {},
  projectBackground: options.projectBackground,
  sharedFields: options.sharedFields,
  validationErrors: validateSelectionProjectData(options.projectData),
  missingMetrics: missingBenefitMetrics(options.metrics),
})

const money = (value: unknown) => {
  try {
    return new Decimal(value === undefined || value === null || value === "" ? 0 : value as any)
  } catch {
    return new Decimal(0)
  }
}

export const formatSelectionMoney = (value: Decimal.Value) => money(value).toFixed(2)

const subjectLabel = (subject: IctSubjectDefinition, item: any) => {
  const resolved = resolveBillingSubjectPresentation(subject, item)
  const baseName = resolved.billingSubjectName || resolved.standardName
  const prefix = `${subject.documentPrefix}-`
  return baseName.startsWith(prefix) ? baseName : `${prefix}${baseName}`
}

const sumSubjects = (projectData: any, subjects: IctSubjectDefinition[], field: "excl" | "incl") =>
  subjects.reduce(
    (total, subject) => total.plus(money(getProjectDataSubjectItem(projectData, subject)?.[field])),
    new Decimal(0),
  )

export const getSelectionRenewalAmount = (projectData: any) =>
  money(projectData.cost?.ct?.renewal?.excl)

export const calculateSelectionApprovalAmount = (
  projectData: any,
  renewalDecision?: SelectionRenewalDecision,
) => {
  const itAmount = sumSubjects(projectData, IT_COST_SUBJECTS, "excl")
  const dedicatedLineAmount = ICT_SUBJECT_DEFINITIONS
    .filter(subject => subject.groupId === "costCt" && CT_DEDICATED_LINE_KEYS.has(subject.key))
    .reduce(
      (total, subject) => total.plus(money(getProjectDataSubjectItem(projectData, subject)?.excl)),
      new Decimal(0),
    )
  const renewalAmount = renewalDecision === "include" ? getSelectionRenewalAmount(projectData) : new Decimal(0)
  return itAmount.plus(dedicatedLineAmount).plus(renewalAmount)
}

const buildRowsForProject = (
  project: SelectionResultBatchProject,
  projectIndex: number,
  subjects: IctSubjectDefinition[],
  prefix: "A" | "B" | "C" | "D",
  sourceData: any,
) => {
  const rows: Record<string, string>[] = []
  subjects.forEach(subject => {
    const item = getProjectDataSubjectItem(sourceData, subject)
    const excl = money(item?.excl)
    if (excl.abs().lt(0.005)) return
    const isFirst = rows.length === 0
    if (prefix === "A") {
      rows.push({
        A_SEQ: isFirst ? String(projectIndex + 1) : "",
        A_NAME: isFirst ? project.projectName : "",
        A_FEE_TYPE: subjectLabel(subject, item),
        A_TAX_RATE: `${Number(item?.tax ?? subject.defaultTaxRate)}%`,
        A_LIMIT: excl.toFixed(2),
      })
    } else {
      rows.push({
        [`${prefix}_SEQ`]: isFirst ? String(projectIndex + 1) : "",
        [`${prefix}_NAME`]: isFirst ? project.projectName : "",
        [`${prefix}_TYPE`]: subjectLabel(subject, item),
        [`${prefix}_EXCL`]: excl.toFixed(2),
        [`${prefix}_TAX_RATE`]: `${Number(item?.tax ?? subject.defaultTaxRate)}%`,
        [`${prefix}_INCL`]: money(item?.incl).toFixed(2),
      })
    }
  })
  return rows
}

const aggregateSituation = (
  projects: SelectionResultBatchProject[],
  side: IctSubjectSide,
  total: Decimal,
) => {
  const labelTotals = new Map<string, Decimal>()
  projects.forEach(project => {
    ICT_SUBJECT_DEFINITIONS.filter(subject => subject.side === side).forEach(subject => {
      const item = getProjectDataSubjectItem(project.projectData, subject)
      const amount = money(item?.excl)
      if (amount.abs().lt(0.005)) return
      const label = subjectLabel(subject, item)
      labelTotals.set(label, (labelTotals.get(label) || new Decimal(0)).plus(amount))
    })
  })
  const action = side === "cost" ? "投入" : "收入"
  const details = [...labelTotals.entries()]
    .map(([label, amount]) => `${label}${action}${amount.toFixed(2)}元`)
    .join("，")
  return `总${action}${total.toFixed(2)}元${details ? `；其中${details}` : ""}。`
}

export function buildSelectionResultBatchModel(
  projects: SelectionResultBatchProject[],
  renewalDecisions: Record<string, SelectionRenewalDecision | undefined>,
): SelectionResultBatchModel {
  const tableA: Record<string, string>[] = []
  const tableB: Record<string, string>[] = []
  const tableC: Record<string, string>[] = []
  const tableD: Record<string, string>[] = []
  const tableE: Record<string, string>[] = []
  let totalLimitExcl = new Decimal(0)
  let totalWinnerExcl = new Decimal(0)
  let totalWinnerIncl = new Decimal(0)
  let totalCostExcl = new Decimal(0)
  let totalCostIncl = new Decimal(0)
  let totalRevenueExcl = new Decimal(0)
  let totalRevenueIncl = new Decimal(0)
  let approvalAmountExcl = new Decimal(0)
  const renewalProjects: SelectionResultBatchModel["renewalProjects"] = []

  const costSubjects = ICT_SUBJECT_DEFINITIONS.filter(subject => subject.side === "cost")
  const revenueSubjects = ICT_SUBJECT_DEFINITIONS.filter(subject => subject.side === "revenue")

  projects.forEach((project, projectIndex) => {
    const manualLimit = money(project.projectData.selectionFee?.limit)
    if (manualLimit.gt(0)) {
      const integrationSubject = IT_COST_SUBJECTS.find(subject => subject.key === "integration") || IT_COST_SUBJECTS[0]
      const integrationItem = project.projectData.cost?.it?.integration
      tableA.push({
        A_SEQ: String(projectIndex + 1),
        A_NAME: project.projectName,
        A_FEE_TYPE: subjectLabel(integrationSubject, integrationItem),
        A_TAX_RATE: `${Number(integrationItem?.tax ?? integrationSubject.defaultTaxRate)}%`,
        A_LIMIT: manualLimit.toFixed(2),
      })
      totalLimitExcl = totalLimitExcl.plus(manualLimit)
    } else {
      const preProjectData = {
        revenue: { it: {}, ct: {}, non_it_ct: null },
        cost: { it: project.preSelectionCostIt, ct: {}, mix: {} },
      }
      tableA.push(...buildRowsForProject(project, projectIndex, IT_COST_SUBJECTS, "A", preProjectData))
      totalLimitExcl = totalLimitExcl.plus(sumSubjects(preProjectData, IT_COST_SUBJECTS, "excl"))
    }

    tableB.push(...buildRowsForProject(project, projectIndex, IT_COST_SUBJECTS, "B", project.projectData))
    tableC.push(...buildRowsForProject(project, projectIndex, costSubjects, "C", project.projectData))
    tableD.push(...buildRowsForProject(project, projectIndex, revenueSubjects, "D", project.projectData))
    tableE.push({
      E_SEQ: String(projectIndex + 1),
      E_NAME: project.projectName,
      E_NPV_RATE: formatSelectionMetric(project.metrics?.npv_rate),
      E_MARGIN: formatSelectionMetric(project.metrics?.margin_rate),
      E_IT_NPV: formatSelectionMetric(project.metrics?.it_npv_rate),
    })

    totalWinnerExcl = totalWinnerExcl.plus(sumSubjects(project.projectData, IT_COST_SUBJECTS, "excl"))
    totalWinnerIncl = totalWinnerIncl.plus(sumSubjects(project.projectData, IT_COST_SUBJECTS, "incl"))
    totalCostExcl = totalCostExcl.plus(sumSubjects(project.projectData, costSubjects, "excl"))
    totalCostIncl = totalCostIncl.plus(sumSubjects(project.projectData, costSubjects, "incl"))
    totalRevenueExcl = totalRevenueExcl.plus(sumSubjects(project.projectData, revenueSubjects, "excl"))
    totalRevenueIncl = totalRevenueIncl.plus(sumSubjects(project.projectData, revenueSubjects, "incl"))
    approvalAmountExcl = approvalAmountExcl.plus(
      calculateSelectionApprovalAmount(project.projectData, renewalDecisions[project.projectId]),
    )

    const renewalAmount = getSelectionRenewalAmount(project.projectData)
    if (renewalAmount.gt(0)) {
      renewalProjects.push({ projectId: project.projectId, projectName: project.projectName, amountExcl: renewalAmount })
    }
  })

  return {
    tableA,
    tableB,
    tableC,
    tableD,
    tableE,
    totalLimitExcl,
    totalWinnerExcl,
    totalWinnerIncl,
    totalCostExcl,
    totalCostIncl,
    totalRevenueExcl,
    totalRevenueIncl,
    approvalAmountExcl,
    investmentSituation: aggregateSituation(projects, "cost", totalCostExcl),
    revenueSituation: aggregateSituation(projects, "revenue", totalRevenueExcl),
    renewalProjects,
  }
}

export const formatSelectionMetric = (value: unknown) => {
  const numeric = Number(value)
  return Number.isFinite(numeric) ? `${new Decimal(numeric).mul(100).toFixed(2)}%` : "--"
}

export const detectSelectionSharedConflicts = (
  projects: SelectionResultBatchProject[],
): SelectionSharedConflict[] => SHARED_FIELD_CONFIG.flatMap(config => {
  const values = projects
    .map(project => ({ projectName: project.projectName, value: asText(project.sharedFields[config.key]) }))
    .filter(entry => entry.value)
  const uniqueValues = new Set(values.map(entry => entry.value))
  return uniqueValues.size > 1 ? [{ ...config, values }] : []
})

export const defaultSelectionBatchName = (projects: SelectionResultBatchProject[]) => {
  if (projects.length === 0) return ""
  if (projects.length === 1) return projects[0].projectName
  const firstName = projects[0].projectName.replace(/(?:ICT)?项目$/i, "").trim() || projects[0].projectName
  return `${firstName}等${projects.length}个ICT项目`
}

export const selectionBatchItContent = (projects: SelectionResultBatchProject[]) =>
  projects.map(project => project.projectName).join("、")

export const selectionBatchCtContent = (projects: SelectionResultBatchProject[]) => {
  const ctSubjects = ICT_SUBJECT_DEFINITIONS.filter(subject => subject.documentPrefix === "CT")
  const names = projects
    .filter(project => ctSubjects.some(subject => money(getProjectDataSubjectItem(project.projectData, subject)?.excl).abs().gte(0.005)))
    .map(project => project.projectName)
  return names.length ? names.join("、") : "无"
}
