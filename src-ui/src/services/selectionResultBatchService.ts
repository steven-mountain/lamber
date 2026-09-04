import { domainSaveService } from "./domainSaveService"
import {
  projectService,
  type BenefitAnalysisScheme,
  type BenefitAnalysisSnapshot,
} from "../utils/projectService"
import {
  missingBenefitMetrics,
  readSelectionSharedFields,
  restoreSelectionProjectData,
  validateSelectionProjectData,
  type SelectionResultBatchProject,
  type SelectionResultCandidate,
} from "../lib/selectionResultBatch"

const latestScheme = (schemes: BenefitAnalysisScheme[], stage: string) =>
  [...schemes]
    .filter(scheme => scheme.stage === stage)
    .sort((left, right) => (right.updated_at || "").localeCompare(left.updated_at || ""))[0] || null

const latestSnapshot = (snapshots: BenefitAnalysisSnapshot[]) =>
  snapshots.reduce<BenefitAnalysisSnapshot | null>(
    (latest, snapshot) => !latest || snapshot.version > latest.version ? snapshot : latest,
    null,
  )

const hasRecordValues = (value: unknown) => Boolean(
  value
  && typeof value === "object"
  && !Array.isArray(value)
  && Object.keys(value as Record<string, unknown>).length > 0,
)

export async function loadSelectionResultCandidates(): Promise<SelectionResultCandidate[]> {
  const projects = (await projectService.getProjects())
    .filter(project => project.project_type === "ict")
    .sort((left, right) => left.name.localeCompare(right.name, "zh-CN"))

  return Promise.all(projects.map(async project => {
    try {
      const schemes = await projectService.getSchemes(project.id)
      const preScheme = latestScheme(schemes, "pre_selection")
      const postScheme = latestScheme(schemes, "post_selection")
      const reason = !postScheme
        ? "缺少甄选后方案"
        : !preScheme
          ? "缺少甄选前方案"
          : ""
      return { project, preScheme, postScheme, eligible: !reason, reason }
    } catch (error) {
      return {
        project,
        preScheme: null,
        postScheme: null,
        eligible: false,
        reason: `读取方案失败：${String(error)}`,
      }
    }
  }))
}

const loadSchemeState = async (projectId: string, scheme: BenefitAnalysisScheme) => {
  const [lifecycleState, cashflowState, snapshots] = await Promise.all([
    domainSaveService.loadLifecycleState(projectId, scheme.id).catch(() => null),
    domainSaveService.loadCashflowState(projectId, scheme.id).catch(() => null),
    projectService.getSnapshots(scheme.id).catch(() => []),
  ])
  const snapshot = latestSnapshot(snapshots)
  return {
    input: hasRecordValues(lifecycleState?.inputPayloadJson)
      ? lifecycleState.inputPayloadJson
      : snapshot?.input_params || null,
    lifecycleState,
    cashflowState,
    metrics: hasRecordValues(cashflowState?.metricsJson)
      ? cashflowState.metricsJson
      : snapshot?.output_metrics || null,
  }
}

export async function loadSelectionResultBatchProject(
  candidate: SelectionResultCandidate,
  templateId: string,
): Promise<SelectionResultBatchProject> {
  if (!candidate.preScheme || !candidate.postScheme) {
    throw new Error(`${candidate.project.name}：${candidate.reason || "甄选方案不完整"}`)
  }

  const [preState, postState, templateState] = await Promise.all([
    loadSchemeState(candidate.project.id, candidate.preScheme),
    loadSchemeState(candidate.project.id, candidate.postScheme),
    domainSaveService.loadTemplateState(candidate.project.id, templateId).catch(() => null),
  ])
  if (!preState.input) throw new Error(`${candidate.project.name}：甄选前方案没有可用测算数据`)
  if (!postState.input) throw new Error(`${candidate.project.name}：甄选后方案没有可用测算数据`)

  const preProjectData = restoreSelectionProjectData(
    candidate.project,
    preState.input,
    preState.cashflowState?.assumptionsJson,
  )
  const projectData = restoreSelectionProjectData(
    candidate.project,
    postState.input,
    postState.cashflowState?.assumptionsJson,
  )
  const metrics = postState.metrics || {}

  return {
    projectId: candidate.project.id,
    projectName: projectData.basic?.proj_name || candidate.project.name,
    customerName: projectData.basic?.customer_name || candidate.project.customer_name,
    postSchemeId: candidate.postScheme.id,
    postSchemeName: candidate.postScheme.name,
    preSchemeId: candidate.preScheme.id,
    preSchemeName: candidate.preScheme.name,
    projectData,
    preSelectionCostIt: preProjectData.cost?.it || {},
    metrics,
    projectBackground: String(
      postState.lifecycleState?.backgroundJson?.projectBackground
        ?? postState.input?.project_background
        ?? "",
    ).trim(),
    sharedFields: readSelectionSharedFields(templateState),
    validationErrors: validateSelectionProjectData(projectData),
    missingMetrics: missingBenefitMetrics(metrics),
  }
}
