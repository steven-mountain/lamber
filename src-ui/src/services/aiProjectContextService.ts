import { invoke } from "@tauri-apps/api/core";

export type AiProjectContextSource =
  | "overview"
  | "lifecycle"
  | "cashflow"
  | "benefit"
  | "templates"
  | "template_detail"
  | "files";

export interface AiProjectContextRequest {
  projectId: string;
  requestedSources?: AiProjectContextSource[];
  activeTemplateId?: string | null;
}

export interface AiProjectContextBundle {
  projectId: string;
  projectName: string;
  overview: AiProjectOverview;
  lifecycle?: AiLifecycleContext | null;
  cashflow?: AiCashflowContext | null;
  benefit?: AiBenefitContext | null;
  templates?: AiTemplateContextSummary[] | null;
  templateDetail?: AiTemplateDetailContext | null;
  files?: AiFileContextSummary | null;
  sources: AiContextSourceMeta[];
  warnings: string[];
}

export interface AiWorkspaceProjectIndexItem {
  projectId: string;
  projectName: string;
  customerName?: string | null;
  status?: string | null;
  updatedAt?: string | null;
  hasLifecycleState: boolean;
  hasCashflowState: boolean;
  hasTemplateState: boolean;
  templateNames: string[];
  hasBenefitSchemes: boolean;
}

export interface AiProjectOverview {
  name: string;
  customerName?: string | null;
  status?: string | null;
  phase?: string | null;
  deadline?: string | null;
  description?: string | null;
  progress?: number | null;
  benefitStatus?: string | null;
  folderLinked: boolean;
  updatedAt?: string | null;
}

export interface AiLifecycleContext {
  hasSavedState: boolean;
  lifecycleVersion?: number | null;
  updatedAt?: string | null;
  summaryJson: Record<string, unknown>;
  profileJson?: unknown | null;
  parametersJson?: unknown | null;
  backgroundJson?: unknown | null;
  inputPayloadJson?: unknown | null;
}

export interface AiCashflowContext {
  hasSavedState: boolean;
  cashflowVersion?: number | null;
  cashflowModel?: string | null;
  hasYearlyCashflow: boolean;
  yearCount?: number | null;
  updatedAt?: string | null;
  summaryJson: Record<string, unknown>;
  paymentModelJson?: unknown | null;
  yearlyCashflowJson?: unknown | null;
  sectorCashflowJson?: unknown | null;
  assumptionsJson?: unknown | null;
  metricsJson?: unknown | null;
}

export interface AiBenefitContext {
  schemeCount: number;
  defaultScheme?: AiBenefitSchemeSummary | null;
  latestScheme?: AiBenefitSchemeSummary | null;
  latestSnapshot?: AiBenefitSnapshotSummary | null;
  projectSummaryMetrics?: unknown | null;
}

export interface AiBenefitSchemeSummary {
  id: string;
  name: string;
  updatedAt?: string | null;
  isDefault: boolean;
}

export interface AiBenefitSnapshotSummary {
  id: string;
  schemeId: string;
  version: number;
  createdAt?: string | null;
  outputMetricsSummary?: unknown | null;
  inputParams?: unknown | null;
}

export interface AiTemplateContextSummary {
  templateId: string;
  templateName?: string | null;
  hasSavedState: boolean;
  updatedAt?: string | null;
  fieldCount?: number | null;
  assetCount?: number | null;
  source: string;
}

export interface AiTemplateDetailContext {
  projectId: string;
  templateId: string;
  templateName?: string | null;
  source: string;
  hasSavedState: boolean;
  updatedAt?: string | null;
  fields: Record<string, unknown>;
  fieldMapping?: unknown | null;
  outputConfig?: unknown | null;
  assets: AiTemplateAssetReference[];
  warnings: string[];
}

export interface AiTemplateAssetReference {
  assetId: string;
  fieldKey?: string | null;
  fileName?: string | null;
  mimeType?: string | null;
  fileSize: number;
  width?: number | null;
  height?: number | null;
  exists?: boolean | null;
  updatedAt?: string | null;
}

export interface AiTemplateAssetImageInput {
  id: string;
  projectId: string;
  name: string;
  mimeType: string;
  size: number;
  width?: number | null;
  height?: number | null;
  dataUrl: string;
  source: string;
}

export interface AiFileContextSummary {
  totalFiles: number;
  existingFiles: number;
  missingFiles: number;
  fileTypeCounts: AiNamedCount[];
  storageModeCounts: AiNamedCount[];
  mainDocumentCount: number;
  mainBudgetFileCount: number;
  files?: AiProjectFileSummary[] | null;
}

export interface AiProjectFileSummary {
  id: string;
  fileName: string;
  fileType: string;
  extension: string;
  size: number;
  exists: boolean;
  storageMode: string;
  isMainDocument: boolean;
  isMainBudgetFile: boolean;
  fileRole?: string | null;
  modifiedAt?: string | null;
  updatedAt?: string | null;
}

export interface AiNamedCount {
  name: string;
  count: number;
}

export interface AiContextSourceMeta {
  sourceType: string;
  sourceId?: string | null;
  updatedAt?: string | null;
}

export function buildAiProjectContext(
  request: AiProjectContextRequest,
): Promise<AiProjectContextBundle> {
  return invoke<AiProjectContextBundle>("build_ai_project_context", { request });
}

export function listAiWorkspaceProjects(): Promise<AiWorkspaceProjectIndexItem[]> {
  return invoke<AiWorkspaceProjectIndexItem[]>("list_ai_workspace_projects");
}

export function loadAiTemplateAsset(projectId: string, assetId: string): Promise<AiTemplateAssetImageInput> {
  return invoke<AiTemplateAssetImageInput>("load_ai_template_asset", {
    request: { projectId, assetId },
  });
}
