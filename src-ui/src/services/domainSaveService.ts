import { invoke } from "@tauri-apps/api/core";
import type {
  BenefitAnalysisScheme,
  BenefitAnalysisSnapshot,
  IctInput,
  IctResult,
  Project,
} from "../utils/projectService";

export interface ProjectDetailPatch {
  name?: string;
  customerName?: string;
  status?: string;
  progress?: number;
  deadline?: string | null;
  note?: string | null;
  linkedFolderType?: string | null;
  linkedFolderRelativePath?: string | null;
  linkedFolderExternalPath?: string | null;
}

export interface LifecycleStatePayload {
  profileJson: Record<string, unknown>;
  parametersJson: Record<string, unknown>;
  backgroundJson: Record<string, unknown>;
  inputPayloadJson: Record<string, unknown>;
}

export interface CashflowStatePayload {
  cashflowModel?: string | null;
  paymentModelJson: Record<string, unknown>;
  yearlyCashflowJson: Record<string, unknown>;
  sectorCashflowJson: Record<string, unknown>;
  assumptionsJson: Record<string, unknown>;
  metricsJson: Record<string, unknown>;
}

export interface TemplateStatePayload {
  templateName?: string | null;
  templateType?: string | null;
  templatePath?: string | null;
  templatePathType?: string | null;
  filledDataJson: Record<string, unknown>;
  fieldMappingJson: Record<string, unknown>;
  outputConfigJson: Record<string, unknown>;
}

export interface TemplateAssetPayload {
  assetType: string;
  usage?: string | null;
  originalFileName?: string | null;
  base64Data: string;
  width?: number | null;
  height?: number | null;
}

export interface StoredTemplateState {
  id: string;
  projectId: string;
  templateId: string;
  templateName?: string | null;
  filledDataJson: Record<string, unknown>;
  fieldMappingJson: Record<string, unknown>;
  outputConfigJson: Record<string, unknown>;
  source: string;
}

export interface ProjectFullState {
  project: Project;
  lifecycleState?: any | null;
  cashflowState?: any | null;
  schemes: BenefitAnalysisScheme[];
  latestSnapshot?: BenefitAnalysisSnapshot | null;
  templateStates: StoredTemplateState[];
  templateAssets: any[];
  legacyLifecycleInput?: IctInput | null;
  legacyCashflowMetrics?: IctResult | null;
}

export const domainSaveService = {
  saveProjectDetail(projectId: string, patch: ProjectDetailPatch): Promise<Project> {
    return invoke<Project>("save_project_detail", { projectId, patch });
  },

  getProjectDetail(projectId: string): Promise<Project> {
    return invoke<Project>("get_project_detail", { projectId });
  },

  saveLifecycleState(projectId: string, lifecycleState: LifecycleStatePayload): Promise<any> {
    return invoke("save_lifecycle_state", { projectId, lifecycleState });
  },

  loadLifecycleState(projectId: string): Promise<any | null> {
    return invoke("get_lifecycle_state", { projectId });
  },

  saveCashflowState(projectId: string, cashflowState: CashflowStatePayload): Promise<any> {
    return invoke("save_cashflow_state", { projectId, cashflowState });
  },

  loadCashflowState(projectId: string): Promise<any | null> {
    return invoke("get_cashflow_state", { projectId });
  },

  saveBenefitAnalysis(
    projectId: string,
    schemeIdOpt: string | null,
    schemeName: string,
    inputParams: IctInput | Record<string, unknown>,
    outputMetrics: IctResult | Record<string, unknown>,
    isSaveAsNew: boolean,
  ): Promise<Project> {
    return invoke<Project>("save_benefit_analysis", {
      projectId,
      schemeIdOpt,
      schemeName,
      inputParams,
      outputMetrics,
      isSaveAsNew,
    });
  },

  loadBenefitSchemes(projectId: string): Promise<BenefitAnalysisScheme[]> {
    return invoke<BenefitAnalysisScheme[]>("get_benefit_schemes", { projectId });
  },

  saveTemplateState(projectId: string, templateId: string, templateState: TemplateStatePayload): Promise<StoredTemplateState> {
    return invoke<StoredTemplateState>("save_template_state", { projectId, templateId, templateState });
  },

  loadTemplateState(projectId: string, templateId: string): Promise<StoredTemplateState | null> {
    return invoke<StoredTemplateState | null>("get_template_state", { projectId, templateId });
  },

  loadTemplateStates(projectId: string): Promise<StoredTemplateState[]> {
    return invoke<StoredTemplateState[]>("list_template_states", { projectId });
  },

  loadTemplateAssets(projectId: string, templateId?: string | null): Promise<any[]> {
    return invoke<any[]>("list_template_assets", { projectId, templateId: templateId || null });
  },

  saveTemplateAsset(projectId: string, templateId: string, asset: TemplateAssetPayload): Promise<string> {
    return invoke<string>("save_template_asset", {
      projectId,
      templateName: templateId,
      assetType: asset.assetType,
      usage: asset.usage || null,
      originalFileName: asset.originalFileName || null,
      base64Data: asset.base64Data,
      width: asset.width ?? null,
      height: asset.height ?? null,
    });
  },

  loadProjectFullState(projectId: string): Promise<ProjectFullState> {
    return invoke<ProjectFullState>("get_project_full_state", { projectId });
  },
};
