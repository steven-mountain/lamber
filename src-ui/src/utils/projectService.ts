import { invoke } from "@tauri-apps/api/core";

export interface IctItem {
  incl_tax: string;
  tax_rate: string;
  custom_subject_name?: string;
  billing_subject_name?: string;
}

export interface CashflowSegment {
  id: string;
  name: string;
  value: number;
  revenueValue: number;
  revenueTax: number;
  revenueScope: string;
  costValue: number;
  costTax: number;
  costScope: string;
  startYear: number;
  serviceYears: number;
  revenueMode: string;
  costMode: string;
  revenueAnnualValues: number[];
  costAnnualValues: number[];
}

export interface BalanceSubjectRef {
  subject_code: string;
  group_id: string;
  key: string;
}

export interface BalanceAllocationRulePayload {
  enabled: boolean;
  total_incl_amount: number | null;
  balancing_subject: BalanceSubjectRef | null;
}

export interface IctInput {
  project_name: string;
  customer_name?: string;
  property_rights: string;
  discount_rate: string;
  project_years?: number;
  cashflow_model?: string;
  cashflow_calculation_source?: "legacy_model" | "subject_funding_plans";
  subject_funding_plan_migration_version?: number;
  cashflow_segment_value_mode?: string;
  cashflow_segments?: CashflowSegment[];
  ignore_tail_difference?: boolean;
  tail_difference_value?: string;
  revenue_balance_rule?: BalanceAllocationRulePayload;
  investment_balance_rule?: BalanceAllocationRulePayload;
  rev_distribution: number[];
  cost_distribution: number[];
  rev_cashflow_excl?: string[] | null;
  cost_cashflow_excl?: string[] | null;
  it_rev_cashflow_excl?: string[] | null;
  it_cost_cashflow_excl?: string[] | null;
  [key: string]: any;
}

export interface IctCashflowRow {
  year: number;
  cash_in: string;
  cash_out: string;
  net_cash: string;
  cum_net_cash: string;
  pv: string;
  cum_pv: string;
}

export interface IctResult {
  npv: string;
  npv_rate: string;
  margin_rate: string;
  dynamic_payback: string;
  irr: string;
  it_npv: string;
  it_npv_rate: string;
  it_margin_rate: string;
  cashflow: IctCashflowRow[];
}

export interface ProjectLog {
  id: string;
  timestamp: string;
  description: string;
}

export interface SummaryMetrics {
  margin_rate: string;
  npv: string;
  npv_rate: string;
  irr: string;
  dynamic_payback: string;
  risk_level: string;
}

export interface Project {
  id: string;
  name: string;
  customer_name: string;
  status: string; // User-editable lifecycle tag, defaults to "需求导入"
  benefit_status: "not_started" | "normal" | "outdated";
  default_scheme_id?: string | null;
  created_at: string;
  updated_at: string;

  total_revenue_incl: number;
  total_cost_incl: number;
  project_years: number;
  discount_rate: number;
  cashflow_model: string;

  summary_metrics?: SummaryMetrics | null;
  folder_path?: string | null;
  main_document_path?: string | null;
  main_budget_file_path?: string | null;
  note?: string | null;
  logs: ProjectLog[];

  // Workspace integration fields
  folder_name?: string | null;
  relative_path?: string | null;
  progress?: number;
  deadline?: string | null;
  linked_folder_type?: "none" | "internal" | "external" | null;
  linked_folder_relative_path?: string | null;
  linked_folder_external_path?: string | null;
  directoryExists?: boolean;
}

export interface UnregisteredProject {
  projectId: string;
  name: string;
  relativePath: string;
  folderName: string;
}

export interface WorkspaceProjectInfo {
  project: Project;
  directoryExists: boolean;
}

export interface BenefitAnalysisScheme {
  id: string;
  project_id: string;
  name: string;
  created_at: string;
  updated_at: string;
}

export interface BenefitAnalysisSnapshot {
  id: string;
  scheme_id: string;
  project_id: string;
  version: number;
  input_params: IctInput;
  output_metrics: IctResult;
  fingerprint: string;
  created_at: string;
}

export const projectService = {
  async getProjects(): Promise<Project[]> {
    return invoke<Project[]>("get_projects");
  },

  async getProject(id: string): Promise<Project | null> {
    return invoke<Project | null>("get_project", { id });
  },

  async createProject(name: string, customerName: string): Promise<Project> {
    return invoke<Project>("create_project", { name, customerName });
  },

  async updateProject(project: Project): Promise<Project> {
    return invoke<Project>("update_project", { project });
  },

  async deleteProject(id: string): Promise<void> {
    return invoke<void>("delete_project", { id });
  },

  async deleteBenefitScheme(projectId: string, schemeId: string): Promise<Project> {
    return invoke<Project>("delete_benefit_scheme", { projectId, schemeId });
  },

  async getSchemes(projectId: string): Promise<BenefitAnalysisScheme[]> {
    return invoke<BenefitAnalysisScheme[]>("get_schemes", { projectId });
  },

  async getSnapshots(schemeId: string): Promise<BenefitAnalysisSnapshot[]> {
    return invoke<BenefitAnalysisSnapshot[]>("get_snapshots", { schemeId });
  },

  async saveBenefitScheme(
    projectId: string,
    schemeIdOpt: string | null,
    schemeName: string,
    inputParams: any,
    outputMetrics: any,
    isSaveAsNew: boolean
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

  async saveTemplateAsset(
    projectId: string,
    templateName: string,
    assetType: string,
    usage: string | null,
    originalFileName: string | null,
    base64Data: string,
    width: number | null,
    height: number | null
  ): Promise<string> {
    return invoke<string>("save_template_asset", {
      projectId,
      templateName,
      assetType,
      usage,
      originalFileName,
      base64Data,
      width,
      height,
    });
  },

  async getTemplateAssetPath(assetId: string): Promise<string> {
    return invoke<string>("get_template_asset_path", { assetId });
  },

  async deleteTemplateAsset(assetId: string): Promise<void> {
    return invoke<void>("delete_template_asset", { assetId });
  },

  async cleanupOrphanTemplateAssets(projectId: string): Promise<[number, string[]]> {
    return invoke<[number, string[]]>("cleanup_orphan_template_assets", { projectId });
  },

  async getProjectSetting(projectId: string, key: string): Promise<string | null> {
    return invoke<string | null>("get_project_setting", { projectId, key });
  },

  async saveProjectSetting(projectId: string, key: string, value: string): Promise<void> {
    return invoke<void>("save_project_setting", { projectId, key, value });
  },

  async createProjectInWorkspace(name: string, customerName: string, projectPresetTemplateId?: string | null): Promise<Project> {
    return invoke<Project>("create_project_in_workspace", {
      name,
      customerName,
      projectPresetTemplateId: projectPresetTemplateId || null,
    });
  },

  async listWorkspaceProjects(): Promise<WorkspaceProjectInfo[]> {
    return invoke<WorkspaceProjectInfo[]>("list_workspace_projects");
  },

  async inspectWorkspaceProjects(): Promise<UnregisteredProject[]> {
    return invoke<UnregisteredProject[]>("inspect_workspace_projects");
  },
};
