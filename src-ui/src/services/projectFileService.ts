import { invoke } from "@tauri-apps/api/core";

export interface ProjectFile {
  id: string;
  projectId: string;
  fileName: string;
  filePath: string;
  originalPath?: string;
  managedPath?: string;
  fileType: 'word' | 'excel' | 'pdf' | 'ppt' | 'image' | 'other';
  extension: string;
  size: number;
  exists: boolean;
  lastScannedAt?: string;
  modifiedAt: string;
  storageMode: 'linked' | 'copied';
  isMainDocument: boolean;
  isMainBudgetFile: boolean;
  note?: string;
  createdAt: string;
  updatedAt: string;
}

export const projectFileService = {
  async getProjectFiles(projectId: string): Promise<ProjectFile[]> {
    return invoke<ProjectFile[]>("get_project_files", { projectId });
  },

  async bindProjectFolder(projectId: string, folderPath: string): Promise<void> {
    return invoke<void>("bind_project_folder", { projectId, folderPath });
  },

  async createProjectFolder(parentPath: string, folderName: string): Promise<string> {
    return invoke<string>("create_project_folder", { parentPath, folderName });
  },

  async scanProjectFolder(projectId: string, recursive: boolean = false): Promise<ProjectFile[]> {
    return invoke<ProjectFile[]>("scan_project_folder", { projectId, recursive });
  },

  async addProjectFile(projectId: string, srcPath: string, storageMode: 'linked' | 'copied'): Promise<ProjectFile> {
    return invoke<ProjectFile>("add_project_file", { projectId, srcPath, storageMode });
  },

  async removeProjectFileRecord(projectId: string, fileId: string): Promise<void> {
    return invoke<void>("remove_project_file_record", { projectId, fileId });
  },

  async deleteManagedProjectFile(projectId: string, fileId: string): Promise<void> {
    return invoke<void>("delete_managed_project_file", { projectId, fileId });
  },

  async markMainDocument(projectId: string, fileId: string | null): Promise<void> {
    return invoke<void>("mark_main_document", { projectId, fileId });
  },

  async markMainBudgetFile(projectId: string, fileId: string | null): Promise<void> {
    return invoke<void>("mark_main_budget_file", { projectId, fileId });
  },

  async openProjectFolder(projectId: string): Promise<void> {
    return invoke<void>("open_project_folder", { projectId });
  },

  async openProjectFile(fileId: string): Promise<void> {
    return invoke<void>("open_project_file", { fileId });
  },

  async revealProjectFile(fileId: string): Promise<void> {
    return invoke<void>("reveal_project_file", { fileId });
  },

  async selectLocalFolder(): Promise<string | null> {
    return invoke<string | null>("select_local_folder");
  },

  async selectLocalFile(title: string, extensions?: string[]): Promise<string | null> {
    return invoke<string | null>("select_local_file", { title, extensions });
  },

  async unbindProjectFolder(projectId: string): Promise<void> {
    return invoke<void>("unbind_project_folder", { projectId });
  },

  async parseBenefitExcel(filePath: string): Promise<ExcelParsedData> {
    return invoke<ExcelParsedData>("parse_benefit_excel", { filePath });
  }
};

export interface ExcelParsedData {
  project_name: string;
  customer_name?: string;
  total_income_incl: number;
  total_cost_incl: number;
  target_margin: number;
  target_npv: number;
  project_years: number;
  discount_rate?: number;
  ct_name: string;
  ct_income_incl: number;
  it_tax: number;
  ct_tax: number;
  payment_collect: string;
  payment_pay: string;
  items?: Record<string, {
    incl_tax: number;
    excl_tax: number;
    tax_rate: number;
  }>;
}
