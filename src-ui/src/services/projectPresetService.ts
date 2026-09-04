import { invoke } from "@tauri-apps/api/core";

export type ProjectPresetValueType =
  | "text"
  | "long_text"
  | "dictionary_value"
  | "boolean";

export type ProjectPresetSourceType =
  | "manual"
  | "from_project"
  | "preset_item"
  | "dictionary";

export interface ProjectPresetTemplateEntry {
  id: string;
  templateId: string;
  fieldKey: string;
  value: unknown;
  valueType: ProjectPresetValueType;
  sourceType: ProjectPresetSourceType;
  sortOrder: number;
  createdAt: string;
  updatedAt: string;
}

export interface ProjectPresetTemplate {
  id: string;
  scope: "workspace";
  name: string;
  description?: string | null;
  category: string;
  tags: string[];
  enabled: boolean;
  createdAt: string;
  updatedAt: string;
  entries: ProjectPresetTemplateEntry[];
}

export interface ProjectPresetTemplateEntryInput {
  id?: string;
  fieldKey: string;
  value: unknown;
  valueType: ProjectPresetValueType;
  sourceType?: ProjectPresetSourceType;
  sortOrder?: number;
}

export interface ProjectPresetTemplateInput {
  id?: string;
  scope?: "workspace";
  name: string;
  description?: string | null;
  category?: string;
  tags?: string[];
  enabled?: boolean;
  entries: ProjectPresetTemplateEntryInput[];
}

export const projectPresetService = {
  list(includeDisabled = false): Promise<ProjectPresetTemplate[]> {
    return invoke<ProjectPresetTemplate[]>("list_project_preset_templates", {
      includeDisabled,
    });
  },

  save(template: ProjectPresetTemplateInput): Promise<ProjectPresetTemplate> {
    return invoke<ProjectPresetTemplate>("save_project_preset_template", { template });
  },

  setEnabled(id: string, enabled: boolean): Promise<ProjectPresetTemplate> {
    return invoke<ProjectPresetTemplate>("set_project_preset_template_enabled", {
      id,
      enabled,
    });
  },

  delete(id: string): Promise<void> {
    return invoke<void>("delete_project_preset_template", { id });
  },
};
