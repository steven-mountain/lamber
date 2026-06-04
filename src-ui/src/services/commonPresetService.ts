import { invoke } from "@tauri-apps/api/core";
import type { CommonPresetKind } from "../lib/presetFieldKeys";

export interface CommonPreset {
  id: string;
  scope: "workspace" | "user";
  kind: CommonPresetKind;
  category: string;
  name: string;
  content: string;
  tags: string[];
  applicableFieldKeys: string[];
  usageCount: number;
  lastUsedAt?: string | null;
  enabled: boolean;
  createdAt: string;
  updatedAt: string;
}

export interface CommonPresetInput {
  id?: string | null;
  scope?: "workspace" | "user";
  kind: CommonPresetKind;
  category: string;
  name: string;
  content: string;
  tags?: string[];
  applicableFieldKeys?: string[];
  enabled?: boolean;
}

export interface CommonPresetFilter {
  kind?: CommonPresetKind | null;
  category?: string | null;
  fieldKey?: string | null;
  includeDisabled?: boolean;
  sortBy?: "recent" | "usage";
}

export const commonPresetService = {
  list(filter: CommonPresetFilter = {}): Promise<CommonPreset[]> {
    return invoke<CommonPreset[]>("list_common_presets", { filter });
  },

  save(preset: CommonPresetInput): Promise<CommonPreset> {
    return invoke<CommonPreset>("save_common_preset", { preset });
  },

  setEnabled(id: string, enabled: boolean): Promise<CommonPreset> {
    return invoke<CommonPreset>("set_common_preset_enabled", { id, enabled });
  },

  delete(id: string): Promise<void> {
    return invoke<void>("delete_common_preset", { id });
  },

  markUsed(id: string): Promise<CommonPreset> {
    return invoke<CommonPreset>("mark_common_preset_used", { id });
  },
};
