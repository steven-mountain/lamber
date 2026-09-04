import {
  getPresetFieldDefinition,
  type PresetFieldDefinition,
} from "./presetFieldKeys";
import type {
  ProjectPresetSourceType,
  ProjectPresetValueType,
} from "../services/projectPresetService";

export interface ProjectPresetFieldBinding {
  fieldKey: string;
  value: unknown;
  valueType: ProjectPresetValueType;
  sourceType?: ProjectPresetSourceType;
  apply: (value: unknown) => void;
}

export interface ProjectPresetFieldDisplay {
  definition: PresetFieldDefinition;
  canApply: boolean;
  reason?: string;
}

export function isProjectPresetFieldAllowed(fieldKey: string): boolean {
  const definition = getPresetFieldDefinition(fieldKey);
  if (!definition) return false;
  return definition.presetEligible || Boolean(definition.dictionaryKey);
}

export function getProjectPresetValueType(
  fieldKey: string,
): ProjectPresetValueType {
  const definition = getPresetFieldDefinition(fieldKey);
  if (definition?.dictionaryKey) return "dictionary_value";
  if (definition?.fieldType === "checkbox") return "boolean";
  if (definition?.fieldType === "long_text") return "long_text";
  return "text";
}

export function isProjectPresetValueEmpty(value: unknown): boolean {
  if (value === null || value === undefined) return true;
  if (typeof value === "string") return value.trim().length === 0;
  return false;
}

export function summarizeProjectPresetValue(value: unknown): string {
  if (typeof value === "boolean") return value ? "是" : "否";
  const text = typeof value === "string" ? value : JSON.stringify(value);
  return text.length > 80 ? `${text.slice(0, 80)}...` : text;
}

export function getProjectPresetFieldDisplay(
  fieldKey: string,
): ProjectPresetFieldDisplay | null {
  const definition = getPresetFieldDefinition(fieldKey);
  if (!definition) return null;
  const canApply = isProjectPresetFieldAllowed(fieldKey);
  return {
    definition,
    canApply,
    reason: canApply ? undefined : "该字段不允许纳入项目预设",
  };
}
