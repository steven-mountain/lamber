import { invoke } from "@tauri-apps/api/core";

export interface AiAgentModelOption {
  id: string;
  name: string;
  supportsImages: boolean;
}

export interface AiAgentSettings {
  model: string;
  baseUrl: string;
  hasApiKey: boolean;
  models: AiAgentModelOption[];
}

export interface AiAgentSettingsUpdate {
  model: string;
  baseUrl: string;
  apiKey?: string;
  clearApiKey: boolean;
}

export function getAiAgentSettings() {
  return invoke<AiAgentSettings>("ai_get_settings");
}

export function saveAiAgentSettings(update: AiAgentSettingsUpdate) {
  return invoke<AiAgentSettings>("ai_save_settings", { update });
}
