import { webBusinessTransport } from '../services/webBusinessTransport';
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
  const remote = webBusinessTransport();
  if (remote) return remote<AiAgentSettings>("settings", {});
  return invoke<AiAgentSettings>("ai_get_settings");
}

export function saveAiAgentSettings(update: AiAgentSettingsUpdate) {
  const remote = webBusinessTransport();
  if (remote) return remote<AiAgentSettings>("save-settings", update);
  return invoke<AiAgentSettings>("ai_save_settings", { update });
}
