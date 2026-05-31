import { emit } from "@tauri-apps/api/event";

export const AI_TEMPLATE_ASSET_SELECTED_EVENT = "lamber-ai-template-asset-selected";
export const AI_TEMPLATE_ASSET_SELECTED_STORAGE_KEY = "lamber_ai_template_asset_selected";

export interface AiTemplateAssetSelection {
  requestId: string;
  projectId: string;
  templateId: string;
  assetId: string;
  fieldKey?: string | null;
  fileName?: string | null;
  mimeType?: string | null;
  size?: number | null;
  width?: number | null;
  height?: number | null;
  selectedAt: number;
}

function isTauriRuntime() {
  return typeof window !== "undefined" && Boolean((window as Window & { __TAURI_INTERNALS__?: unknown }).__TAURI_INTERNALS__);
}

function createRequestId() {
  return typeof crypto !== "undefined" && "randomUUID" in crypto
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

export function createTemplateAssetSelection(input: Omit<AiTemplateAssetSelection, "requestId" | "selectedAt">): AiTemplateAssetSelection {
  return {
    ...input,
    requestId: createRequestId(),
    selectedAt: Date.now(),
  };
}

export async function publishTemplateAssetSelection(selection: AiTemplateAssetSelection) {
  if (typeof window !== "undefined") {
    window.localStorage.setItem(AI_TEMPLATE_ASSET_SELECTED_STORAGE_KEY, JSON.stringify(selection));
    window.dispatchEvent(new CustomEvent(AI_TEMPLATE_ASSET_SELECTED_EVENT, { detail: selection }));
  }

  if (isTauriRuntime()) {
    await emit(AI_TEMPLATE_ASSET_SELECTED_EVENT, selection).catch((error) => {
      console.warn("Failed to emit template asset selection:", error);
    });
  }
}

export function parseTemplateAssetSelection(value: unknown): AiTemplateAssetSelection | null {
  try {
    const parsed = typeof value === "string" ? JSON.parse(value) : value;
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return null;
    const record = parsed as Record<string, unknown>;
    if (
      typeof record.requestId !== "string" ||
      typeof record.projectId !== "string" ||
      typeof record.templateId !== "string" ||
      typeof record.assetId !== "string"
    ) {
      return null;
    }
    return record as unknown as AiTemplateAssetSelection;
  } catch {
    return null;
  }
}
