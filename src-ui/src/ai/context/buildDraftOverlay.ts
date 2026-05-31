import { AI_CONTEXT_KEY, getAiContextScope, isAiContextKeyForView } from "../../utils/aiContextKeys";
import type { AiContextSnapshot } from "../../store/useAiContextStore";
import type { DirtyScope } from "../../store/useSaveStore";
import type { AiDraftOverlay } from "./types";

const DRAFT_SCOPE_BY_VIEW: Record<string, DirtyScope[]> = {
  project_board: ["project-detail"],
  ict_lifecycle: ["lifecycle", "cashflow", "benefit-analysis", "template-forms"],
  ict: ["lifecycle", "cashflow", "benefit-analysis", "template-forms"],
};

const MAX_ARRAY_ITEMS = 20;
const MAX_OBJECT_KEYS = 80;
const MAX_STRING_LENGTH = 1200;
const MAX_DEPTH = 5;

interface DraftOverlayInput {
  projectId?: string | null;
  currentView: string;
  dirtyScopes: DirtyScope[];
  aiSnapshot: AiContextSnapshot;
}

function normalizeContextView(view: string) {
  if (view === "ict_lifecycle") return "ict";
  return view;
}

function relevantDirtyScopes(view: string, dirtyScopes: DirtyScope[]) {
  const allowed = DRAFT_SCOPE_BY_VIEW[view] || [];
  return dirtyScopes.filter(scope => allowed.includes(scope));
}

function extractProjectId(value: unknown): string | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const candidates = [
    record.projectId,
    record.project_id,
    record.activeProjectId,
    record.active_project_id,
  ];
  for (const candidate of candidates) {
    if (typeof candidate === "string" && candidate.trim()) {
      return candidate.trim();
    }
  }
  return null;
}

function isAbsolutePath(value: string) {
  return /^[A-Za-z]:[\\/]/.test(value) || value.startsWith("\\\\") || value.startsWith("/");
}

function isSensitiveKey(key: string) {
  const normalized = key.toLowerCase();
  return (
    normalized.includes("base64") ||
    normalized.includes("dataurl") ||
    normalized.includes("preview") ||
    normalized === "src" ||
    normalized === "url"
  );
}

function sanitizeDraftValue(value: unknown, key = "", depth = 0): unknown {
  if (value === null || value === undefined) return value;
  if (typeof value === "string") {
    if (value.startsWith("data:") || isSensitiveKey(key)) return "[omitted large or binary preview]";
    if (key.toLowerCase().includes("path") && isAbsolutePath(value)) return "[omitted absolute path]";
    return value.length > MAX_STRING_LENGTH ? `${value.slice(0, MAX_STRING_LENGTH)}... [truncated]` : value;
  }
  if (typeof value !== "object") return value;
  if (depth >= MAX_DEPTH) return "[truncated nested object]";

  if (Array.isArray(value)) {
    const items = value.slice(0, MAX_ARRAY_ITEMS).map(item => sanitizeDraftValue(item, key, depth + 1));
    if (value.length > MAX_ARRAY_ITEMS) {
      items.push(`[truncated ${value.length - MAX_ARRAY_ITEMS} items]`);
    }
    return items;
  }

  const entries = Object.entries(value as Record<string, unknown>).slice(0, MAX_OBJECT_KEYS);
  const sanitized: Record<string, unknown> = {};
  for (const [childKey, childValue] of entries) {
    sanitized[childKey] = sanitizeDraftValue(childValue, childKey, depth + 1);
  }
  const totalKeys = Object.keys(value as Record<string, unknown>).length;
  if (totalKeys > MAX_OBJECT_KEYS) {
    sanitized.__truncatedKeys = totalKeys - MAX_OBJECT_KEYS;
  }
  return sanitized;
}

function selectDraftData(input: DraftOverlayInput) {
  const contextView = normalizeContextView(input.currentView);
  const activeModule = input.aiSnapshot.activeModule;
  const activeModuleData = activeModule && activeModule !== AI_CONTEXT_KEY.HUB
    ? input.aiSnapshot.businessData[activeModule]
    : undefined;

  const shouldUseActiveModule =
    activeModule &&
    activeModule !== AI_CONTEXT_KEY.HUB &&
    isAiContextKeyForView(activeModule, contextView) &&
    activeModuleData;

  if (shouldUseActiveModule) {
    return {
      module: activeModule,
      data: activeModuleData,
    };
  }

  const coreKey = contextView === "ict"
    ? AI_CONTEXT_KEY.ICT_CORE
    : contextView === "project_board"
      ? AI_CONTEXT_KEY.PROJECT_BOARD_CORE
      : `${contextView}.core`;
  return {
    module: coreKey,
    data: input.aiSnapshot.businessData[coreKey],
  };
}

export function buildDraftOverlay(input: DraftOverlayInput): AiDraftOverlay | undefined {
  const projectId = input.projectId?.trim();
  if (!projectId) return undefined;

  const dirtyScopes = relevantDirtyScopes(input.currentView, input.dirtyScopes);
  if (dirtyScopes.length === 0) return undefined;

  const selected = selectDraftData(input);
  if (!selected.data) return undefined;

  const draftProjectId = extractProjectId(selected.data);
  if (draftProjectId && draftProjectId !== projectId) return undefined;

  const scope = getAiContextScope(selected.module);
  if ((input.currentView === "ict_lifecycle" || input.currentView === "ict") && scope !== "ict") return undefined;

  return {
    source: "unsaved_frontend_draft",
    projectId,
    view: input.currentView,
    dirtyScopes,
    data: {
      module: selected.module,
      state: sanitizeDraftValue(selected.data),
    },
  };
}
