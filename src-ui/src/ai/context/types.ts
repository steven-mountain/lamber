import type { ContextNode } from "../types";
import type { AiProjectContextBundle } from "../../services/aiProjectContextService";
import type { DirtyScope } from "../../store/useSaveStore";

export interface AiSavedProjectContext {
  source: "workspace_sqlite";
  projectId: string;
  projectName?: string;
  resolution?: {
    reason: "active_project" | "specified_project";
    matchedName?: string;
    templateId?: string;
    templateName?: string | null;
  };
  data: AiProjectContextBundle;
}

export interface AiDraftOverlay {
  source: "unsaved_frontend_draft";
  projectId: string;
  view: string;
  dirtyScopes: DirtyScope[];
  data: unknown;
}

export interface AiComposedChatContext {
  savedProjectContext?: AiSavedProjectContext;
  savedProjectContexts?: AiSavedProjectContext[];
  draftOverlay?: AiDraftOverlay;
  contextNodes: {
    savedOfficial: ContextNode[];
    pageContext: ContextNode[];
    draftOverlay: ContextNode[];
    warnings: ContextNode[];
  };
  warnings: string[];
}
