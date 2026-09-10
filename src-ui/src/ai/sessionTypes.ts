import type { AiChatMessage } from './types';

export const AI_SESSION_STORAGE_KEY = 'lamber_ai_session_workspace';
export const AI_SESSION_STORAGE_VERSION = 1;
export const DEFAULT_AI_SESSION_TITLE = '新会话';

export type AiSessionTitleSource = 'default' | 'manual' | 'generated';

/** Retired frontend history schema, retained only for lossless read-only migration.
 * Frontend project/harness mirrors never authorize a restored WebUI session.
 */
export interface AiSession {
  id: string;
  title: string;
  projectId?: string;
  harnessSessionId?: string;
  createdAt: number;
  updatedAt: number;
  messages: AiChatMessage[];
  titleSource?: AiSessionTitleSource;
}

export interface AiSessionSnapshot {
  version: typeof AI_SESSION_STORAGE_VERSION;
  sessions: AiSession[];
  currentSessionId: string | null;
}
