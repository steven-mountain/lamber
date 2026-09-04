import { create } from 'zustand';
import {
  AI_SESSION_STORAGE_KEY,
  AI_SESSION_STORAGE_VERSION,
  DEFAULT_AI_SESSION_TITLE,
  type AiSession,
  type AiSessionSnapshot,
  type AiSessionTitleSource,
} from '../ai/sessionTypes';
import type { AiChatMessage, AiImageAttachment } from '../ai/types';

const DEFAULT_WELCOME_MESSAGE: AiChatMessage = {
  role: 'assistant',
  content: '您好！我是 Lamber 智能售前顾问。我可以帮您分析当前页面的项目效益、推荐内置产品。请问有什么可以帮您？',
};

const PERSIST_DEBOUNCE_MS = 160;

interface AiSessionState {
  sessions: AiSession[];
  currentSessionId: string | null;
  createSession: (projectId?: string) => string;
  ensureActiveSession: (projectId?: string) => string;
  selectSession: (sessionId: string) => void;
  deleteSession: (sessionId: string) => void;
  appendMessages: (sessionId: string, messages: AiChatMessage[]) => void;
  updateLastAssistantMessage: (
    sessionId: string,
    patch: Pick<AiChatMessage, 'content' | 'think'>,
  ) => void;
  resetSessionMessages: (sessionId: string, message?: AiChatMessage) => void;
  setSessionTitle: (
    sessionId: string,
    title: string,
    source?: Exclude<AiSessionTitleSource, 'default'>,
  ) => void;
  hydrateFromStorage: () => void;
  flushPersistence: () => void;
}

let pendingSnapshot: AiSessionSnapshot | null = null;
let persistTimer: ReturnType<typeof setTimeout> | null = null;

function createSessionId() {
  return typeof crypto !== 'undefined' && 'randomUUID' in crypto
    ? crypto.randomUUID()
    : `ai-session-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function createSessionRecord(projectId?: string): AiSession {
  const now = Date.now();
  return {
    id: createSessionId(),
    title: DEFAULT_AI_SESSION_TITLE,
    projectId: projectId || undefined,
    harnessSessionId: undefined,
    createdAt: now,
    updatedAt: now,
    messages: [{ ...DEFAULT_WELCOME_MESSAGE }],
    titleSource: 'default',
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function normalizeAttachment(value: unknown): AiImageAttachment | null {
  if (!isRecord(value) || typeof value.id !== 'string' || typeof value.name !== 'string') {
    return null;
  }

  return {
    id: value.id,
    name: value.name,
    mimeType: typeof value.mimeType === 'string' ? value.mimeType : 'image/png',
    size: typeof value.size === 'number' ? value.size : 0,
    dataUrl: typeof value.dataUrl === 'string' ? value.dataUrl : undefined,
    source: value.source === 'template_asset' ? 'template_asset' : 'user_upload',
    projectId: typeof value.projectId === 'string' ? value.projectId : undefined,
    templateId: typeof value.templateId === 'string' ? value.templateId : undefined,
    assetId: typeof value.assetId === 'string' ? value.assetId : undefined,
    fieldKey: typeof value.fieldKey === 'string' ? value.fieldKey : undefined,
  };
}

function normalizeMessage(value: unknown): AiChatMessage | null {
  if (!isRecord(value) || (value.role !== 'user' && value.role !== 'assistant')) {
    return null;
  }

  return {
    role: value.role,
    content: typeof value.content === 'string' ? value.content : '',
    think: typeof value.think === 'string' ? value.think : undefined,
    images: Array.isArray(value.images)
      ? value.images.map(normalizeAttachment).filter((image): image is AiImageAttachment => Boolean(image))
      : undefined,
  };
}

function normalizeSession(value: unknown): AiSession | null {
  if (!isRecord(value) || typeof value.id !== 'string') return null;

  const createdAt = typeof value.createdAt === 'number' ? value.createdAt : Date.now();
  const updatedAt = typeof value.updatedAt === 'number' ? value.updatedAt : createdAt;
  const messages = Array.isArray(value.messages)
    ? value.messages.map(normalizeMessage).filter((message): message is AiChatMessage => Boolean(message))
    : [];
  const titleSource = value.titleSource === 'manual' || value.titleSource === 'generated'
    ? value.titleSource
    : 'default';

  return {
    id: value.id,
    title: typeof value.title === 'string' && value.title.trim()
      ? value.title.trim()
      : DEFAULT_AI_SESSION_TITLE,
    projectId: typeof value.projectId === 'string' && value.projectId.trim()
      ? value.projectId
      : undefined,
    harnessSessionId: typeof value.harnessSessionId === 'string' && value.harnessSessionId.trim()
      ? value.harnessSessionId
      : undefined,
    createdAt,
    updatedAt,
    messages: messages.length > 0 ? messages : [{ ...DEFAULT_WELCOME_MESSAGE }],
    titleSource,
  };
}

export function readAiSessionSnapshot(): AiSessionSnapshot {
  const emptySnapshot: AiSessionSnapshot = {
    version: AI_SESSION_STORAGE_VERSION,
    sessions: [],
    currentSessionId: null,
  };
  if (typeof window === 'undefined') return emptySnapshot;

  try {
    const raw = window.localStorage.getItem(AI_SESSION_STORAGE_KEY);
    if (!raw) return emptySnapshot;

    const parsed = JSON.parse(raw) as unknown;
    if (!isRecord(parsed) || !Array.isArray(parsed.sessions)) return emptySnapshot;

    const sessions = parsed.sessions
      .map(normalizeSession)
      .filter((session): session is AiSession => Boolean(session));
    const requestedCurrentId = typeof parsed.currentSessionId === 'string'
      ? parsed.currentSessionId
      : null;
    const currentSessionId = sessions.some(session => session.id === requestedCurrentId)
      ? requestedCurrentId
      : [...sessions].sort((a, b) => b.updatedAt - a.updatedAt)[0]?.id ?? null;

    return {
      version: AI_SESSION_STORAGE_VERSION,
      sessions,
      currentSessionId,
    };
  } catch (error) {
    console.warn('Failed to read AI session workspace:', error);
    return emptySnapshot;
  }
}

function stripVolatileAttachmentData(session: AiSession): AiSession {
  return {
    ...session,
    messages: session.messages.map(message => ({
      ...message,
      images: message.images?.map(image => ({
        ...image,
        // User-uploaded data URLs can be several MB each. Keep the attachment
        // record without risking failure of the complete localStorage snapshot.
        dataUrl: undefined,
      })),
    })),
  };
}

function writeSnapshot(snapshot: AiSessionSnapshot) {
  if (typeof window === 'undefined') return;

  try {
    const persistable: AiSessionSnapshot = {
      ...snapshot,
      sessions: snapshot.sessions.map(stripVolatileAttachmentData),
    };
    window.localStorage.setItem(AI_SESSION_STORAGE_KEY, JSON.stringify(persistable));
  } catch (error) {
    console.warn('Failed to persist AI session workspace:', error);
  }
}

function persistImmediately(snapshot: AiSessionSnapshot) {
  if (persistTimer) clearTimeout(persistTimer);
  persistTimer = null;
  pendingSnapshot = null;
  writeSnapshot(snapshot);
}

function schedulePersistence(snapshot: AiSessionSnapshot) {
  pendingSnapshot = snapshot;
  if (persistTimer) return;
  persistTimer = setTimeout(() => {
    persistTimer = null;
    if (!pendingSnapshot) return;
    const snapshotToWrite = pendingSnapshot;
    pendingSnapshot = null;
    writeSnapshot(snapshotToWrite);
  }, PERSIST_DEBOUNCE_MS);
}

function toSnapshot(state: Pick<AiSessionState, 'sessions' | 'currentSessionId'>): AiSessionSnapshot {
  return {
    version: AI_SESSION_STORAGE_VERSION,
    sessions: state.sessions,
    currentSessionId: state.currentSessionId,
  };
}

const initialSnapshot = readAiSessionSnapshot();

export const useAiSessionStore = create<AiSessionState>((set, get) => ({
  sessions: initialSnapshot.sessions,
  currentSessionId: initialSnapshot.currentSessionId,

  createSession: (projectId) => {
    const session = createSessionRecord(projectId);
    set(state => ({
      sessions: [session, ...state.sessions],
      currentSessionId: session.id,
    }));
    persistImmediately(toSnapshot(get()));
    return session.id;
  },

  ensureActiveSession: (projectId) => {
    const state = get();
    if (state.currentSessionId && state.sessions.some(session => session.id === state.currentSessionId)) {
      return state.currentSessionId;
    }
    if (state.sessions.length > 0) {
      const latest = [...state.sessions].sort((a, b) => b.updatedAt - a.updatedAt)[0];
      set({ currentSessionId: latest.id });
      persistImmediately(toSnapshot(get()));
      return latest.id;
    }
    return get().createSession(projectId);
  },

  selectSession: (sessionId) => {
    const state = get();
    if (state.currentSessionId === sessionId || !state.sessions.some(session => session.id === sessionId)) {
      return;
    }
    set({ currentSessionId: sessionId });
    persistImmediately(toSnapshot(get()));
  },

  deleteSession: (sessionId) => {
    const state = get();
    if (!state.sessions.some(session => session.id === sessionId)) return;

    const sessions = state.sessions.filter(session => session.id !== sessionId);
    const currentSessionId = state.currentSessionId === sessionId
      ? [...sessions].sort((a, b) => b.updatedAt - a.updatedAt)[0]?.id ?? null
      : state.currentSessionId;
    set({ sessions, currentSessionId });
    persistImmediately(toSnapshot(get()));
  },

  appendMessages: (sessionId, messages) => {
    const now = Date.now();
    set(state => ({
      sessions: state.sessions.map(session => session.id === sessionId
        ? { ...session, messages: [...session.messages, ...messages], updatedAt: now }
        : session),
    }));
    schedulePersistence(toSnapshot(get()));
  },

  updateLastAssistantMessage: (sessionId, patch) => {
    const now = Date.now();
    set(state => ({
      sessions: state.sessions.map(session => {
        if (session.id !== sessionId || session.messages.length === 0) return session;
        const lastIndex = session.messages.length - 1;
        const lastMessage = session.messages[lastIndex];
        if (lastMessage.role !== 'assistant') return session;
        if (lastMessage.content === patch.content && lastMessage.think === patch.think) return session;

        const messages = [...session.messages];
        messages[lastIndex] = { ...lastMessage, ...patch };
        return { ...session, messages, updatedAt: now };
      }),
    }));
    schedulePersistence(toSnapshot(get()));
  },

  resetSessionMessages: (sessionId, message = DEFAULT_WELCOME_MESSAGE) => {
    const now = Date.now();
    set(state => ({
      sessions: state.sessions.map(session => session.id === sessionId
        ? { ...session, messages: [{ ...message }], updatedAt: now }
        : session),
    }));
    persistImmediately(toSnapshot(get()));
  },

  setSessionTitle: (sessionId, title, source = 'manual') => {
    const normalizedTitle = title.trim();
    if (!normalizedTitle) return;
    const now = Date.now();
    set(state => ({
      sessions: state.sessions.map(session => session.id === sessionId
        ? { ...session, title: normalizedTitle, titleSource: source, updatedAt: now }
        : session),
    }));
    persistImmediately(toSnapshot(get()));
  },

  hydrateFromStorage: () => {
    const snapshot = readAiSessionSnapshot();
    set({
      sessions: snapshot.sessions,
      currentSessionId: snapshot.currentSessionId,
    });
  },

  flushPersistence: () => {
    persistImmediately(toSnapshot(get()));
  },
}));

if (typeof window !== 'undefined') {
  window.addEventListener('storage', event => {
    if (event.key === AI_SESSION_STORAGE_KEY) {
      useAiSessionStore.getState().hydrateFromStorage();
    }
  });
  window.addEventListener('pagehide', () => {
    useAiSessionStore.getState().flushPersistence();
  });
}
