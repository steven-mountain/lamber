import { create } from 'zustand';
import { emit } from '@tauri-apps/api/event';
import { migrateLegacyAiContextKey, normalizeAiActiveModule } from '../utils/aiContextKeys';

export const AI_CONTEXT_STORAGE_KEY = 'lamber_ai_context_state';
export const AI_CONTEXT_UPDATED_EVENT = 'lamber-ai-context-updated';
export const AI_CONTEXT_REFRESH_REQUEST_EVENT = 'lamber-ai-context-refresh-request';

interface AiContextState {
  activeModule: string;
  // businessData stores the full snapshot of each module's business state
  businessData: Record<string, any>;
  lastUpdated: Record<string, number>;

  // Actions
  setActiveModule: (module: string) => void;
  updateBusinessData: (module: string, data: any) => void;
  hydrateFromStorage: () => void;
}

export type AiContextSnapshot = Pick<AiContextState, 'activeModule' | 'businessData' | 'lastUpdated'>;

const createDefaultSnapshot = (): AiContextSnapshot => ({
  activeModule: 'hub',
  businessData: {},
  lastUpdated: {},
});

function isTauriRuntime() {
  return typeof window !== 'undefined' && Boolean((window as Window & { __TAURI_INTERNALS__?: unknown }).__TAURI_INTERNALS__);
}

function isRecord(value: unknown): value is Record<string, any> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function normalizeModuleKey(module: string) {
  return migrateLegacyAiContextKey(module) ?? normalizeAiActiveModule(module);
}

export function readAiContextSnapshot(): AiContextSnapshot {
  const defaultSnapshot = createDefaultSnapshot();
  if (typeof window === 'undefined') return defaultSnapshot;

  try {
    const raw = window.localStorage.getItem(AI_CONTEXT_STORAGE_KEY);
    if (!raw) return defaultSnapshot;

    const parsed = JSON.parse(raw) as Partial<AiContextSnapshot>;
    const parsedBusinessData = isRecord(parsed.businessData)
      ? parsed.businessData
      : defaultSnapshot.businessData;
    const parsedLastUpdated = isRecord(parsed.lastUpdated)
      ? parsed.lastUpdated
      : defaultSnapshot.lastUpdated;
    const businessData: Record<string, any> = {};
    const lastUpdated: Record<string, number> = {};

    Object.entries(parsedBusinessData).forEach(([key, value]) => {
      const nextKey = migrateLegacyAiContextKey(key);
      if (!nextKey) return;
      businessData[nextKey] = value;
      if (typeof parsedLastUpdated[key] === 'number') {
        lastUpdated[nextKey] = parsedLastUpdated[key];
      }
    });

    return {
      activeModule: normalizeAiActiveModule(parsed.activeModule),
      businessData,
      lastUpdated,
    };
  } catch (error) {
    console.warn('Failed to read AI context snapshot:', error);
    return defaultSnapshot;
  }
}

export function persistAiContextSnapshot(snapshot: AiContextSnapshot) {
  if (typeof window === 'undefined') return;

  try {
    window.localStorage.setItem(AI_CONTEXT_STORAGE_KEY, JSON.stringify(snapshot));
  } catch (error) {
    console.warn('Failed to persist AI context snapshot:', error);
  }
}

export function publishAiContextSnapshot(snapshot: AiContextSnapshot) {
  persistAiContextSnapshot(snapshot);

  if (!isTauriRuntime()) return;

  emit(AI_CONTEXT_UPDATED_EVENT, snapshot).catch((error) => {
    console.warn('Failed to emit AI context update:', error);
  });
}

const initialSnapshot = readAiContextSnapshot();

export const useAiContextStore = create<AiContextState>((set, get) => ({
  ...initialSnapshot,

  setActiveModule: (module) => {
    const nextModule = normalizeModuleKey(module);
    let nextSnapshot: AiContextSnapshot | null = null;

    set((state) => {
      nextSnapshot = {
        activeModule: nextModule,
        businessData: state.businessData,
        lastUpdated: state.lastUpdated,
      };

      return { activeModule: nextModule };
    });

    if (nextSnapshot) {
      publishAiContextSnapshot(nextSnapshot);
    }
  },

  updateBusinessData: (module, data) => {
    const nextModule = normalizeModuleKey(module);
    let nextSnapshot: AiContextSnapshot | null = null;

    set((state) => {
      console.log('AI Store Updated:', nextModule, data);
      const existingModuleData = state.businessData[nextModule];
      const nextModuleData = isRecord(existingModuleData) && isRecord(data)
        ? { ...existingModuleData, ...data }
        : data;

      nextSnapshot = {
        activeModule: nextModule,
        businessData: {
          ...state.businessData,
          [nextModule]: nextModuleData,
        },
        lastUpdated: {
          ...state.lastUpdated,
          [nextModule]: Date.now(),
        },
      };

      return nextSnapshot;
    });

    if (nextSnapshot) {
      publishAiContextSnapshot(nextSnapshot);
    }
  },

  hydrateFromStorage: () => {
    const snapshot = readAiContextSnapshot();
    const current = get();
    if (
      snapshot.activeModule === current.activeModule &&
      snapshot.businessData === current.businessData &&
      snapshot.lastUpdated === current.lastUpdated
    ) {
      return;
    }

    set(snapshot);
  },
}));

if (typeof window !== 'undefined') {
  window.addEventListener('storage', (event) => {
    if (event.key === AI_CONTEXT_STORAGE_KEY) {
      useAiContextStore.getState().hydrateFromStorage();
    }
  });
}
