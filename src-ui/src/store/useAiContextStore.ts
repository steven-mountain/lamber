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
  replaceBusinessData: (module: string, data: any) => void;
  clearBusinessData: (module: string) => void;
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

// Simple debounce helper
function debounce<T extends (...args: any[]) => void>(func: T, wait: number): T {
  let timeout: any;
  return function(this: any, ...args: any[]) {
    clearTimeout(timeout);
    timeout = setTimeout(() => func.apply(this, args), wait);
  } as any;
}

// Simple shallow equality check for records
function isShallowEqual(a: any, b: any): boolean {
  if (a === b) return true;
  if (!isRecord(a) || !isRecord(b)) return false;
  const keysA = Object.keys(a);
  const keysB = Object.keys(b);
  if (keysA.length !== keysB.length) return false;
  for (const key of keysA) {
    if (a[key] !== b[key]) return false;
  }
  return true;
}

export const useAiContextStore = create<AiContextState>((set, get) => {
  // Debounced publisher to write to localStorage and emit Tauri events after typing stops
  const debouncedPublish = debounce(() => {
    const state = get();
    const snapshot: AiContextSnapshot = {
      activeModule: state.activeModule,
      businessData: state.businessData,
      lastUpdated: state.lastUpdated,
    };
    publishAiContextSnapshot(snapshot);
  }, 300);

  return {
    ...initialSnapshot,

    setActiveModule: (module) => {
      const nextModule = normalizeModuleKey(module);
      set((state) => {
        if (state.activeModule === nextModule) return {};

        const nextSnapshot = {
          activeModule: nextModule,
          businessData: state.businessData,
          lastUpdated: state.lastUpdated,
        };

        // Tab changes should publish immediately
        publishAiContextSnapshot(nextSnapshot);
        return { activeModule: nextModule };
      });
    },

    updateBusinessData: (module, data) => {
      const nextModule = normalizeModuleKey(module);

      set((state) => {
        const existingModuleData = state.businessData[nextModule];

        // If data is identical to existing, skip update to prevent redraw cascades
        if (existingModuleData && isShallowEqual(existingModuleData, data)) {
          return {};
        }

        const nextModuleData = isRecord(existingModuleData) && isRecord(data)
          ? { ...existingModuleData, ...data }
          : data;

        console.log('AI Store Updated (in-memory):', nextModule);

        return {
          businessData: {
            ...state.businessData,
            [nextModule]: nextModuleData,
          },
          lastUpdated: {
            ...state.lastUpdated,
            [nextModule]: Date.now(),
          },
        };
      });

      // Schedule persistence and event emission in background
      debouncedPublish();
    },

    replaceBusinessData: (module, data) => {
      const nextModule = normalizeModuleKey(module);

      set((state) => ({
        businessData: {
          ...state.businessData,
          [nextModule]: data,
        },
        lastUpdated: {
          ...state.lastUpdated,
          [nextModule]: Date.now(),
        },
      }));

      debouncedPublish();
    },

    clearBusinessData: (module) => {
      const nextModule = normalizeModuleKey(module);

      set((state) => {
        const businessData = { ...state.businessData };
        const lastUpdated = { ...state.lastUpdated };
        delete businessData[nextModule];
        delete lastUpdated[nextModule];
        return { businessData, lastUpdated };
      });

      debouncedPublish();
    },

    hydrateFromStorage: () => {
      const snapshot = readAiContextSnapshot();
      const current = get();
      if (
        snapshot.activeModule === current.activeModule &&
        JSON.stringify(snapshot.businessData) === JSON.stringify(current.businessData)
      ) {
        return;
      }

      set(snapshot);
    },
  };
});

if (typeof window !== 'undefined') {
  window.addEventListener('storage', (event) => {
    if (event.key === AI_CONTEXT_STORAGE_KEY) {
      useAiContextStore.getState().hydrateFromStorage();
    }
  });
}
