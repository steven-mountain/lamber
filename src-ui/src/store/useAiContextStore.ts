import { create } from 'zustand';

const AI_CONTEXT_STORAGE_KEY = 'lamber_ai_context_state';

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

type AiContextSnapshot = Pick<AiContextState, 'activeModule' | 'businessData' | 'lastUpdated'>;

const defaultSnapshot: AiContextSnapshot = {
  activeModule: 'hub',
  businessData: {},
  lastUpdated: {},
};

function readPersistedSnapshot(): AiContextSnapshot {
  if (typeof window === 'undefined') return defaultSnapshot;

  try {
    const raw = window.localStorage.getItem(AI_CONTEXT_STORAGE_KEY);
    if (!raw) return defaultSnapshot;

    const parsed = JSON.parse(raw) as Partial<AiContextSnapshot>;
    return {
      activeModule: typeof parsed.activeModule === 'string' ? parsed.activeModule : defaultSnapshot.activeModule,
      businessData: parsed.businessData && typeof parsed.businessData === 'object' ? parsed.businessData : defaultSnapshot.businessData,
      lastUpdated: parsed.lastUpdated && typeof parsed.lastUpdated === 'object' ? parsed.lastUpdated : defaultSnapshot.lastUpdated,
    };
  } catch (error) {
    console.warn('Failed to read AI context snapshot:', error);
    return defaultSnapshot;
  }
}

function persistSnapshot(snapshot: AiContextSnapshot) {
  if (typeof window === 'undefined') return;

  try {
    window.localStorage.setItem(AI_CONTEXT_STORAGE_KEY, JSON.stringify(snapshot));
  } catch (error) {
    console.warn('Failed to persist AI context snapshot:', error);
  }
}

const initialSnapshot = readPersistedSnapshot();

export const useAiContextStore = create<AiContextState>((set, get) => ({
  ...initialSnapshot,

  setActiveModule: (module) => set((state) => {
    persistSnapshot({
      activeModule: module,
      businessData: state.businessData,
      lastUpdated: state.lastUpdated,
    });
    return { activeModule: module };
  }),

  updateBusinessData: (module, data) => set((state) => {
    console.log('AI Store Updated:', module, data);
    const nextSnapshot = {
      activeModule: state.activeModule,
      businessData: {
        ...state.businessData,
        [module]: data,
      },
      lastUpdated: {
        ...state.lastUpdated,
        [module]: Date.now(),
      },
    };

    persistSnapshot(nextSnapshot);
    return {
      businessData: nextSnapshot.businessData,
      lastUpdated: nextSnapshot.lastUpdated,
    };
  }),

  hydrateFromStorage: () => {
    const snapshot = readPersistedSnapshot();
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
