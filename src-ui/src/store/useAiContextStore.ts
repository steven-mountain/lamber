import { create } from 'zustand';

interface AiContextState {
  activeModule: string;
  // businessData stores the full snapshot of each module's business state
  businessData: Record<string, any>;
  lastUpdated: Record<string, number>;
  
  // Actions
  setActiveModule: (module: string) => void;
  updateBusinessData: (module: string, data: any) => void;
}

export const useAiContextStore = create<AiContextState>((set) => ({
  activeModule: 'hub',
  businessData: {},
  lastUpdated: {},
  
  setActiveModule: (module) => set({ activeModule: module }),
  
  updateBusinessData: (module, data) => set((state) => {
    console.log('AI Store Updated:', module, data);
    return {
      businessData: {
        ...state.businessData,
        [module]: data
      },
      lastUpdated: {
        ...state.lastUpdated,
        [module]: Date.now()
      }
    };
  }),
}));
