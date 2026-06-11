import { create } from "zustand";
import { projectService } from "../../utils/projectService";
import { calculateQuoteBlueprint } from "./calculations";
import { getFormulaParameterReferences, normalizeQuoteFormula } from "./formulaEngine";
import { createH200Blueprint } from "./presets";
import type {
  AiComputeQuoteBlueprint,
  AiComputeQuoteFormula,
  AiComputeQuoteLineItem,
  AiComputeQuoteParameter,
  AiComputeQuotePersistedState,
  AiComputeQuoteSubjectMapping,
} from "./types";

export const AI_COMPUTE_QUOTE_SETTING_KEY = "ai_compute_quote::active";

function cloneH200Blueprint() {
  return calculateQuoteBlueprint(createH200Blueprint());
}

function normalizeBlueprint(value: unknown): AiComputeQuoteBlueprint {
  const fallback = cloneH200Blueprint();
  if (!value || typeof value !== "object") return fallback;
  const candidate = value as Partial<AiComputeQuoteBlueprint>;
  if (!Array.isArray(candidate.parameters) || !Array.isArray(candidate.revenueItems) || !Array.isArray(candidate.costItems)) {
    return fallback;
  }
  return calculateQuoteBlueprint({
    ...fallback,
    ...candidate,
    revenueItems: candidate.revenueItems.map(item => ({
      ...item,
      formula: normalizeQuoteFormula(item.formula, candidate.parameters),
    })),
    costItems: candidate.costItems.map(item => ({
      ...item,
      formula: normalizeQuoteFormula(item.formula, candidate.parameters),
    })),
    mappings: Array.isArray(candidate.mappings) ? candidate.mappings : [],
  });
}

interface AiComputeQuoteStore {
  blueprint: AiComputeQuoteBlueprint;
  projectId: string | null;
  isLoading: boolean;
  isSaving: boolean;
  isDirty: boolean;
  error: string | null;
  lastSavedAt: string | null;
  load: (projectId: string | null) => Promise<void>;
  save: (projectId: string | null) => Promise<boolean>;
  resetToH200: () => void;
  renameBlueprint: (name: string) => void;
  updateParameter: (parameterId: string, patch: Partial<AiComputeQuoteParameter>) => void;
  addParameter: (parameter: AiComputeQuoteParameter) => void;
  removeParameter: (parameterId: string) => boolean;
  duplicateParameter: (parameterId: string, newId: string) => void;
  updateLineItem: (side: AiComputeQuoteLineItem["side"], itemId: string, patch: Partial<AiComputeQuoteLineItem>) => void;
  addLineItem: (item: AiComputeQuoteLineItem) => void;
  removeLineItem: (side: AiComputeQuoteLineItem["side"], itemId: string) => void;
  updateFormula: (side: AiComputeQuoteLineItem["side"], itemId: string, formula: AiComputeQuoteFormula) => void;
  updateMapping: (lineItemId: string, mapping: AiComputeQuoteSubjectMapping | null) => void;
}

function recalculate(blueprint: AiComputeQuoteBlueprint) {
  return calculateQuoteBlueprint(blueprint);
}

export const useAiComputeQuoteStore = create<AiComputeQuoteStore>((set, get) => ({
  blueprint: cloneH200Blueprint(),
  projectId: null,
  isLoading: false,
  isSaving: false,
  isDirty: false,
  error: null,
  lastSavedAt: null,

  load: async projectId => {
    if (!projectId) {
      set({ blueprint: cloneH200Blueprint(), projectId: null, isDirty: false, error: null, lastSavedAt: null });
      return;
    }
    set({ isLoading: true, error: null });
    try {
      const raw = await projectService.getProjectSetting(projectId, AI_COMPUTE_QUOTE_SETTING_KEY);
      if (!raw) {
        set({ blueprint: cloneH200Blueprint(), projectId, isLoading: false, isDirty: false, lastSavedAt: null });
        return;
      }
      const persisted = JSON.parse(raw) as Partial<AiComputeQuotePersistedState>;
      set({
        blueprint: normalizeBlueprint(persisted.blueprint),
        projectId,
        isLoading: false,
        isDirty: false,
        lastSavedAt: typeof persisted.savedAt === "string" ? persisted.savedAt : null,
      });
    } catch (error) {
      set({ isLoading: false, error: `智算报价加载失败：${String(error)}` });
    }
  },

  save: async projectId => {
    if (!projectId) {
      set({ error: "请先选择项目后再保存智算蓝图。" });
      return false;
    }
    set({ isSaving: true, error: null });
    try {
      const savedAt = new Date().toISOString();
      const payload: AiComputeQuotePersistedState = {
        version: 1,
        blueprint: recalculate(get().blueprint),
        savedAt,
      };
      await projectService.saveProjectSetting(projectId, AI_COMPUTE_QUOTE_SETTING_KEY, JSON.stringify(payload));
      set({ projectId, isSaving: false, isDirty: false, lastSavedAt: savedAt });
      return true;
    } catch (error) {
      set({ isSaving: false, error: `智算报价保存失败：${String(error)}` });
      return false;
    }
  },

  resetToH200: () => set({ blueprint: cloneH200Blueprint(), isDirty: true, error: null }),
  renameBlueprint: name => set(state => ({
    blueprint: { ...state.blueprint, name: name.trim() || state.blueprint.name },
    isDirty: true,
  })),
  updateParameter: (parameterId, patch) => set(state => ({
    blueprint: recalculate({
      ...state.blueprint,
      parameters: state.blueprint.parameters.map(parameter =>
        parameter.id === parameterId ? { ...parameter, ...patch } : parameter
      ),
    }),
    isDirty: true,
    error: null,
  })),
  addParameter: parameter => set(state => ({
    blueprint: recalculate({ ...state.blueprint, parameters: [...state.blueprint.parameters, parameter] }),
    isDirty: true,
  })),
  removeParameter: parameterId => {
    const blueprint = get().blueprint;
    const isReferenced = [...blueprint.revenueItems, ...blueprint.costItems]
      .some(item => getFormulaParameterReferences(item.formula).includes(parameterId));
    if (isReferenced) {
      set({ error: "该参数仍被收入或成本公式引用，请先调整公式。" });
      return false;
    }
    set({
      blueprint: recalculate({
        ...blueprint,
        parameters: blueprint.parameters.filter(parameter => parameter.id !== parameterId),
      }),
      isDirty: true,
      error: null,
    });
    return true;
  },
  duplicateParameter: (parameterId, newId) => set(state => {
    const source = state.blueprint.parameters.find(parameter => parameter.id === parameterId);
    if (!source) return state;
    return {
      blueprint: recalculate({
        ...state.blueprint,
        parameters: [...state.blueprint.parameters, {
          ...source,
          id: newId,
          key: `${source.key}_copy`,
          name: `${source.name} 副本`,
          locked: false,
        }],
      }),
      isDirty: true,
    };
  }),
  updateLineItem: (side, itemId, patch) => set(state => {
    const key = side === "revenue" ? "revenueItems" : "costItems";
    return {
      blueprint: recalculate({
        ...state.blueprint,
        [key]: state.blueprint[key].map(item => item.id === itemId ? { ...item, ...patch } : item),
      }),
      isDirty: true,
      error: null,
    };
  }),
  addLineItem: item => set(state => {
    const key = item.side === "revenue" ? "revenueItems" : "costItems";
    return {
      blueprint: recalculate({ ...state.blueprint, [key]: [...state.blueprint[key], item] }),
      isDirty: true,
    };
  }),
  removeLineItem: (side, itemId) => set(state => {
    const key = side === "revenue" ? "revenueItems" : "costItems";
    return {
      blueprint: recalculate({
        ...state.blueprint,
        [key]: state.blueprint[key].filter(item => item.id !== itemId),
        mappings: state.blueprint.mappings.filter(mapping => mapping.lineItemId !== itemId),
      }),
      isDirty: true,
    };
  }),
  updateFormula: (side, itemId, formula) => get().updateLineItem(side, itemId, { formula }),
  updateMapping: (lineItemId, mapping) => set(state => ({
    blueprint: {
      ...state.blueprint,
      mappings: mapping
        ? [...state.blueprint.mappings.filter(item => item.lineItemId !== lineItemId), mapping]
        : state.blueprint.mappings.filter(item => item.lineItemId !== lineItemId),
    },
    isDirty: true,
  })),
}));
