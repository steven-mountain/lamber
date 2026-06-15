import { create } from "zustand";
import { projectService } from "../../utils/projectService";
import { calculateQuoteBlueprint } from "./calculations";
import { getFormulaParameterReferences, normalizeQuoteFormula } from "./formulaEngine";
import {
  canDeleteAiComputeParameterGroup,
  initializeAiComputeDiscountRate,
  migrateAiComputeDiscountRateOwnership,
  moveAiComputeParameter,
  moveAiComputeParameterByOffset,
  moveAiComputeParameterGroupByOffset,
  normalizeAiComputeParameterLayout,
  reorderAiComputeParameterGroup,
} from "./parameterLayout";
import { createH200Blueprint } from "./presets";
import {
  ictDiscountRateToAiComputePercent,
  isAiComputeProjectCycleParameter,
  isAiComputeStableIctParameter,
  normalizeAiComputeDiscountRatePercent,
  normalizeAiComputeProjectCycleValue,
} from "./fundingPlans";
import { clearAiComputeControlForMappingChange } from "./ictSync";
import type {
  AiComputeQuoteBlueprint,
  AiComputeQuoteFormula,
  AiComputeQuoteLineItem,
  AiComputeQuoteParameter,
  AiComputeQuoteParameterGroup,
  AiComputeQuotePersistedState,
  AiComputeQuoteSubjectMapping,
} from "./types";

export const AI_COMPUTE_QUOTE_SETTING_KEY = "ai_compute_quote::active";

function cloneH200Blueprint() {
  return calculateQuoteBlueprint(createH200Blueprint());
}

export function normalizeBlueprint(
  value: unknown,
  options: { fallbackDiscountRatePercent?: number } = {},
): AiComputeQuoteBlueprint {
  const fallback = cloneH200Blueprint();
  if (!value || typeof value !== "object") return fallback;
  const candidate = value as Partial<AiComputeQuoteBlueprint>;
  if (!Array.isArray(candidate.parameters) || !Array.isArray(candidate.revenueItems) || !Array.isArray(candidate.costItems)) {
    return fallback;
  }
  return calculateQuoteBlueprint(normalizeAiComputeParameterLayout({
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
  }, options));
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
  addParameterGroup: (group: AiComputeQuoteParameterGroup) => void;
  renameParameterGroup: (groupId: string, name: string, description?: string) => void;
  removeParameterGroup: (groupId: string) => boolean;
  reorderParameterGroup: (groupId: string, targetGroupId: string) => void;
  moveParameterGroupByOffset: (groupId: string, offset: -1 | 1) => void;
  moveParameter: (parameterId: string, targetGroupId: string, targetParameterId?: string | null) => void;
  moveParameterByOffset: (parameterId: string, offset: -1 | 1) => void;
  updateLineItem: (side: AiComputeQuoteLineItem["side"], itemId: string, patch: Partial<AiComputeQuoteLineItem>) => void;
  addLineItem: (item: AiComputeQuoteLineItem) => void;
  removeLineItem: (side: AiComputeQuoteLineItem["side"], itemId: string) => void;
  updateFormula: (side: AiComputeQuoteLineItem["side"], itemId: string, formula: AiComputeQuoteFormula) => void;
  updateMapping: (lineItemId: string, mapping: AiComputeQuoteSubjectMapping | null) => void;
  replaceBlueprint: (blueprint: AiComputeQuoteBlueprint, options?: { dirty?: boolean; savedAt?: string | null }) => void;
  restoreFormulaControl: (lineItemId: string) => void;
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
      const [raw, project] = await Promise.all([
        projectService.getProjectSetting(projectId, AI_COMPUTE_QUOTE_SETTING_KEY),
        projectService.getProject(projectId),
      ]);
      if (!raw) {
        const discountRatePercent = ictDiscountRateToAiComputePercent(project?.discount_rate);
        const blueprint = initializeAiComputeDiscountRate(
          cloneH200Blueprint(),
          discountRatePercent,
        );
        set({
          blueprint: recalculate(blueprint),
          projectId,
          isLoading: false,
          isDirty: false,
          lastSavedAt: null,
        });
        return;
      }
      const persisted = JSON.parse(raw) as Partial<AiComputeQuotePersistedState>;
      const discountRatePercent = ictDiscountRateToAiComputePercent(project?.discount_rate);
      const normalized = normalizeBlueprint(persisted.blueprint, {
        fallbackDiscountRatePercent: discountRatePercent,
      });
      set({
        blueprint: recalculate(migrateAiComputeDiscountRateOwnership(
          normalized,
          Number(persisted.version || 1),
          discountRatePercent,
        )),
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
        version: 4,
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
  updateParameter: (parameterId, patch) => set(state => {
    const current = state.blueprint.parameters.find(parameter => parameter.id === parameterId);
    const nextPatch = current && isAiComputeStableIctParameter(current)
      ? {
          ...patch,
          key: current.key,
          value: patch.value === undefined
            ? current.value
            : isAiComputeProjectCycleParameter(current)
              ? normalizeAiComputeProjectCycleValue(patch.value)
              : normalizeAiComputeDiscountRatePercent(patch.value),
          isKey: true,
          locked: true,
        }
      : patch;
    return {
      blueprint: recalculate({
        ...state.blueprint,
        parameters: state.blueprint.parameters.map(parameter =>
          parameter.id === parameterId ? { ...parameter, ...nextPatch } : parameter
        ),
      }),
      isDirty: true,
      error: null,
    };
  }),
  addParameter: parameter => set(state => ({
    blueprint: recalculate({ ...state.blueprint, parameters: [...state.blueprint.parameters, parameter] }),
    isDirty: true,
  })),
  removeParameter: parameterId => {
    const blueprint = get().blueprint;
    const parameter = blueprint.parameters.find(candidate => candidate.id === parameterId);
    if (parameter && isAiComputeStableIctParameter(parameter)) {
      set({ error: "项目周期和项目折现率是 ICT 同步必需参数，不能删除或修改字段 key。" });
      return false;
    }
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
    if (isAiComputeStableIctParameter(source)) {
      return {
        ...state,
        error: "项目周期和项目折现率是唯一同步参数，不能复制。",
      };
    }
    const siblings = state.blueprint.parameters.filter(parameter => parameter.groupId === source.groupId);
    const sourceSiblingIndex = siblings.findIndex(parameter => parameter.id === parameterId);
    const nextSiblingId = siblings[sourceSiblingIndex + 1]?.id;
    return {
      blueprint: recalculate({
        ...state.blueprint,
        parameters: moveAiComputeParameter([...state.blueprint.parameters, {
          ...source,
          id: newId,
          key: `${source.key}_copy`,
          name: `${source.name} 副本`,
          locked: false,
        }], newId, source.groupId || "", nextSiblingId),
      }),
      isDirty: true,
    };
  }),
  addParameterGroup: group => set(state => ({
    blueprint: {
      ...state.blueprint,
      parameterGroups: [...state.blueprint.parameterGroups, group],
    },
    isDirty: true,
    error: null,
  })),
  renameParameterGroup: (groupId, name, description) => set(state => ({
    blueprint: {
      ...state.blueprint,
      parameterGroups: state.blueprint.parameterGroups.map(group =>
        group.id === groupId
          ? { ...group, name: name.trim() || group.name, description: description?.trim() || undefined }
          : group
      ),
    },
    isDirty: true,
    error: null,
  })),
  removeParameterGroup: groupId => {
    const state = get();
    const group = state.blueprint.parameterGroups.find(candidate => candidate.id === groupId);
    if (!group || !canDeleteAiComputeParameterGroup(group, state.blueprint.parameters)) {
      set({
        error: group?.builtin
          ? "内置类别不能删除。"
          : "请先将类别中的参数移动到其他类别，再删除该类别。",
      });
      return false;
    }
    set({
      blueprint: {
        ...state.blueprint,
        parameterGroups: state.blueprint.parameterGroups.filter(candidate => candidate.id !== groupId),
      },
      isDirty: true,
      error: null,
    });
    return true;
  },
  reorderParameterGroup: (groupId, targetGroupId) => set(state => ({
    blueprint: {
      ...state.blueprint,
      parameterGroups: reorderAiComputeParameterGroup(
        state.blueprint.parameterGroups,
        groupId,
        targetGroupId,
      ),
    },
    isDirty: true,
    error: null,
  })),
  moveParameterGroupByOffset: (groupId, offset) => set(state => ({
    blueprint: {
      ...state.blueprint,
      parameterGroups: moveAiComputeParameterGroupByOffset(
        state.blueprint.parameterGroups,
        groupId,
        offset,
      ),
    },
    isDirty: true,
    error: null,
  })),
  moveParameter: (parameterId, targetGroupId, targetParameterId) => set(state => ({
    blueprint: recalculate({
      ...state.blueprint,
      parameters: moveAiComputeParameter(
        state.blueprint.parameters,
        parameterId,
        targetGroupId,
        targetParameterId,
      ),
    }),
    isDirty: true,
    error: null,
  })),
  moveParameterByOffset: (parameterId, offset) => set(state => ({
    blueprint: recalculate({
      ...state.blueprint,
      parameters: moveAiComputeParameterByOffset(state.blueprint.parameters, parameterId, offset),
    }),
    isDirty: true,
    error: null,
  })),
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
  updateMapping: (lineItemId, mapping) => set(state => {
    const released = clearAiComputeControlForMappingChange(state.blueprint, lineItemId);
    return {
      blueprint: recalculate({
        ...released,
        mappings: mapping
          ? [...released.mappings.filter(item => item.lineItemId !== lineItemId), mapping]
          : released.mappings.filter(item => item.lineItemId !== lineItemId),
      }),
      isDirty: true,
      error: null,
    };
  }),
  replaceBlueprint: (blueprint, options) => set({
    blueprint: recalculate(blueprint),
    isDirty: options?.dirty ?? false,
    lastSavedAt: options?.savedAt === undefined ? get().lastSavedAt : options.savedAt,
    error: null,
  }),
  restoreFormulaControl: lineItemId => set(state => {
    const restore = (item: AiComputeQuoteLineItem) => item.id === lineItemId
      ? {
          ...item,
          formulaControlStatus: "formula" as const,
          ictOverride: undefined,
          ictControlMessage: undefined,
        }
      : item;
    return {
      blueprint: recalculate({
        ...state.blueprint,
        revenueItems: state.blueprint.revenueItems.map(restore),
        costItems: state.blueprint.costItems.map(restore),
      }),
      isDirty: true,
      error: null,
    };
  }),
}));
