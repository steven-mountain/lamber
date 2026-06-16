import { create } from "zustand";
import { intelligentComputeService } from "../../services/intelligentComputeService";
import { calculateQuoteBlueprint } from "./calculations";
import { getFormulaParameterReferences, normalizeQuoteFormula } from "./formulaEngine";
import {
  canDeleteAiComputeParameterGroup,
  moveAiComputeParameter,
  moveAiComputeParameterByOffset,
  moveAiComputeParameterGroupByOffset,
  normalizeAiComputeParameterLayout,
  reorderAiComputeParameterGroup,
} from "./parameterLayout";
import { createH200Blueprint } from "./presets";
import {
  getAiComputeDiscountRatePercent,
  getAiComputeProjectCycleYears,
  isAiComputeProjectCycleParameter,
  isAiComputeStableIctParameter,
  normalizeAiComputeDiscountRatePercent,
  normalizeAiComputeProjectCycleValue,
} from "./fundingPlans";
import type {
  AiComputeQuoteBlueprint,
  AiComputeQuoteFormula,
  AiComputeQuoteLineItem,
  AiComputeQuoteParameter,
  AiComputeQuoteParameterGroup,
  AiComputeQuoteSubjectMapping,
  IntelligentAmountSource,
  IntelligentComputeProjectState,
} from "./types";

export const AI_COMPUTE_QUOTE_SETTING_KEY = "ai_compute_quote::active";

export type CreateAmountSourceBaseMode = "blank" | "h200" | "current" | "source";

export interface CreateAmountSourceRequest {
  name: string;
  baseMode: CreateAmountSourceBaseMode;
  baseSourceId?: string | null;
  enabled?: boolean;
}

function cloneH200Blueprint() {
  return calculateQuoteBlueprint(createH200Blueprint());
}

let loadSequence = 0;

function injectProjectParameters(
  blueprint: AiComputeQuoteBlueprint,
  projectState: IntelligentComputeProjectState,
) {
  return {
    ...blueprint,
    parameters: blueprint.parameters.map(parameter =>
      isAiComputeProjectCycleParameter(parameter)
        ? { ...parameter, value: projectState.projectYears, locked: true, isKey: true }
        : parameter.id === "discount-rate" || parameter.key === "discount_rate"
          ? { ...parameter, value: projectState.discountRate * 100, locked: true, isKey: true }
          : parameter
    ),
  };
}

export function buildBlueprintForAmountSource(
  source: IntelligentAmountSource,
  projectState: IntelligentComputeProjectState,
) {
  const hasBusinessData = source.parameters.length > 0
    || source.revenueItems.length > 0
    || source.costItems.length > 0;
  const raw: AiComputeQuoteBlueprint = hasBusinessData
    ? {
        id: source.id,
        scenarioId: String(source.metadata.scenarioId || source.id),
        name: source.name,
        description: source.description || undefined,
        parameterGroups: source.parameterGroups,
        parameters: source.parameters,
        revenueItems: source.revenueItems,
        costItems: source.costItems,
        mappings: source.mappings,
        syncState: source.calculationSnapshot.syncState,
      }
    : {
        ...cloneH200Blueprint(),
        id: source.id,
        scenarioId: source.id,
        name: source.name,
        description: source.description || "智算项目金额来源",
      };
  return calculateQuoteBlueprint(injectProjectParameters(normalizeBlueprint(raw), projectState));
}

function blueprintToSource(
  source: IntelligentAmountSource,
  blueprint: AiComputeQuoteBlueprint,
): IntelligentAmountSource {
  return {
    ...source,
    name: blueprint.name,
    description: blueprint.description || null,
    enabled: true,
    metadata: {
      ...source.metadata,
      scenarioId: blueprint.scenarioId || blueprint.id,
    },
    parameterGroups: blueprint.parameterGroups,
    parameters: blueprint.parameters,
    revenueItems: blueprint.revenueItems,
    costItems: blueprint.costItems,
    mappings: blueprint.mappings,
    calculationSnapshot: {
      ...source.calculationSnapshot,
      syncState: blueprint.syncState,
      calculatedAt: new Date().toISOString(),
    },
  };
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
  amountSources: IntelligentAmountSource[];
  projectState: IntelligentComputeProjectState | null;
  activeAmountSourceId: string | null;
  projectId: string | null;
  isLoading: boolean;
  isSaving: boolean;
  isDirty: boolean;
  error: string | null;
  lastSavedAt: string | null;
  load: (projectId: string | null) => Promise<void>;
  save: (projectId: string | null) => Promise<boolean>;
  setActiveAmountSource: (sourceId: string) => Promise<boolean>;
  createAmountSource: (request: CreateAmountSourceRequest) => Promise<boolean>;
  deleteAmountSource: (sourceId: string) => Promise<boolean>;
  updateAmountSourceMeta: (sourceId: string, patch: Partial<Pick<IntelligentAmountSource, "name" | "description" | "enabled">>) => void;
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
  amountSources: [],
  projectState: null,
  activeAmountSourceId: null,
  projectId: null,
  isLoading: false,
  isSaving: false,
  isDirty: false,
  error: null,
  lastSavedAt: null,

  load: async projectId => {
    const requestId = ++loadSequence;
    if (!projectId) {
      set({
        blueprint: cloneH200Blueprint(),
        amountSources: [],
        projectState: null,
        activeAmountSourceId: null,
        projectId: null,
        isLoading: false,
        isSaving: false,
        isDirty: false,
        error: null,
        lastSavedAt: null,
      });
      return;
    }
    set({
      blueprint: cloneH200Blueprint(),
      amountSources: [],
      projectState: null,
      activeAmountSourceId: null,
      projectId,
      isLoading: true,
      isDirty: false,
      error: null,
      lastSavedAt: null,
    });
    try {
      const data = await intelligentComputeService.loadProject(projectId);
      if (requestId !== loadSequence || get().projectId !== projectId) return;
      const activeId = data.state.activeAmountSourceId || data.amountSources[0]?.id || null;
      const activeSource = data.amountSources.find(source => source.id === activeId)
        || data.amountSources[0];
      if (!activeSource) throw new Error("智算项目缺少金额来源");
      set({
        blueprint: buildBlueprintForAmountSource(activeSource, data.state),
        amountSources: data.amountSources,
        projectState: data.state,
        activeAmountSourceId: activeSource.id,
        projectId,
        isLoading: false,
        isDirty: false,
        lastSavedAt: activeSource.updatedAt || null,
      });
    } catch (error) {
      if (requestId === loadSequence) {
        set({ isLoading: false, error: `智算金额来源加载失败：${String(error)}` });
      }
    }
  },

  save: async projectId => {
    if (!projectId) {
      set({ error: "请先打开智算项目后再保存金额来源。" });
      return false;
    }
    const current = get();
    const source = current.amountSources.find(item => item.id === current.activeAmountSourceId);
    if (!source || !current.projectState) {
      set({ error: "当前金额来源不存在。" });
      return false;
    }
    set({ isSaving: true, error: null });
    try {
      const blueprint = recalculate(get().blueprint);
      const savedSource = await intelligentComputeService.saveAmountSource(
        projectId,
        blueprintToSource(source, blueprint),
        source.sourceVersion,
      );
      const latestState = get().projectState;
      if (!latestState || get().projectId !== projectId) return false;
      const projectYears = getAiComputeProjectCycleYears(blueprint.parameters);
      const discountRate = getAiComputeDiscountRatePercent(blueprint.parameters) / 100;
      const savedState = await intelligentComputeService.saveProjectState(projectId, {
        expectedVersion: latestState.stateVersion,
        activeAmountSourceId: savedSource.id,
        projectYears,
        discountRate,
      });
      if (get().projectId !== projectId) return false;
      set(state => ({
        amountSources: state.amountSources.map(item => item.id === savedSource.id ? savedSource : item),
        projectState: savedState,
        projectId,
        isSaving: false,
        isDirty: false,
        lastSavedAt: savedSource.updatedAt,
      }));
      return true;
    } catch (error) {
      set({ isSaving: false, error: `智算金额来源保存失败：${String(error)}` });
      return false;
    }
  },

  setActiveAmountSource: async sourceId => {
    const state = get();
    if (!state.projectId || !state.projectState || sourceId === state.activeAmountSourceId) return true;
    if (state.isDirty && !(await state.save(state.projectId))) return false;
    const latest = get();
    const source = latest.amountSources.find(item => item.id === sourceId);
    if (!source || !latest.projectState) return false;
    try {
      const savedState = await intelligentComputeService.saveProjectState(latest.projectId!, {
        expectedVersion: latest.projectState.stateVersion,
        activeAmountSourceId: sourceId,
        projectYears: latest.projectState.projectYears,
        discountRate: latest.projectState.discountRate,
      });
      set({
        activeAmountSourceId: sourceId,
        blueprint: buildBlueprintForAmountSource(source, savedState),
        amountSources: latest.amountSources.map(item => ({
          ...item,
          enabled: item.id === sourceId,
        })),
        projectState: savedState,
        isDirty: false,
        lastSavedAt: source.updatedAt,
        error: null,
      });
      return true;
    } catch (error) {
      set({ error: `切换金额来源失败：${String(error)}` });
      return false;
    }
	  },

	  createAmountSource: async request => {
	    const state = get();
	    if (!state.projectId || !state.projectState) return false;
	    if (state.isDirty && !(await state.save(state.projectId))) return false;
	    const latest = get();
	    const name = request.name.trim();
	    if (!name) {
	      set({ error: "请先填写金额来源名称。" });
	      return false;
	    }
	    const sourceId = `amount_source_${Date.now()}_${Math.random().toString(16).slice(2)}`;
	    const h200 = cloneH200Blueprint();
	    const selectedBaseSource = latest.amountSources.find(source => source.id === request.baseSourceId);
	    const base = request.baseMode === "current"
	      ? latest.blueprint
	      : request.baseMode === "source" && selectedBaseSource && latest.projectState
	        ? buildBlueprintForAmountSource(selectedBaseSource, latest.projectState)
	        : request.baseMode === "h200"
	          ? h200
	          : {
	              ...h200,
	              parameters: h200.parameters.filter(isAiComputeStableIctParameter),
	              revenueItems: [],
	              costItems: [],
	              mappings: [],
	            };
	    const blueprint = {
	      ...base,
	      id: sourceId,
	      scenarioId: sourceId,
	      name,
	      syncState: undefined,
	    };
	    const now = new Date().toISOString();
	    const draft: IntelligentAmountSource = {
	      id: sourceId,
	      projectId: latest.projectId!,
	      name: blueprint.name,
	      description: blueprint.description || null,
	      enabled: true,
	      sourceVersion: 0,
      metadata: { scenarioId: sourceId },
      parameterGroups: blueprint.parameterGroups,
      parameters: blueprint.parameters,
      revenueItems: blueprint.revenueItems,
      costItems: blueprint.costItems,
      mappings: blueprint.mappings,
      calculationSnapshot: {},
      createdAt: now,
      updatedAt: now,
    };
    try {
      const saved = await intelligentComputeService.saveAmountSource(latest.projectId!, draft, 0);
      const projectState = await intelligentComputeService.saveProjectState(latest.projectId!, {
        expectedVersion: latest.projectState!.stateVersion,
        activeAmountSourceId: saved.id,
        projectYears: latest.projectState!.projectYears,
        discountRate: latest.projectState!.discountRate,
      });
      set(current => ({
        amountSources: [...current.amountSources.map(source => ({ ...source, enabled: false })), saved],
        activeAmountSourceId: saved.id,
        blueprint: buildBlueprintForAmountSource(saved, projectState),
        projectState,
        isDirty: false,
        lastSavedAt: saved.updatedAt,
        error: null,
      }));
      return true;
    } catch (error) {
      set({ error: `新增金额来源失败：${String(error)}` });
      return false;
    }
  },

  deleteAmountSource: async sourceId => {
    const state = get();
    if (!state.projectId || state.amountSources.length <= 1) {
      set({ error: "智算项目至少保留一个金额来源。" });
      return false;
    }
    try {
      await intelligentComputeService.deleteAmountSource(state.projectId, sourceId);
      await get().load(state.projectId);
      return true;
    } catch (error) {
      set({ error: `删除金额来源失败：${String(error)}` });
      return false;
    }
  },

  updateAmountSourceMeta: (sourceId, patch) => set(state => {
    const amountSources = state.amountSources.map(source =>
      source.id === sourceId
        ? { ...source, ...patch }
        : patch.enabled
          ? { ...source, enabled: false }
          : source
    );
    const active = sourceId === state.activeAmountSourceId;
    return {
      amountSources,
      blueprint: active && patch.name
        ? { ...state.blueprint, name: patch.name }
        : state.blueprint,
      isDirty: active ? true : state.isDirty,
      error: null,
    };
  }),

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
    return {
      blueprint: recalculate({
        ...state.blueprint,
        mappings: mapping
          ? [...state.blueprint.mappings.filter(item => item.lineItemId !== lineItemId), mapping]
          : state.blueprint.mappings.filter(item => item.lineItemId !== lineItemId),
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
