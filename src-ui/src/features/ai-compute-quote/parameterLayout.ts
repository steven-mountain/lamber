import type {
  AiComputeQuoteBlueprint,
  AiComputeQuoteParameter,
  AiComputeQuoteParameterCategory,
  AiComputeQuoteParameterGroup,
} from "./types";
import {
  AI_COMPUTE_DEFAULT_DISCOUNT_RATE_PERCENT,
  AI_COMPUTE_DISCOUNT_RATE_PARAMETER_ID,
  AI_COMPUTE_DISCOUNT_RATE_PARAMETER_KEY,
  AI_COMPUTE_PROJECT_CYCLE_PARAMETER_ID,
  AI_COMPUTE_PROJECT_CYCLE_PARAMETER_KEY,
  normalizeAiComputeDiscountRatePercent,
  normalizeAiComputeProjectCycleValue,
} from "./fundingPlans";

export const PARAMETER_GROUP_IDS = {
  scale: "parameter-group-scale",
  pricing: "parameter-group-pricing",
  investment: "parameter-group-investment",
  operations: "parameter-group-operations",
  finance: "parameter-group-finance",
  unclassified: "parameter-group-unclassified",
} as const;

export const DEFAULT_PARAMETER_GROUPS: AiComputeQuoteParameterGroup[] = [
  {
    id: PARAMETER_GROUP_IDS.scale,
    name: "规模类",
    description: "决定业务体量与服务周期",
    builtin: true,
  },
  {
    id: PARAMETER_GROUP_IDS.pricing,
    name: "定价类",
    description: "决定收入水平与报价结构",
    builtin: true,
  },
  {
    id: PARAMETER_GROUP_IDS.investment,
    name: "投入类",
    description: "决定建设和资源投入规模",
    builtin: true,
  },
  {
    id: PARAMETER_GROUP_IDS.operations,
    name: "运营类",
    description: "决定持续运营与维护成本",
    builtin: true,
  },
  {
    id: PARAMETER_GROUP_IDS.finance,
    name: "财务类",
    description: "决定资金占用与收益评价",
    builtin: true,
  },
  {
    id: PARAMETER_GROUP_IDS.unclassified,
    name: "未分类",
    description: "等待归类的参数",
    builtin: true,
  },
];

export const DEFAULT_KEY_PARAMETER_KEYS = new Set([
  "device_count",
  "years",
  "discount_rate",
  "gpu_service_price",
  "bandwidth_revenue_price",
  "cabinet_revenue_price",
  "machine_price",
  "network_price",
  "cabinet_cost_price",
  "power_kw_per_device",
  "maintenance_price",
  "bandwidth_per_device",
  "capital_rate",
]);

const LEGACY_CATEGORY_GROUP: Record<AiComputeQuoteParameterCategory, string> = {
  scale: PARAMETER_GROUP_IDS.scale,
  price: PARAMETER_GROUP_IDS.pricing,
  cost: PARAMETER_GROUP_IDS.investment,
  finance: PARAMETER_GROUP_IDS.finance,
  technical: PARAMETER_GROUP_IDS.operations,
  custom: PARAMETER_GROUP_IDS.unclassified,
};

function cloneDefaultGroups() {
  return DEFAULT_PARAMETER_GROUPS.map(group => ({ ...group }));
}

function normalizeGroups(value: unknown) {
  if (!Array.isArray(value)) return cloneDefaultGroups();
  const seen = new Set<string>();
  const groups: AiComputeQuoteParameterGroup[] = value.flatMap(candidate => {
    if (!candidate || typeof candidate !== "object") return [];
    const source = candidate as Partial<AiComputeQuoteParameterGroup>;
    const id = typeof source.id === "string" ? source.id.trim() : "";
    const name = typeof source.name === "string" ? source.name.trim() : "";
    if (!id || !name || seen.has(id)) return [];
    seen.add(id);
    return [{
      id,
      name,
      description: typeof source.description === "string" ? source.description.trim() : undefined,
      builtin: Boolean(source.builtin),
    }];
  });

  DEFAULT_PARAMETER_GROUPS.forEach(defaultGroup => {
    const existing = groups.find(group => group.id === defaultGroup.id);
    if (existing) {
      existing.builtin = true;
      if (!existing.description) existing.description = defaultGroup.description;
      return;
    }
    groups.push({ ...defaultGroup });
  });
  return groups;
}

function resolveLegacyGroupId(parameter: AiComputeQuoteParameter, validGroupIds: Set<string>) {
  if (parameter.groupId && validGroupIds.has(parameter.groupId)) return parameter.groupId;
  if (parameter.category && LEGACY_CATEGORY_GROUP[parameter.category]) {
    return LEGACY_CATEGORY_GROUP[parameter.category];
  }
  return PARAMETER_GROUP_IDS.unclassified;
}

function normalizeProjectCycleParameter(
  parameters: AiComputeQuoteParameter[],
): AiComputeQuoteParameter[] {
  const projectCycleIndex = parameters.findIndex(parameter =>
    parameter.id === AI_COMPUTE_PROJECT_CYCLE_PARAMETER_ID
  );
  const fallbackIndex = projectCycleIndex >= 0
    ? projectCycleIndex
    : parameters.findIndex(parameter => parameter.key === AI_COMPUTE_PROJECT_CYCLE_PARAMETER_KEY);
  if (fallbackIndex < 0) {
    return [
      {
        id: AI_COMPUTE_PROJECT_CYCLE_PARAMETER_ID,
        name: "项目周期",
        key: AI_COMPUTE_PROJECT_CYCLE_PARAMETER_KEY,
        value: 1,
        unit: "年",
        category: "scale",
        groupId: PARAMETER_GROUP_IDS.scale,
        isKey: true,
        sensitivityEnabled: true,
        locked: true,
      },
      ...parameters,
    ];
  }
  return parameters.map((parameter, index) => {
    if (index !== fallbackIndex) {
      return parameter.key === AI_COMPUTE_PROJECT_CYCLE_PARAMETER_KEY
        ? { ...parameter, key: `${parameter.key}_custom` }
        : parameter;
    }
    return {
      ...parameter,
      id: AI_COMPUTE_PROJECT_CYCLE_PARAMETER_ID,
      key: AI_COMPUTE_PROJECT_CYCLE_PARAMETER_KEY,
      value: normalizeAiComputeProjectCycleValue(parameter.value),
      groupId: PARAMETER_GROUP_IDS.scale,
      isKey: true,
      locked: true,
    };
  });
}

function normalizeDiscountRateParameter(
  parameters: AiComputeQuoteParameter[],
  fallbackDiscountRatePercent: number,
): AiComputeQuoteParameter[] {
  const discountRateIndex = parameters.findIndex(parameter =>
    parameter.id === AI_COMPUTE_DISCOUNT_RATE_PARAMETER_ID
  );
  const fallbackIndex = discountRateIndex >= 0
    ? discountRateIndex
    : parameters.findIndex(parameter => parameter.key === AI_COMPUTE_DISCOUNT_RATE_PARAMETER_KEY);
  if (fallbackIndex < 0) {
    return [
      ...parameters,
      {
        id: AI_COMPUTE_DISCOUNT_RATE_PARAMETER_ID,
        name: "项目折现率",
        key: AI_COMPUTE_DISCOUNT_RATE_PARAMETER_KEY,
        value: normalizeAiComputeDiscountRatePercent(fallbackDiscountRatePercent),
        unit: "%",
        category: "finance",
        groupId: PARAMETER_GROUP_IDS.finance,
        isKey: true,
        sensitivityEnabled: true,
        locked: true,
      },
    ];
  }
  return parameters.map((parameter, index) => {
    if (index !== fallbackIndex) {
      return parameter.key === AI_COMPUTE_DISCOUNT_RATE_PARAMETER_KEY
        ? { ...parameter, key: `${parameter.key}_custom` }
        : parameter;
    }
    return {
      ...parameter,
      id: AI_COMPUTE_DISCOUNT_RATE_PARAMETER_ID,
      name: parameter.name || "项目折现率",
      key: AI_COMPUTE_DISCOUNT_RATE_PARAMETER_KEY,
      value: normalizeAiComputeDiscountRatePercent(parameter.value),
      unit: "%",
      groupId: PARAMETER_GROUP_IDS.finance,
      isKey: true,
      locked: true,
    };
  });
}

export function normalizeAiComputeParameterLayout(
  blueprint: AiComputeQuoteBlueprint,
  options: { fallbackDiscountRatePercent?: number } = {},
): AiComputeQuoteBlueprint {
  const parameterGroups = normalizeGroups(blueprint.parameterGroups);
  const validGroupIds = new Set(parameterGroups.map(group => group.id));
  const parameters = normalizeDiscountRateParameter(
    normalizeProjectCycleParameter(blueprint.parameters),
    options.fallbackDiscountRatePercent ?? AI_COMPUTE_DEFAULT_DISCOUNT_RATE_PERCENT,
  );
  return {
    ...blueprint,
    parameterGroups,
    parameters: parameters.map(parameter => ({
      ...parameter,
      groupId: resolveLegacyGroupId(parameter, validGroupIds),
      isKey: typeof parameter.isKey === "boolean"
        ? parameter.isKey
        : DEFAULT_KEY_PARAMETER_KEYS.has(parameter.key),
    })),
  };
}

export function initializeAiComputeDiscountRate(
  blueprint: AiComputeQuoteBlueprint,
  discountRatePercent: number,
): AiComputeQuoteBlueprint {
  const value = normalizeAiComputeDiscountRatePercent(discountRatePercent);
  return {
    ...blueprint,
    parameters: blueprint.parameters.map(parameter =>
      parameter.id === AI_COMPUTE_DISCOUNT_RATE_PARAMETER_ID
        ? { ...parameter, value }
        : parameter
    ),
  };
}

export function migrateAiComputeDiscountRateOwnership(
  blueprint: AiComputeQuoteBlueprint,
  persistedVersion: number,
  currentProjectDiscountRatePercent: number,
): AiComputeQuoteBlueprint {
  return persistedVersion < 4
    ? initializeAiComputeDiscountRate(blueprint, currentProjectDiscountRatePercent)
    : blueprint;
}

function moveArrayItem<T>(items: T[], fromIndex: number, toIndex: number) {
  if (fromIndex < 0 || toIndex < 0 || fromIndex === toIndex) return items;
  const next = [...items];
  const [moved] = next.splice(fromIndex, 1);
  next.splice(toIndex, 0, moved);
  return next;
}

export function reorderAiComputeParameterGroup(
  groups: AiComputeQuoteParameterGroup[],
  groupId: string,
  targetGroupId: string,
) {
  const fromIndex = groups.findIndex(group => group.id === groupId);
  if (fromIndex < 0 || groupId === targetGroupId) return groups;
  const moved = groups[fromIndex];
  const remaining = groups.filter(group => group.id !== groupId);
  const targetIndex = remaining.findIndex(group => group.id === targetGroupId);
  if (targetIndex < 0) return groups;
  return [...remaining.slice(0, targetIndex), moved, ...remaining.slice(targetIndex)];
}

export function moveAiComputeParameterGroupByOffset(
  groups: AiComputeQuoteParameterGroup[],
  groupId: string,
  offset: -1 | 1,
) {
  const fromIndex = groups.findIndex(group => group.id === groupId);
  const toIndex = Math.max(0, Math.min(groups.length - 1, fromIndex + offset));
  return moveArrayItem(groups, fromIndex, toIndex);
}

export function moveAiComputeParameter(
  parameters: AiComputeQuoteParameter[],
  parameterId: string,
  targetGroupId: string,
  targetParameterId?: string | null,
) {
  const sourceIndex = parameters.findIndex(parameter => parameter.id === parameterId);
  if (sourceIndex < 0) return parameters;
  const moved = { ...parameters[sourceIndex], groupId: targetGroupId };
  const remaining = parameters.filter(parameter => parameter.id !== parameterId);

  if (targetParameterId) {
    const targetIndex = remaining.findIndex(parameter =>
      parameter.id === targetParameterId && parameter.groupId === targetGroupId
    );
    if (targetIndex >= 0) {
      return [...remaining.slice(0, targetIndex), moved, ...remaining.slice(targetIndex)];
    }
  }

  let insertIndex = -1;
  remaining.forEach((parameter, index) => {
    if (parameter.groupId === targetGroupId) insertIndex = index + 1;
  });
  if (insertIndex < 0) return [...remaining, moved];
  return [...remaining.slice(0, insertIndex), moved, ...remaining.slice(insertIndex)];
}

export function moveAiComputeParameterByOffset(
  parameters: AiComputeQuoteParameter[],
  parameterId: string,
  offset: -1 | 1,
) {
  const source = parameters.find(parameter => parameter.id === parameterId);
  if (!source?.groupId) return parameters;
  const siblings = parameters.filter(parameter => parameter.groupId === source.groupId);
  const sourceSiblingIndex = siblings.findIndex(parameter => parameter.id === parameterId);
  const targetSiblingIndex = sourceSiblingIndex + offset;
  if (targetSiblingIndex < 0 || targetSiblingIndex >= siblings.length) return parameters;
  return moveAiComputeParameter(
    parameters,
    parameterId,
    source.groupId,
    offset < 0 ? siblings[targetSiblingIndex].id : siblings[targetSiblingIndex + 1]?.id,
  );
}

export function canDeleteAiComputeParameterGroup(
  group: AiComputeQuoteParameterGroup,
  parameters: AiComputeQuoteParameter[],
) {
  return !group.builtin && !parameters.some(parameter => parameter.groupId === group.id);
}
