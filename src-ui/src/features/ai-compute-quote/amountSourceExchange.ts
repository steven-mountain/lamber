import {
  AI_COMPUTE_DISCOUNT_RATE_PARAMETER_ID,
  AI_COMPUTE_DISCOUNT_RATE_PARAMETER_KEY,
  AI_COMPUTE_PROJECT_CYCLE_PARAMETER_ID,
  AI_COMPUTE_PROJECT_CYCLE_PARAMETER_KEY,
  ictDiscountRateToAiComputePercent,
  normalizeAiComputeDiscountRatePercent,
  normalizeAiComputeProjectCycleValue,
} from "./fundingPlans";
import type {
  AiComputeQuoteBlueprint,
  AiComputeQuoteLineItem,
  AiComputeQuoteParameter,
  AiComputeQuoteParameterGroup,
  AiComputeQuoteSubjectMapping,
  IntelligentAmountSource,
  IntelligentComputeProjectState,
} from "./types";

export const AMOUNT_SOURCE_PACKAGE_KIND = "lamber.intelligentCompute.amountSource";
export const AMOUNT_SOURCE_PACKAGE_SCHEMA_VERSION = 1;

export interface AiComputeAmountSourcePackage {
  kind: typeof AMOUNT_SOURCE_PACKAGE_KIND;
  schemaVersion: typeof AMOUNT_SOURCE_PACKAGE_SCHEMA_VERSION;
  exportedAt: string;
  projectSettings: {
    projectYears: number;
    discountRate: number;
  };
  source: {
    name: string;
    description?: string | null;
    metadata: Record<string, unknown>;
    parameterGroups: AiComputeQuoteParameterGroup[];
    parameters: AiComputeQuoteParameter[];
    revenueItems: AiComputeQuoteLineItem[];
    costItems: AiComputeQuoteLineItem[];
    mappings: AiComputeQuoteSubjectMapping[];
    calculationSnapshot: Record<string, unknown>;
  };
}

export interface ImportedAmountSourceBuildOptions {
  sourceId: string;
  name: string;
  projectYears: number;
  discountRate: number;
}

function cloneJson<T>(value: T): T {
  return JSON.parse(JSON.stringify(value ?? null)) as T;
}

function cloneArray<T>(value: unknown): T[] {
  return Array.isArray(value) ? cloneJson(value) : [];
}

function cloneRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value)
    ? cloneJson(value as Record<string, unknown>)
    : {};
}

function sanitizeMetadata(metadata: unknown) {
  const next = cloneRecord(metadata);
  delete next.sourceRole;
  delete next.scenarioId;
  return next;
}

function sanitizeCalculationSnapshot(snapshot: unknown) {
  const next = cloneRecord(snapshot);
  delete next.syncState;
  delete next.ictResult;
  delete next.formalResult;
  return next;
}

function sanitizeLineItem(item: AiComputeQuoteLineItem): AiComputeQuoteLineItem {
  const next = cloneJson(item);
  delete next.formulaControlStatus;
  delete next.ictOverride;
  delete next.ictControlMessage;
  return next;
}

function patchProjectSettingsParameter(
  parameter: AiComputeQuoteParameter,
  projectYears: number,
  discountRate: number,
): AiComputeQuoteParameter {
  if (parameter.id === AI_COMPUTE_PROJECT_CYCLE_PARAMETER_ID || parameter.key === AI_COMPUTE_PROJECT_CYCLE_PARAMETER_KEY) {
    return {
      ...parameter,
      id: parameter.id || AI_COMPUTE_PROJECT_CYCLE_PARAMETER_ID,
      key: AI_COMPUTE_PROJECT_CYCLE_PARAMETER_KEY,
      value: normalizeAiComputeProjectCycleValue(projectYears),
      isKey: true,
      locked: true,
    };
  }
  if (parameter.id === AI_COMPUTE_DISCOUNT_RATE_PARAMETER_ID || parameter.key === AI_COMPUTE_DISCOUNT_RATE_PARAMETER_KEY) {
    return {
      ...parameter,
      id: parameter.id || AI_COMPUTE_DISCOUNT_RATE_PARAMETER_ID,
      key: AI_COMPUTE_DISCOUNT_RATE_PARAMETER_KEY,
      value: ictDiscountRateToAiComputePercent(discountRate),
      isKey: true,
      locked: true,
    };
  }
  return parameter;
}

function ensureProjectSettingsParameters(
  parameters: AiComputeQuoteParameter[],
  projectYears: number,
  discountRate: number,
) {
  const normalizedYears = normalizeAiComputeProjectCycleValue(projectYears);
  const normalizedDiscountPercent = normalizeAiComputeDiscountRatePercent(
    ictDiscountRateToAiComputePercent(discountRate),
  );
  let hasYears = false;
  let hasDiscountRate = false;
  const patched = parameters.map(parameter => {
    const next = patchProjectSettingsParameter(parameter, normalizedYears, discountRate);
    if (next.id === AI_COMPUTE_PROJECT_CYCLE_PARAMETER_ID || next.key === AI_COMPUTE_PROJECT_CYCLE_PARAMETER_KEY) {
      hasYears = true;
    }
    if (next.id === AI_COMPUTE_DISCOUNT_RATE_PARAMETER_ID || next.key === AI_COMPUTE_DISCOUNT_RATE_PARAMETER_KEY) {
      hasDiscountRate = true;
    }
    return next;
  });
  if (!hasYears) {
    patched.unshift({
      id: AI_COMPUTE_PROJECT_CYCLE_PARAMETER_ID,
      name: "项目周期",
      key: AI_COMPUTE_PROJECT_CYCLE_PARAMETER_KEY,
      value: normalizedYears,
      unit: "年",
      category: "scale",
      isKey: true,
      locked: true,
    });
  }
  if (!hasDiscountRate) {
    patched.unshift({
      id: AI_COMPUTE_DISCOUNT_RATE_PARAMETER_ID,
      name: "项目折现率",
      key: AI_COMPUTE_DISCOUNT_RATE_PARAMETER_KEY,
      value: normalizedDiscountPercent,
      unit: "%",
      category: "finance",
      isKey: true,
      locked: true,
    });
  }
  return patched;
}

export function buildAmountSourcePackage(
  source: IntelligentAmountSource,
  blueprint: AiComputeQuoteBlueprint,
  projectState: IntelligentComputeProjectState,
): AiComputeAmountSourcePackage {
  return {
    kind: AMOUNT_SOURCE_PACKAGE_KIND,
    schemaVersion: AMOUNT_SOURCE_PACKAGE_SCHEMA_VERSION,
    exportedAt: new Date().toISOString(),
    projectSettings: {
      projectYears: normalizeAiComputeProjectCycleValue(projectState.projectYears),
      discountRate: Number.isFinite(projectState.discountRate) ? projectState.discountRate : 0,
    },
    source: {
      name: blueprint.name || source.name,
      description: blueprint.description ?? source.description ?? null,
      metadata: sanitizeMetadata(source.metadata),
      parameterGroups: cloneArray(blueprint.parameterGroups),
      parameters: cloneArray(blueprint.parameters),
      revenueItems: cloneArray<AiComputeQuoteLineItem>(blueprint.revenueItems).map(sanitizeLineItem),
      costItems: cloneArray<AiComputeQuoteLineItem>(blueprint.costItems).map(sanitizeLineItem),
      mappings: cloneArray(blueprint.mappings),
      calculationSnapshot: sanitizeCalculationSnapshot(source.calculationSnapshot),
    },
  };
}

export function normalizeAmountSourcePackage(value: unknown): AiComputeAmountSourcePackage {
  if (!value || typeof value !== "object") {
    throw new Error("文件不是有效的智算金额来源包。");
  }
  const candidate = value as Partial<AiComputeAmountSourcePackage>;
  if (candidate.kind !== AMOUNT_SOURCE_PACKAGE_KIND) {
    throw new Error("文件类型不匹配，请选择智算金额来源 JSON。");
  }
  if (candidate.schemaVersion !== AMOUNT_SOURCE_PACKAGE_SCHEMA_VERSION) {
    throw new Error("金额来源文件版本不受支持。");
  }
  const source = candidate.source as Partial<AiComputeAmountSourcePackage["source"]> | undefined;
  if (!source || typeof source !== "object") {
    throw new Error("金额来源文件缺少来源数据。");
  }
  const requiredArrays: Array<keyof AiComputeAmountSourcePackage["source"]> = [
    "parameterGroups",
    "parameters",
    "revenueItems",
    "costItems",
    "mappings",
  ];
  requiredArrays.forEach(key => {
    if (!Array.isArray(source[key])) {
      throw new Error(`金额来源文件字段 ${String(key)} 不是数组。`);
    }
  });
  const projectSettings = cloneRecord(candidate.projectSettings);
  return {
    kind: AMOUNT_SOURCE_PACKAGE_KIND,
    schemaVersion: AMOUNT_SOURCE_PACKAGE_SCHEMA_VERSION,
    exportedAt: String(candidate.exportedAt || new Date().toISOString()),
    projectSettings: {
      projectYears: normalizeAiComputeProjectCycleValue(projectSettings.projectYears),
      discountRate: Number.isFinite(Number(projectSettings.discountRate))
        ? Math.max(0, Math.min(1, Number(projectSettings.discountRate)))
        : 0,
    },
    source: {
      name: String(source.name || "导入金额来源"),
      description: source.description == null ? null : String(source.description),
      metadata: sanitizeMetadata(source.metadata),
      parameterGroups: cloneArray(source.parameterGroups),
      parameters: cloneArray(source.parameters),
      revenueItems: cloneArray<AiComputeQuoteLineItem>(source.revenueItems).map(sanitizeLineItem),
      costItems: cloneArray<AiComputeQuoteLineItem>(source.costItems).map(sanitizeLineItem),
      mappings: cloneArray(source.mappings),
      calculationSnapshot: sanitizeCalculationSnapshot(source.calculationSnapshot),
    },
  };
}

export function buildBlueprintFromAmountSourcePackage(
  pkg: AiComputeAmountSourcePackage,
  options: ImportedAmountSourceBuildOptions,
): AiComputeQuoteBlueprint {
  const name = options.name.trim() || `${pkg.source.name}（导入）`;
  return {
    id: options.sourceId,
    scenarioId: options.sourceId,
    name,
    description: pkg.source.description || undefined,
    parameterGroups: cloneArray(pkg.source.parameterGroups),
    parameters: ensureProjectSettingsParameters(
      cloneArray(pkg.source.parameters),
      options.projectYears,
      options.discountRate,
    ),
    revenueItems: cloneArray<AiComputeQuoteLineItem>(pkg.source.revenueItems).map(sanitizeLineItem),
    costItems: cloneArray<AiComputeQuoteLineItem>(pkg.source.costItems).map(sanitizeLineItem),
    mappings: cloneArray(pkg.source.mappings),
    syncState: undefined,
  };
}

export function getDefaultImportedAmountSourceName(pkg: AiComputeAmountSourcePackage) {
  return `${pkg.source.name || "导入金额来源"}（导入）`;
}

export function sanitizeAmountSourceFileName(name: string) {
  const cleaned = name.replace(/[\\/:*?"<>|\s]+/g, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");
  return `${cleaned || "智算金额来源"}.json`;
}
