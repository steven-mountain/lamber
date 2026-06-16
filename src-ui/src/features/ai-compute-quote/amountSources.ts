import type { IntelligentAmountSource } from "./types";

export type CreateAmountSourceBaseMode = "blank" | "h200" | "current" | "source";

const H200_BASELINE_SOURCE_ROLE = "h200_baseline";
const H200_BASELINE_DEFAULT_DESCRIPTION = "智算项目默认金额来源";
const H200_BASELINE_PRESET_DESCRIPTION = "64 台 H200、5 年服务期的标准报价预设。金额口径为元、含税。";

export function getDefaultCreateAmountSourceBaseMode(): CreateAmountSourceBaseMode {
  return "h200";
}

export function isH200BaselineAmountSource(
  source: Pick<IntelligentAmountSource, "id" | "metadata" | "description" | "createdAt">,
  sources?: Array<Pick<IntelligentAmountSource, "id" | "createdAt">>,
) {
  if (source.metadata?.sourceRole === H200_BASELINE_SOURCE_ROLE) return true;
  const hasLegacyBaselineDescription = source.description === H200_BASELINE_DEFAULT_DESCRIPTION
    || source.description === H200_BASELINE_PRESET_DESCRIPTION;
  if (!hasLegacyBaselineDescription) return false;
  if (!sources || sources.length === 0) return true;
  const firstSource = [...sources].sort((left, right) => (
    String(left.createdAt || "").localeCompare(String(right.createdAt || ""))
  ))[0];
  return firstSource?.id === source.id;
}

export function canDeleteIntelligentAmountSource(
  sources: IntelligentAmountSource[],
  sourceId: string | null,
) {
  if (!sourceId || sources.length <= 1) return false;
  const source = sources.find(item => item.id === sourceId);
  return Boolean(source && !isH200BaselineAmountSource(source, sources));
}
