// 测算方案的甄选阶段标签（甄选前 / 甄选后）。
// 与后端 benefit_schemes.stage 字段一一对应，空值表示未标注。

export type SchemeStage = "pre_selection" | "post_selection";

export interface SchemeStageOption {
  value: SchemeStage;
  label: string;
  short: string;
  /** Tailwind 类：背景 + 文字，用于 chip 展示。遵循无边框、低饱和的设计系统。 */
  chipClass: string;
}

export const SCHEME_STAGE_OPTIONS: SchemeStageOption[] = [
  {
    value: "pre_selection",
    label: "甄选前",
    short: "甄选前",
    chipClass: "bg-amber-100 text-amber-700",
  },
  {
    value: "post_selection",
    label: "甄选后",
    short: "甄选后",
    chipClass: "bg-emerald-100 text-emerald-700",
  },
];

export function normalizeSchemeStage(stage?: string | null): SchemeStage | null {
  return stage === "pre_selection" || stage === "post_selection" ? stage : null;
}

export function getSchemeStageOption(
  stage?: string | null
): SchemeStageOption | null {
  const normalized = normalizeSchemeStage(stage);
  return normalized
    ? SCHEME_STAGE_OPTIONS.find((option) => option.value === normalized) ?? null
    : null;
}

export function getSchemeStageLabel(stage?: string | null): string {
  return getSchemeStageOption(stage)?.label ?? "未标注";
}
