export type ColorMode = "light" | "dark" | "system";

export type ThemePreset =
  | "lamber"
  | "graphite"
  | "navy"
  | "forest"
  | "warmStone";

export type FontScalePreset =
  | "compact"
  | "standard"
  | "comfortable"
  | "large";

export type DensityPreset =
  | "compact"
  | "standard"
  | "comfortable";

export type ContrastPreference = "standard" | "high";

export interface CustomAccentSettings {
  enabled: boolean;
  value: string | null;
}

export interface AppearanceSettings {
  colorMode: ColorMode;
  themePreset: ThemePreset;
  fontScale: FontScalePreset;
  density: DensityPreset;
  contrastPreference: ContrastPreference;
  aiLauncherVisible: boolean;
  customAccent: CustomAccentSettings;
  version: number;
}

export const DEFAULT_APPEARANCE_SETTINGS: AppearanceSettings = {
  colorMode: "light",
  themePreset: "lamber",
  fontScale: "standard",
  density: "standard",
  contrastPreference: "standard",
  aiLauncherVisible: false,
  customAccent: {
    enabled: false,
    value: null,
  },
  version: 4,
};

export const FONT_SCALES: Record<FontScalePreset, number> = {
  compact: 0.93,
  standard: 1.00,
  comfortable: 1.08,
  large: 1.16,
};
