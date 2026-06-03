import { ThemePreset } from "./appearance";

export interface ThemeColors {
  background: string;
  foreground: string;
  card: string;
  cardForeground: string;
  popover: string;
  popoverForeground: string;
  primary: string;
  primaryForeground: string;
  primarySoft: string;
  secondary: string;
  secondaryForeground: string;
  muted: string;
  mutedForeground: string;
  accent: string;
  accentForeground: string;
  border: string;
  input: string;
  ring: string;
  success: string;
  successForeground: string;
  successSoft: string;
  warning: string;
  warningForeground: string;
  warningSoft: string;
  destructive: string;
  destructiveForeground: string;
  destructiveSoft: string;
}

export const LIGHT_THEMES: Record<ThemePreset, ThemeColors> = {
  lamber: {
    background: "210 40% 98%", // slate-50 #f8fafc
    foreground: "222 47% 11%", // slate-900 #0f172a
    card: "0 0% 100%", // white
    cardForeground: "222 47% 11%",
    popover: "0 0% 100%",
    popoverForeground: "222 47% 11%",
    primary: "221 83% 53%", // blue-600 #2563eb
    primaryForeground: "210 40% 98%",
    primarySoft: "221 83% 95%",
    secondary: "210 40% 96%", // slate-100 #f1f5f9
    secondaryForeground: "215 16% 47%", // slate-500
    muted: "210 40% 96%",
    mutedForeground: "215 16% 47%",
    accent: "210 40% 96%",
    accentForeground: "222 47% 11%",
    border: "214 32% 91%", // slate-200
    input: "213 27% 84%", // slate-300
    ring: "217 91% 60%", // blue-500
    success: "142 70% 45%",
    successForeground: "142 76% 20%",
    successSoft: "142 70% 96%",
    warning: "38 92% 50%",
    warningForeground: "38 92% 20%",
    warningSoft: "38 92% 96%",
    destructive: "2 45% 43%",
    destructiveForeground: "0 0% 100%",
    destructiveSoft: "2 45% 95%",
  },
  graphite: {
    background: "220 10% 97%", // light cool grey
    foreground: "220 20% 15%", // very dark graphite
    card: "0 0% 100%",
    cardForeground: "220 20% 15%",
    popover: "0 0% 100%",
    popoverForeground: "220 20% 15%",
    primary: "220 15% 25%", // dark graphite accent
    primaryForeground: "210 10% 98%",
    primarySoft: "220 10% 90%",
    secondary: "210 10% 94%",
    secondaryForeground: "220 10% 45%",
    muted: "210 10% 94%",
    mutedForeground: "220 10% 45%",
    accent: "210 10% 94%",
    accentForeground: "220 20% 15%",
    border: "220 10% 88%",
    input: "220 10% 82%",
    ring: "220 15% 25%",
    success: "142 60% 40%",
    successForeground: "142 60% 15%",
    successSoft: "142 40% 94%",
    warning: "35 80% 45%",
    warningForeground: "35 80% 15%",
    warningSoft: "35 60% 94%",
    destructive: "0 50% 40%",
    destructiveForeground: "0 0% 100%",
    destructiveSoft: "0 40% 94%",
  },
  navy: {
    background: "218 25% 97%", // cool slate
    foreground: "222 47% 11%",
    card: "0 0% 100%",
    cardForeground: "222 47% 11%",
    popover: "0 0% 100%",
    popoverForeground: "222 47% 11%",
    primary: "222 47% 31%", // navy blue
    primaryForeground: "210 40% 98%",
    primarySoft: "218 30% 92%",
    secondary: "214 20% 94%",
    secondaryForeground: "215 16% 47%",
    muted: "214 20% 94%",
    mutedForeground: "215 16% 47%",
    accent: "214 20% 94%",
    accentForeground: "222 47% 11%",
    border: "214 25% 88%",
    input: "214 20% 80%",
    ring: "222 47% 31%",
    success: "142 65% 42%",
    successForeground: "142 76% 18%",
    successSoft: "142 60% 94%",
    warning: "38 85% 48%",
    warningForeground: "38 85% 18%",
    warningSoft: "38 80% 94%",
    destructive: "2 45% 43%",
    destructiveForeground: "0 0% 100%",
    destructiveSoft: "2 45% 95%",
  },
  forest: {
    background: "120 15% 98%", // very light greenish grey
    foreground: "140 25% 10%",
    card: "0 0% 100%",
    cardForeground: "140 25% 10%",
    popover: "0 0% 100%",
    popoverForeground: "140 25% 10%",
    primary: "142 50% 30%", // forest green
    primaryForeground: "120 20% 98%",
    primarySoft: "142 30% 93%",
    secondary: "120 15% 94%",
    secondaryForeground: "140 15% 40%",
    muted: "120 15% 94%",
    mutedForeground: "140 15% 40%",
    accent: "120 15% 94%",
    accentForeground: "140 25% 10%",
    border: "120 15% 88%",
    input: "120 15% 80%",
    ring: "142 50% 30%",
    success: "142 60% 35%",
    successForeground: "142 70% 15%",
    successSoft: "142 40% 94%",
    warning: "38 80% 45%",
    warningForeground: "38 80% 15%",
    warningSoft: "38 70% 94%",
    destructive: "2 45% 43%",
    destructiveForeground: "0 0% 100%",
    destructiveSoft: "2 45% 95%",
  },
  warmStone: {
    background: "30 15% 97%", // warm stone / linen
    foreground: "30 25% 12%",
    card: "0 0% 100%",
    cardForeground: "30 25% 12%",
    popover: "0 0% 100%",
    popoverForeground: "30 25% 12%",
    primary: "30 20% 30%", // warm grey-brown stone
    primaryForeground: "30 10% 98%",
    primarySoft: "30 20% 91%",
    secondary: "30 15% 93%",
    secondaryForeground: "30 15% 45%",
    muted: "30 15% 93%",
    mutedForeground: "30 15% 45%",
    accent: "30 15% 93%",
    accentForeground: "30 25% 12%",
    border: "30 15% 88%",
    input: "30 15% 80%",
    ring: "30 20% 30%",
    success: "142 50% 38%",
    successForeground: "142 60% 15%",
    successSoft: "142 40% 94%",
    warning: "38 80% 45%",
    warningForeground: "38 80% 15%",
    warningSoft: "38 70% 94%",
    destructive: "2 45% 43%",
    destructiveForeground: "0 0% 100%",
    destructiveSoft: "2 45% 95%",
  },
};

export const DARK_THEMES: Record<ThemePreset, ThemeColors> = {
  lamber: {
    background: "222 47% 11%", // #0f172a
    foreground: "210 40% 98%", // #f8fafc
    card: "217 32% 17%", // #1e293b
    cardForeground: "210 40% 98%",
    popover: "217 32% 17%",
    popoverForeground: "210 40% 98%",
    primary: "217 91% 60%", // #3b82f6
    primaryForeground: "222 47% 11%",
    primarySoft: "217 91% 20%",
    secondary: "217 32% 17%",
    secondaryForeground: "215 20% 65%",
    muted: "217 32% 17%",
    mutedForeground: "215 20% 65%",
    accent: "217 32% 20%",
    accentForeground: "210 40% 98%",
    border: "217 32% 20%",
    input: "217 32% 25%",
    ring: "217 91% 60%",
    success: "142 70% 45%",
    successForeground: "142 70% 90%",
    successSoft: "142 70% 15%",
    warning: "38 92% 50%",
    warningForeground: "38 92% 90%",
    warningSoft: "38 92% 15%",
    destructive: "2 45% 43%",
    destructiveForeground: "2 45% 90%",
    destructiveSoft: "2 45% 15%",
  },
  graphite: {
    background: "220 10% 8%", // very dark neutral grey
    foreground: "210 10% 98%",
    card: "220 10% 13%",
    cardForeground: "210 10% 98%",
    popover: "220 10% 13%",
    popoverForeground: "210 10% 98%",
    primary: "220 10% 70%", // cold grey-silver
    primaryForeground: "220 10% 10%",
    primarySoft: "220 10% 25%",
    secondary: "220 10% 13%",
    secondaryForeground: "220 10% 60%",
    muted: "220 10% 13%",
    mutedForeground: "220 10% 60%",
    accent: "220 10% 18%",
    accentForeground: "210 10% 98%",
    border: "220 10% 18%",
    input: "220 10% 22%",
    ring: "220 10% 70%",
    success: "142 60% 40%",
    successForeground: "142 60% 90%",
    successSoft: "142 50% 15%",
    warning: "35 80% 45%",
    warningForeground: "35 80% 90%",
    warningSoft: "35 70% 15%",
    destructive: "0 50% 40%",
    destructiveForeground: "0 50% 90%",
    destructiveSoft: "0 40% 15%",
  },
  navy: {
    background: "222 35% 8%", // deep navy
    foreground: "210 40% 98%",
    card: "222 35% 13%",
    cardForeground: "210 40% 98%",
    popover: "222 35% 13%",
    popoverForeground: "210 40% 98%",
    primary: "217 91% 65%",
    primaryForeground: "222 35% 8%",
    primarySoft: "217 91% 20%",
    secondary: "222 35% 13%",
    secondaryForeground: "215 20% 65%",
    muted: "222 35% 13%",
    mutedForeground: "215 20% 65%",
    accent: "222 30% 18%",
    accentForeground: "210 40% 98%",
    border: "222 30% 18%",
    input: "222 30% 22%",
    ring: "217 91% 65%",
    success: "142 65% 42%",
    successForeground: "142 65% 90%",
    successSoft: "142 60% 15%",
    warning: "38 85% 48%",
    warningForeground: "38 85% 90%",
    warningSoft: "38 80% 15%",
    destructive: "2 45% 43%",
    destructiveForeground: "2 45% 90%",
    destructiveSoft: "2 45% 15%",
  },
  forest: {
    background: "140 20% 8%", // deep forest green grey
    foreground: "120 20% 98%",
    card: "140 18% 13%",
    cardForeground: "120 20% 98%",
    popover: "140 18% 13%",
    popoverForeground: "120 20% 98%",
    primary: "142 60% 55%",
    primaryForeground: "140 20% 8%",
    primarySoft: "142 60% 18%",
    secondary: "140 18% 13%",
    secondaryForeground: "140 15% 60%",
    muted: "140 18% 13%",
    mutedForeground: "140 15% 60%",
    accent: "140 15% 18%",
    accentForeground: "120 20% 98%",
    border: "140 15% 18%",
    input: "140 15% 22%",
    ring: "142 60% 55%",
    success: "142 60% 40%",
    successForeground: "142 60% 90%",
    successSoft: "142 40% 15%",
    warning: "38 80% 45%",
    warningForeground: "38 80% 90%",
    warningSoft: "38 70% 15%",
    destructive: "2 45% 43%",
    destructiveForeground: "2 45% 90%",
    destructiveSoft: "2 45% 15%",
  },
  warmStone: {
    background: "30 15% 8%", // warm dark slate stone
    foreground: "30 10% 98%",
    card: "30 15% 13%",
    cardForeground: "30 10% 98%",
    popover: "30 15% 13%",
    popoverForeground: "30 10% 98%",
    primary: "30 35% 65%",
    primaryForeground: "30 15% 8%",
    primarySoft: "30 35% 20%",
    secondary: "30 15% 13%",
    secondaryForeground: "30 15% 60%",
    muted: "30 15% 13%",
    mutedForeground: "30 15% 60%",
    accent: "30 12% 18%",
    accentForeground: "30 10% 98%",
    border: "30 12% 18%",
    input: "30 12% 22%",
    ring: "30 35% 65%",
    success: "142 50% 40%",
    successForeground: "142 50% 90%",
    successSoft: "142 40% 15%",
    warning: "38 80% 45%",
    warningForeground: "38 80% 90%",
    warningSoft: "38 70% 15%",
    destructive: "2 45% 43%",
    destructiveForeground: "2 45% 90%",
    destructiveSoft: "2 45% 15%",
  },
};

export function getDarkTheme(preset: ThemePreset): ThemeColors {
  return DARK_THEMES[preset] || DARK_THEMES.lamber;
}

