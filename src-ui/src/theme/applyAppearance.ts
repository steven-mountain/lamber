import { AppearanceSettings, FONT_SCALES } from "./appearance";
import { LIGHT_THEMES, getDarkTheme } from "./presets";
import { deriveAccentTokens } from "./deriveAccentTokens";

export function applyAppearance(settings: AppearanceSettings, resolvedColorMode: "light" | "dark") {
  if (typeof document === "undefined") return;
  const doc = document.documentElement;

  // 1. Set theme preset data attributes
  doc.setAttribute("data-theme", settings.themePreset);
  doc.setAttribute("data-color-mode", settings.colorMode);
  doc.setAttribute("data-density", settings.density);
  doc.setAttribute("data-contrast", settings.contrastPreference);
  
  // Also set dark class on document for Tailwind dark mode class compatibility (dark:...)
  if (resolvedColorMode === "dark") {
    doc.classList.add("dark");
  } else {
    doc.classList.remove("dark");
  }

  // 2. Load the colors based on preset and resolved mode
  let colors = { ...(resolvedColorMode === "dark" 
    ? getDarkTheme(settings.themePreset)
    : LIGHT_THEMES[settings.themePreset]) };

  // 3. Layer Custom Accent Overrides
  const isHighContrast = settings.contrastPreference === "high";
  if (settings.customAccent?.enabled && settings.customAccent?.value) {
    const customAccentTokens = deriveAccentTokens(
      settings.customAccent.value,
      resolvedColorMode === "dark",
      isHighContrast
    );
    colors.primary = customAccentTokens.primary;
    colors.primaryForeground = customAccentTokens.primaryForeground;
    colors.primarySoft = customAccentTokens.primarySoft;
    colors.ring = customAccentTokens.ring;
    colors.accent = customAccentTokens.accent;
    colors.accentForeground = customAccentTokens.accentForeground;
  }

  // 4. Layer High Contrast Overrides
  if (isHighContrast) {
    if (resolvedColorMode === "dark") {
      colors.background = "0 0% 0%"; // pure black
      colors.foreground = "0 0% 100%"; // pure white
      colors.card = "222 47% 5%";
      colors.popover = "222 47% 5%";
      colors.muted = "222 47% 12%";
      colors.mutedForeground = "210 40% 90%";
      colors.secondary = "222 47% 12%";
      colors.secondaryForeground = "0 0% 100%";
      colors.border = "217 32% 60%";
      colors.input = "217 32% 60%";
      colors.success = "142 80% 65%";
      colors.successSoft = "142 70% 15%";
      colors.warning = "38 90% 70%";
      colors.warningSoft = "38 92% 15%";
      colors.destructive = "2 75% 65%";
      colors.destructiveSoft = "2 45% 15%";
      
      if (!settings.customAccent?.enabled) {
        colors.primary = "217 91% 75%";
        colors.primaryForeground = "222 47% 5%";
        colors.ring = "217 91% 75%";
        colors.primarySoft = "217 91% 25%";
      }
    } else {
      colors.background = "0 0% 100%"; // pure white
      colors.foreground = "222 47% 2%"; // near black
      colors.card = "0 0% 100%";
      colors.popover = "0 0% 100%";
      colors.muted = "220 10% 96%";
      colors.mutedForeground = "220 10% 25%";
      colors.secondary = "220 10% 92%";
      colors.secondaryForeground = "222 47% 2%";
      colors.border = "222 47% 30%"; // visible border line
      colors.input = "222 47% 30%";
      colors.success = "142 80% 30%";
      colors.successSoft = "142 70% 93%";
      colors.warning = "38 90% 40%";
      colors.warningSoft = "38 92% 92%";
      colors.destructive = "2 70% 35%";
      colors.destructiveSoft = "2 45% 92%";

      if (!settings.customAccent?.enabled) {
        colors.primary = "221 83% 40%";
        colors.primaryForeground = "210 40% 98%";
        colors.ring = "221 83% 40%";
        colors.primarySoft = "221 83% 92%";
      }
    }
  }

  // 5. Write each color token as a CSS variable
  doc.style.setProperty("--background", colors.background);
  doc.style.setProperty("--foreground", colors.foreground);
  doc.style.setProperty("--card", colors.card);
  doc.style.setProperty("--card-foreground", colors.cardForeground);
  doc.style.setProperty("--popover", colors.popover);
  doc.style.setProperty("--popover-foreground", colors.popoverForeground);
  doc.style.setProperty("--primary", colors.primary);
  doc.style.setProperty("--primary-foreground", colors.primaryForeground);
  doc.style.setProperty("--primary-soft", colors.primarySoft);
  doc.style.setProperty("--secondary", colors.secondary);
  doc.style.setProperty("--secondary-foreground", colors.secondaryForeground);
  doc.style.setProperty("--muted", colors.muted);
  doc.style.setProperty("--muted-foreground", colors.mutedForeground);
  doc.style.setProperty("--accent", colors.accent);
  doc.style.setProperty("--accent-foreground", colors.accentForeground);
  doc.style.setProperty("--border", colors.border);
  doc.style.setProperty("--input", colors.input);
  doc.style.setProperty("--ring", colors.ring);
  doc.style.setProperty("--success", colors.success);
  doc.style.setProperty("--success-foreground", colors.successForeground);
  doc.style.setProperty("--success-soft", colors.successSoft);
  doc.style.setProperty("--warning", colors.warning);
  doc.style.setProperty("--warning-foreground", colors.warningForeground);
  doc.style.setProperty("--warning-soft", colors.warningSoft);
  doc.style.setProperty("--destructive", colors.destructive);
  doc.style.setProperty("--destructive-foreground", colors.destructiveForeground);
  doc.style.setProperty("--destructive-soft", colors.destructiveSoft);

  // 6. Set font-scale
  const resolvedScale = FONT_SCALES[settings.fontScale] || 1.00;
  doc.style.setProperty("--font-scale", String(resolvedScale));
}

