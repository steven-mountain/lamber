import { hexToRgb, rgbToHsl, hslToRgb, getLuminance, getContrastRatio, hslToHex } from "./colorUtils";

export interface AccentTokens {
  primary: string;
  primaryForeground: string;
  primarySoft: string;
  ring: string;
  accent: string;
  accentForeground: string;
}

// Representative backgrounds for contrast verification
const LIGHT_BG_HEX = "#FFFFFF";
const DARK_BG_HEX = "#0F172A";

export function deriveAccentTokens(
  customAccentColor: string,
  isDark: boolean,
  isHighContrast: boolean
): AccentTokens {
  const rgb = hexToRgb(customAccentColor);
  if (!rgb) {
    // Return standard fallback if invalid
    return getFallbackAccent(isDark, isHighContrast);
  }

  const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
  const minRatio = isHighContrast ? 7.0 : 4.5;

  const finalH = hsl.h;
  const finalS = hsl.s;
  let finalL = hsl.l;

  if (isDark) {
    // Dark mode: Primary needs to stand out against dark background (DARK_BG_HEX)
    const darkBgRgb = hexToRgb(DARK_BG_HEX)!;
    const darkBgLum = getLuminance(darkBgRgb.r, darkBgRgb.g, darkBgRgb.b);

    let currentRgb = hslToRgb(finalH, finalS, finalL);
    let currentLum = getLuminance(currentRgb.r, currentRgb.g, currentRgb.b);
    let ratio = getContrastRatio(currentLum, darkBgLum);

    // If contrast against dark bg is too low, lighten the color
    if (ratio < minRatio) {
      while (finalL < 100 && ratio < minRatio) {
        finalL += 1;
        currentRgb = hslToRgb(finalH, finalS, finalL);
        currentLum = getLuminance(currentRgb.r, currentRgb.g, currentRgb.b);
        ratio = getContrastRatio(currentLum, darkBgLum);
      }
    }

    // primaryForeground: text on the primary button background.
    // Determine whether light or dark text has higher contrast against primary background
    const primaryHex = hslToHex(finalH, finalS, finalL);
    const contrastWithWhite = getContrastRatioBetweenHex(primaryHex, "#FFFFFF");
    const contrastWithDarkSlate = getContrastRatioBetweenHex(primaryHex, "#0F172A");

    const primaryForeground = contrastWithWhite >= contrastWithDarkSlate
      ? "210 40% 98%" // light slate
      : "222 47% 11%"; // dark slate

    // primarySoft: low lightness, low saturation container
    const primarySoftL = isHighContrast ? 25 : 20;
    const primarySoftS = Math.min(finalS, 20);
    const primarySoft = `${finalH} ${primarySoftS}% ${primarySoftL}%`;

    // accent: hover state background
    const accentL = isHighContrast ? 28 : 22;
    const accent = `${finalH} ${Math.min(finalS, 25)}% ${accentL}%`;

    return {
      primary: `${finalH} ${finalS}% ${finalL}%`,
      primaryForeground,
      primarySoft,
      ring: `${finalH} ${finalS}% ${finalL}%`,
      accent,
      accentForeground: "210 40% 98%",
    };
  } else {
    // Light mode: Primary needs to have contrast against light background (LIGHT_BG_HEX)
    const lightBgRgb = hexToRgb(LIGHT_BG_HEX)!;
    const lightBgLum = getLuminance(lightBgRgb.r, lightBgRgb.g, lightBgRgb.b);

    let currentRgb = hslToRgb(finalH, finalS, finalL);
    let currentLum = getLuminance(currentRgb.r, currentRgb.g, currentRgb.b);
    let ratio = getContrastRatio(currentLum, lightBgLum);

    // If contrast against white bg is too low, darken the color
    if (ratio < minRatio) {
      while (finalL > 0 && ratio < minRatio) {
        finalL -= 1;
        currentRgb = hslToRgb(finalH, finalS, finalL);
        currentLum = getLuminance(currentRgb.r, currentRgb.g, currentRgb.b);
        ratio = getContrastRatio(currentLum, lightBgLum);
      }
    }

    // primaryForeground: text on the primary button background.
    const primaryHex = hslToHex(finalH, finalS, finalL);
    const contrastWithWhite = getContrastRatioBetweenHex(primaryHex, "#FFFFFF");
    const contrastWithDarkSlate = getContrastRatioBetweenHex(primaryHex, "#0F172A");

    const primaryForeground = contrastWithWhite >= contrastWithDarkSlate
      ? "210 40% 98%" // light slate
      : "222 47% 11%"; // dark slate

    // primarySoft: high lightness, low saturation container
    const primarySoftL = isHighContrast ? 92 : 95;
    const primarySoftS = Math.min(finalS, 15);
    const primarySoft = `${finalH} ${primarySoftS}% ${primarySoftL}%`;

    // accent: hover state background
    const accentL = isHighContrast ? 88 : 92;
    const accent = `${finalH} ${Math.min(finalS, 20)}% ${accentL}%`;

    return {
      primary: `${finalH} ${finalS}% ${finalL}%`,
      primaryForeground,
      primarySoft,
      ring: `${finalH} ${finalS}% ${finalL}%`,
      accent,
      accentForeground: `${finalH} ${finalS}% ${finalL}%`,
    };
  }
}

export function validateAccentColor(
  customAccentColor: string,
  isDark: boolean,
  isHighContrast: boolean
): { isValid: boolean; adjustedHex: string } {
  const rgb = hexToRgb(customAccentColor);
  if (!rgb) {
    return { isValid: false, adjustedHex: "#2563EB" };
  }

  const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
  const minRatio = isHighContrast ? 7.0 : 4.5;

  const finalH = hsl.h;
  const finalS = hsl.s;
  let finalL = hsl.l;

  let ratio = 1;
  if (isDark) {
    const darkBgRgb = hexToRgb(DARK_BG_HEX)!;
    const darkBgLum = getLuminance(darkBgRgb.r, darkBgRgb.g, darkBgRgb.b);
    const currentLum = getLuminance(rgb.r, rgb.g, rgb.b);
    ratio = getContrastRatio(currentLum, darkBgLum);

    if (ratio >= minRatio) {
      return { isValid: true, adjustedHex: customAccentColor };
    }

    // Lighten
    while (finalL < 100 && ratio < minRatio) {
      finalL += 1;
      const currentRgb = hslToRgb(finalH, finalS, finalL);
      const currentLum = getLuminance(currentRgb.r, currentRgb.g, currentRgb.b);
      ratio = getContrastRatio(currentLum, darkBgLum);
    }
  } else {
    const lightBgRgb = hexToRgb(LIGHT_BG_HEX)!;
    const lightBgLum = getLuminance(lightBgRgb.r, lightBgRgb.g, lightBgRgb.b);
    const currentLum = getLuminance(rgb.r, rgb.g, rgb.b);
    ratio = getContrastRatio(currentLum, lightBgLum);

    if (ratio >= minRatio) {
      return { isValid: true, adjustedHex: customAccentColor };
    }

    // Darken
    while (finalL > 0 && ratio < minRatio) {
      finalL -= 1;
      const currentRgb = hslToRgb(finalH, finalS, finalL);
      const currentLum = getLuminance(currentRgb.r, currentRgb.g, currentRgb.b);
      ratio = getContrastRatio(currentLum, lightBgLum);
    }
  }

  return {
    isValid: false,
    adjustedHex: hslToHex(finalH, finalS, finalL),
  };
}

function getFallbackAccent(isDark: boolean, isHighContrast: boolean): AccentTokens {
  if (isDark) {
    return {
      primary: "217 91% 60%",
      primaryForeground: "222 47% 11%",
      primarySoft: "217 91% 20%",
      ring: "217 91% 60%",
      accent: "217 32% 20%",
      accentForeground: "210 40% 98%",
    };
  } else {
    return {
      primary: isHighContrast ? "221 83% 40%" : "221 83% 53%",
      primaryForeground: "210 40% 98%",
      primarySoft: "221 83% 95%",
      ring: isHighContrast ? "221 83% 40%" : "217 91% 60%",
      accent: "210 40% 96%",
      accentForeground: "222 47% 11%",
    };
  }
}

function getContrastRatioBetweenHex(hex1: string, hex2: string): number {
  const rgb1 = hexToRgb(hex1);
  const rgb2 = hexToRgb(hex2);
  if (!rgb1 || !rgb2) return 1;
  const lum1 = getLuminance(rgb1.r, rgb1.g, rgb1.b);
  const lum2 = getLuminance(rgb2.r, rgb2.g, rgb2.b);
  return getContrastRatio(lum1, lum2);
}
