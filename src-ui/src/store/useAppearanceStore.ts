import { create } from "zustand";
import { 
  AppearanceSettings, 
  DEFAULT_APPEARANCE_SETTINGS, 
  ColorMode, 
  ThemePreset, 
  FontScalePreset, 
  DensityPreset,
  ContrastPreference
} from "../theme/appearance";
import { applyAppearance } from "../theme/applyAppearance";
import { emit, listen } from "@tauri-apps/api/event";

interface AppearanceState {
  settings: AppearanceSettings;
  resolvedColorMode: "light" | "dark";
  hydrated: boolean;
  
  hydrate: () => void;
  setColorMode: (mode: ColorMode) => void;
  setThemePreset: (preset: ThemePreset) => void;
  setFontScale: (scale: FontScalePreset) => void;
  setDensity: (density: DensityPreset) => void;
  setContrastPreference: (contrast: ContrastPreference) => void;
  setCustomAccent: (enabled: boolean, value: string | null) => void;
  resetAppearance: () => void;
}

const STORAGE_KEY = "lamber_appearance_settings";
const TAURI_SYNC_EVENT = "appearance-settings-updated";

function isTauriRuntime() {
  return typeof window !== "undefined" && Boolean((window as any).__TAURI_INTERNALS__);
}

function getSystemColorMode(): "light" | "dark" {
  if (typeof window === "undefined") return "light";
  return window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
}

export const useAppearanceStore = create<AppearanceState>((set, get) => {
  let mediaQueryList: MediaQueryList | null = null;
  
  const handleSystemThemeChange = (e: MediaQueryListEvent) => {
    if (get().settings.colorMode === "system") {
      const mode = e.matches ? "dark" : "light";
      set({ resolvedColorMode: mode });
      applyAppearance(get().settings, mode);
    }
  };

  const setupSystemThemeListener = () => {
    if (typeof window === "undefined") return;
    if (mediaQueryList) {
      try {
        mediaQueryList.removeEventListener("change", handleSystemThemeChange);
      } catch (err) {}
    }
    mediaQueryList = window.matchMedia("(prefers-color-scheme: dark)");
    try {
      mediaQueryList.addEventListener("change", handleSystemThemeChange);
    } catch (err) {}
  };

  const syncStateAndDOM = (newSettings: AppearanceSettings) => {
    const resolved = newSettings.colorMode === "system"
      ? getSystemColorMode()
      : (newSettings.colorMode as "light" | "dark");

    set({ settings: newSettings, resolvedColorMode: resolved });
    applyAppearance(newSettings, resolved);

    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(newSettings));
    } catch (e) {
      console.warn("Failed to write appearance settings to localStorage:", e);
    }

    if (isTauriRuntime()) {
      emit(TAURI_SYNC_EVENT, newSettings).catch(err => {
        console.warn("Failed to emit appearance sync event:", err);
      });
    }
  };

  return {
    settings: DEFAULT_APPEARANCE_SETTINGS,
    resolvedColorMode: "light",
    hydrated: false,

    hydrate: () => {
      if (get().hydrated) return;

      let loadedSettings = DEFAULT_APPEARANCE_SETTINGS;
      try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (raw) {
          const parsed = JSON.parse(raw);
          loadedSettings = {
            colorMode: ["light", "dark", "system"].includes(parsed.colorMode) ? parsed.colorMode : DEFAULT_APPEARANCE_SETTINGS.colorMode,
            themePreset: ["lamber", "graphite", "navy", "forest", "warmStone"].includes(parsed.themePreset) ? parsed.themePreset : DEFAULT_APPEARANCE_SETTINGS.themePreset,
            fontScale: ["compact", "standard", "comfortable", "large"].includes(parsed.fontScale) ? parsed.fontScale : DEFAULT_APPEARANCE_SETTINGS.fontScale,
            density: ["compact", "standard", "comfortable"].includes(parsed.density) ? parsed.density : DEFAULT_APPEARANCE_SETTINGS.density,
            contrastPreference: ["standard", "high"].includes(parsed.contrastPreference) ? parsed.contrastPreference : DEFAULT_APPEARANCE_SETTINGS.contrastPreference,
            customAccent: (parsed.customAccent && typeof parsed.customAccent.enabled === "boolean") 
              ? { enabled: parsed.customAccent.enabled, value: typeof parsed.customAccent.value === "string" ? parsed.customAccent.value : null }
              : DEFAULT_APPEARANCE_SETTINGS.customAccent,
            version: typeof parsed.version === "number" ? parsed.version : 3,
          };
        }
      } catch (e) {
        console.warn("Failed to load appearance settings, fallback to default", e);
      }

      setupSystemThemeListener();

      const resolved = loadedSettings.colorMode === "system"
        ? getSystemColorMode()
        : (loadedSettings.colorMode as "light" | "dark");

      set({ settings: loadedSettings, resolvedColorMode: resolved, hydrated: true });
      applyAppearance(loadedSettings, resolved);

      // Listen to cross-window Tauri sync events
      if (isTauriRuntime()) {
        listen<AppearanceSettings>(TAURI_SYNC_EVENT, (event) => {
          const received = event.payload;
          if (JSON.stringify(received) !== JSON.stringify(get().settings)) {
            const resolved = received.colorMode === "system"
              ? getSystemColorMode()
              : (received.colorMode as "light" | "dark");
            set({ settings: received, resolvedColorMode: resolved });
            applyAppearance(received, resolved);
            try {
              localStorage.setItem(STORAGE_KEY, JSON.stringify(received));
            } catch (e) {}
          }
        }).catch(err => {
          console.warn("Failed to listen for appearance sync event:", err);
        });
      }
    },

    setColorMode: (mode) => {
      const current = get().settings;
      if (current.colorMode === mode) return;
      syncStateAndDOM({ ...current, colorMode: mode });
    },

    setThemePreset: (preset) => {
      const current = get().settings;
      if (current.themePreset === preset) return;
      syncStateAndDOM({ ...current, themePreset: preset });
    },

    setFontScale: (scale) => {
      const current = get().settings;
      if (current.fontScale === scale) return;
      syncStateAndDOM({ ...current, fontScale: scale });
    },

    setDensity: (density) => {
      const current = get().settings;
      if (current.density === density) return;
      syncStateAndDOM({ ...current, density: density });
    },

    setContrastPreference: (contrast) => {
      const current = get().settings;
      if (current.contrastPreference === contrast) return;
      syncStateAndDOM({ ...current, contrastPreference: contrast });
    },

    setCustomAccent: (enabled, value) => {
      const current = get().settings;
      if (current.customAccent.enabled === enabled && current.customAccent.value === value) return;
      syncStateAndDOM({
        ...current,
        customAccent: { enabled, value }
      });
    },

    resetAppearance: () => {
      syncStateAndDOM(DEFAULT_APPEARANCE_SETTINGS);
    },
  };
});

