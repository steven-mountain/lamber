import { create } from "zustand";

/**
 * 测算行为偏好（与外观无关的计算口径开关），localStorage 持久化。
 *
 * taxInclAutoFix —— 财务口径自动修正：
 * 含税录入值在当前税率下不可精确表示时（round(不含税×(1+税率)) ≠ 录入值，
 * 如 6% 下 1038 → 1038.01），是否自动把含税改写为财务口径反推值。
 * 默认关闭：只提示不改数；智能反算等程序化写入遇到不可表示金额时明确拒绝。
 */

const STORAGE_KEY = "lamber_calc_preferences";

export interface CalcPreferences {
  taxInclAutoFix: boolean;
}

const DEFAULT_CALC_PREFERENCES: CalcPreferences = {
  taxInclAutoFix: false,
};

interface CalcPreferencesState extends CalcPreferences {
  setTaxInclAutoFix: (enabled: boolean) => void;
}

const loadPreferences = (): CalcPreferences => {
  try {
    const raw = typeof window !== "undefined" ? window.localStorage.getItem(STORAGE_KEY) : null;
    const parsed = raw ? JSON.parse(raw) : null;
    return {
      taxInclAutoFix: typeof parsed?.taxInclAutoFix === "boolean"
        ? parsed.taxInclAutoFix
        : DEFAULT_CALC_PREFERENCES.taxInclAutoFix,
    };
  } catch {
    return { ...DEFAULT_CALC_PREFERENCES };
  }
};

const persistPreferences = (preferences: CalcPreferences) => {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(preferences));
  } catch {
    // localStorage 不可用时静默降级为会话内生效
  }
};

export const useCalcPreferencesStore = create<CalcPreferencesState>(set => ({
  ...loadPreferences(),
  setTaxInclAutoFix: (enabled: boolean) => {
    set({ taxInclAutoFix: enabled });
    persistPreferences({ taxInclAutoFix: enabled });
  },
}));

/** 非 React 上下文（hooks 回调、工具函数）读取当前开关。 */
export const isTaxInclAutoFixEnabled = (): boolean =>
  useCalcPreferencesStore.getState().taxInclAutoFix;
