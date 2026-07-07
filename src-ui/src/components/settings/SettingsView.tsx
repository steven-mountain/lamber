import { useState, useEffect, useMemo } from "react";
import { useAppearanceStore } from "../../store/useAppearanceStore";
import { useCalcPreferencesStore } from "../../store/useCalcPreferencesStore";
import { 
  ColorMode, 
  ThemePreset, 
  FontScalePreset, 
  DensityPreset,
  ContrastPreference,
  validateAccentColor,
  applyAppearance
} from "../../theme";
import AppIcon from "../icons/AppIcon";
import { Button } from "../ui/button";
import { Input } from "../ui/input";
import { Card, CardHeader, CardContent, CardTitle, CardDescription } from "../ui/card";
import { Label } from "../ui/label";

interface SettingsViewProps {
  onBack: () => void;
}

export default function SettingsView({ onBack }: SettingsViewProps) {
  const { 
    settings, 
    resolvedColorMode,
    setColorMode, 
    setThemePreset, 
    setFontScale, 
    setDensity, 
    setContrastPreference,
    setCustomAccent,
    setAiLauncherVisible,
    resetAppearance 
  } = useAppearanceStore();

  const { taxInclAutoFix, setTaxInclAutoFix } = useCalcPreferencesStore();

  const [customColorInput, setCustomColorInput] = useState(
    settings.customAccent.value || "#2563eb"
  );

  useEffect(() => {
    if (settings.customAccent.value) {
      setCustomColorInput(settings.customAccent.value);
    } else {
      setCustomColorInput("#2563eb");
    }
  }, [settings.customAccent.value]);

  // Real-time preview of the user's typed color (including self-adjusted safe color if low contrast)
  useEffect(() => {
    if (settings.customAccent.enabled) {
      let cleanColor = customColorInput.trim();
      if (cleanColor && !cleanColor.startsWith('#')) {
        cleanColor = '#' + cleanColor;
      }
      
      if (/^#[0-9A-Fa-f]{6}$/.test(cleanColor) || /^#[0-9A-Fa-f]{3}$/.test(cleanColor)) {
        const v = validateAccentColor(
          cleanColor,
          resolvedColorMode === "dark",
          settings.contrastPreference === "high"
        );
        
        const tempSettings = {
          ...settings,
          customAccent: {
            enabled: true,
            value: v.adjustedHex
          }
        };
        applyAppearance(tempSettings, resolvedColorMode);
      }
    } else {
      applyAppearance(settings, resolvedColorMode);
    }

    return () => {
      // Restore official settings on cleanup or unmount
      applyAppearance(useAppearanceStore.getState().settings, resolvedColorMode);
    };
  }, [customColorInput, settings, resolvedColorMode]);

  const recommendedColors = [
    { value: "#2563eb", name: "商务蓝" },
    { value: "#1e3a8a", name: "海军蓝" },
    { value: "#0f766e", name: "青绿色" },
    { value: "#166534", name: "森林绿" },
    { value: "#d97706", name: "琥珀色" },
    { value: "#4b5563", name: "石墨色" },
  ];

  const validation = useMemo(() => {
    let cleanColor = customColorInput.trim();
    if (cleanColor && !cleanColor.startsWith('#')) {
      cleanColor = '#' + cleanColor;
    }
    return validateAccentColor(
      cleanColor,
      resolvedColorMode === "dark",
      settings.contrastPreference === "high"
    );
  }, [customColorInput, resolvedColorMode, settings.contrastPreference]);

  const themeOptions: { value: ThemePreset; label: string; bg: string; primary: string; desc: string }[] = [
    { value: "lamber", label: "Lamber 默认", bg: "bg-slate-50", primary: "bg-blue-600", desc: "极浅冷灰背景 + 克制蓝主色" },
    { value: "graphite", label: "石墨灰", bg: "bg-[#f1f3f5]", primary: "bg-[#343a40]", desc: "中性灰背景 + 深灰强调" },
    { value: "navy", label: "海军蓝", bg: "bg-[#f4f6f9]", primary: "bg-[#1e3a8a]", desc: "冷灰背景 + 稳重深蓝强调" },
    { value: "forest", label: "森林绿", bg: "bg-[#f4f9f4]", primary: "bg-[#15803d]", desc: "中性浅背景 + 低饱和绿色强调" },
    { value: "warmStone", label: "暖石色", bg: "bg-[#f7f5f2]", primary: "bg-[#5c5346]", desc: "柔和暖灰背景 + 暖灰棕强调" },
  ];

  const modeOptions: { value: ColorMode; label: string; icon: "ai" | "settings" | "quickAction" }[] = [
    { value: "light", label: "浅色模式", icon: "quickAction" },
    { value: "dark", label: "深色模式", icon: "ai" },
    { value: "system", label: "跟随系统", icon: "settings" },
  ];

  const fontOptions: { value: FontScalePreset; label: string; scale: string }[] = [
    { value: "compact", label: "紧凑", scale: "93%" },
    { value: "standard", label: "标准", scale: "100%" },
    { value: "comfortable", label: "舒适", scale: "108%" },
    { value: "large", label: "大字号", scale: "116%" },
  ];

  const densityOptions: { value: DensityPreset; label: string; desc: string }[] = [
    { value: "compact", label: "紧凑", desc: "表格密集、信息量大" },
    { value: "standard", label: "标准", desc: "平衡舒适的默认间距" },
    { value: "comfortable", label: "宽松", desc: "适合长时间大段阅读" },
  ];


  return (
    <div className="flex flex-col flex-1 h-full overflow-hidden bg-background animate-in fade-in duration-300">
      {/* Top Header */}
      <header className="flex h-16 bg-card border-b border-border items-center px-6 gap-4 shrink-0 select-none">
        <button 
          onClick={onBack} 
          className="text-secondary-foreground hover:text-primary hover:bg-secondary font-semibold flex items-center gap-1.5 px-3 py-2 rounded-lg transition-colors text-body"
        >
          <span>←</span> 返回
        </button>
        <h2 className="m-0 text-lg font-bold text-foreground border-l-2 border-border pl-4">外观设置中心</h2>
      </header>

      {/* Main Settings Body */}
      <div className="flex-1 overflow-y-auto p-6 md:p-8">
        <div className="mx-auto max-w-6xl grid grid-cols-1 lg:grid-cols-3 gap-8">
          
          {/* Left Columns - Inputs */}
          <div className="lg:col-span-2 space-y-6">
            
            {/* Section 1: Color Mode */}
            <Card className="border border-border/40 shadow-sm">
              <CardHeader>
                <CardTitle className="text-section-title">明暗模式</CardTitle>
                <CardDescription className="text-caption">选择应用在浅色、深色模式下运行或自动跟随系统设置</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="grid grid-cols-3 gap-4">
                  {modeOptions.map(opt => {
                    const isSelected = settings.colorMode === opt.value;
                    return (
                      <button
                        key={opt.value}
                        onClick={() => setColorMode(opt.value)}
                        className={`flex flex-col items-center justify-center p-4 rounded-xl border transition-all text-center gap-2 ${
                          isSelected 
                            ? "border-primary bg-primary-soft/20 text-primary shadow-sm" 
                            : "border-border/60 hover:bg-muted text-secondary-foreground hover:text-foreground"
                        }`}
                      >
                        <AppIcon name={opt.icon} size={20} />
                        <span className="text-sm font-semibold">{opt.label}</span>
                      </button>
                    );
                  })}
                </div>
              </CardContent>
            </Card>

            {/* Section 2: Color Preset Theme */}
            <Card className="border border-border/40 shadow-sm">
              <CardHeader>
                <CardTitle className="text-section-title">主题预置方案</CardTitle>
                <CardDescription className="text-caption">选择一套受控的、长时间使用友好的企业主题预置方案</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {themeOptions.map(opt => {
                    const isSelected = settings.themePreset === opt.value;
                    return (
                      <button
                        key={opt.value}
                        onClick={() => setThemePreset(opt.value)}
                        className={`flex items-center p-4 rounded-xl border text-left gap-4 transition-all ${
                          isSelected 
                            ? "border-primary bg-primary-soft/20 text-primary shadow-sm" 
                            : "border-border/60 hover:bg-muted text-secondary-foreground hover:text-foreground"
                        }`}
                      >
                        {/* Preset preview dot */}
                        <div className={`h-12 w-12 rounded-lg ${opt.bg} border border-border flex items-center justify-center shrink-0`}>
                          <div className={`h-5 w-5 rounded-full ${opt.primary}`} />
                        </div>
                        <div className="min-w-0 flex-1">
                          <div className="text-sm font-semibold text-foreground">{opt.label}</div>
                          <div className="text-caption text-secondary-foreground truncate mt-0.5">{opt.desc}</div>
                        </div>
                        {isSelected && (
                          <div className="text-primary shrink-0 mr-1">
                            <AppIcon name="check" size={16} />
                          </div>
                        )}
                      </button>
                    );
                  })}
                </div>
              </CardContent>
            </Card>

            {/* Section 3: Contrast Preference */}
            <Card className="border border-border/40 shadow-sm">
              <CardHeader>
                <CardTitle className="text-section-title">显示对比度</CardTitle>
                <CardDescription className="text-caption">增强文字、图标及页面边界的视觉辨识度，服务长时间阅读需求</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="grid grid-cols-2 gap-4">
                  {[
                    { value: "standard", label: "标准对比度", desc: "追求色彩层级的自然与温和" },
                    { value: "high", label: "高对比度", desc: "大幅提升文本、表单边框与焦点清晰度" }
                  ].map(opt => {
                    const isSelected = settings.contrastPreference === opt.value;
                    return (
                      <button
                        key={opt.value}
                        onClick={() => setContrastPreference(opt.value as ContrastPreference)}
                        className={`flex flex-col items-center justify-center p-4 rounded-xl border transition-all text-center gap-1 ${
                          isSelected 
                            ? "border-primary bg-primary-soft/20 text-primary shadow-sm" 
                            : "border-border/60 hover:bg-muted text-secondary-foreground hover:text-foreground"
                        }`}
                      >
                        <span className="text-sm font-semibold">{opt.label}</span>
                        <span className="text-caption text-secondary-foreground mt-0.5">{opt.desc}</span>
                      </button>
                    );
                  })}
                </div>
              </CardContent>
            </Card>

            {/* Section 4: Custom Accent Color */}
            <Card className="border border-border/40 shadow-sm">
              <CardHeader>
                <CardTitle className="text-section-title">界面强调色</CardTitle>
                <CardDescription className="text-caption">选择或自定义应用于按钮、高亮状态和主标签的强调色</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                {/* Accent Mode Switcher */}
                <div className="flex gap-2 p-1 bg-muted rounded-lg border border-border/40">
                  <button
                    onClick={() => setCustomAccent(false, null)}
                    className={`flex-1 py-1.5 text-xs font-semibold rounded-md transition-all ${
                      !settings.customAccent.enabled
                        ? "bg-card text-foreground shadow-sm font-bold"
                        : "text-secondary-foreground hover:text-foreground"
                    }`}
                  >
                    使用预置主题强调色
                  </button>
                  <button
                    onClick={() => setCustomAccent(true, settings.customAccent.value || "#2563eb")}
                    className={`flex-1 py-1.5 text-xs font-semibold rounded-md transition-all ${
                      settings.customAccent.enabled
                        ? "bg-card text-foreground shadow-sm font-bold"
                        : "text-secondary-foreground hover:text-foreground"
                    }`}
                  >
                    自定义界面强调色
                  </button>
                </div>

                {settings.customAccent.enabled && (
                  <div className="space-y-4 animate-in fade-in slide-in-from-top-2 duration-200">
                    {/* Recommended Colors Grid */}
                    <div className="space-y-2">
                      <Label className="text-label">推荐高可读色彩</Label>
                      <div className="grid grid-cols-3 sm:grid-cols-6 gap-2">
                        {recommendedColors.map(color => {
                          const isSelected = settings.customAccent.value?.toLowerCase() === color.value.toLowerCase();
                          return (
                            <button
                              key={color.value}
                              onClick={() => {
                                setCustomColorInput(color.value);
                                setCustomAccent(true, color.value);
                              }}
                              className={`group relative flex h-10 w-full items-center justify-center rounded-lg border transition-all ${
                                isSelected
                                  ? "border-primary bg-primary-soft/10 ring-2 ring-primary/20"
                                  : "border-border/60 hover:bg-muted"
                              }`}
                              title={color.name}
                            >
                              <div
                                className="h-5 w-5 rounded-full shadow-inner transition-transform group-hover:scale-110"
                                style={{ backgroundColor: color.value }}
                              />
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Custom Color Inputs */}
                    <div className="space-y-2 pt-2 border-t border-border/40">
                      <Label className="text-label">自定义选色器</Label>
                      <div className="flex items-center gap-3">
                        {/* Native Color Picker */}
                        <div className="relative h-9 w-9 shrink-0 rounded-lg border border-input overflow-hidden bg-card hover:border-primary transition-colors">
                          <input
                            type="color"
                            value={customColorInput}
                            onChange={(e) => {
                              const val = e.target.value;
                              setCustomColorInput(val);
                              // Auto-validate and apply if valid
                              const v = validateAccentColor(
                                val,
                                resolvedColorMode === "dark",
                                settings.contrastPreference === "high"
                              );
                              if (v.isValid) {
                                setCustomAccent(true, val);
                              }
                            }}
                            className="absolute -inset-1 h-[200%] w-[200%] cursor-pointer border-none bg-transparent p-0"
                          />
                        </div>
                        {/* Hex Text input */}
                        <Input
                          type="text"
                          value={customColorInput}
                          onChange={(e) => {
                            const val = e.target.value;
                            setCustomColorInput(val);
                            
                            let cleanColor = val.trim();
                            if (cleanColor && !cleanColor.startsWith('#')) {
                              cleanColor = '#' + cleanColor;
                            }
                            
                            if (/^#[0-9A-Fa-f]{6}$/.test(cleanColor) || /^#[0-9A-Fa-f]{3}$/.test(cleanColor)) {
                              const v = validateAccentColor(
                                cleanColor,
                                resolvedColorMode === "dark",
                                settings.contrastPreference === "high"
                              );
                              if (v.isValid) {
                                setCustomAccent(true, cleanColor);
                              }
                            }
                          }}
                          placeholder="#2563EB"
                          className="font-mono text-sm max-w-[150px]"
                        />
                        <span className="text-caption text-secondary-foreground">
                          支持输入 6 位十六进制 RGB 颜色值
                        </span>
                      </div>
                    </div>

                    {/* Contrast Validation Banner */}
                    {!validation.isValid && (
                      <div className="p-3 rounded-xl bg-warning-soft border border-warning/20 text-xs text-warning-foreground space-y-2.5 animate-in fade-in duration-200">
                        <div className="flex items-start gap-2">
                          <AppIcon name="warning" size={14} className="text-warning shrink-0 mt-0.5" />
                          <div className="leading-relaxed">
                            <strong>可读性警示：</strong>
                            当前选择的强调色在明暗或高对比模式下对比度不足（低于最低 WCAG 对比度标准），这可能导致按钮文字或关键标识灰化、模糊。
                          </div>
                        </div>
                        <div className="flex items-center justify-between pl-6 gap-2">
                          <div className="flex items-center gap-2">
                            <span>推荐安全替代色：</span>
                            <div className="h-4.5 w-4.5 rounded-full border border-border" style={{ backgroundColor: validation.adjustedHex }} />
                            <span className="font-mono font-bold">{validation.adjustedHex}</span>
                          </div>
                          <Button
                            size="sm"
                            variant="outline"
                            onClick={() => {
                              setCustomColorInput(validation.adjustedHex);
                              setCustomAccent(true, validation.adjustedHex);
                            }}
                            className="h-7 text-[11px] px-2.5 border-warning/30 hover:bg-warning hover:text-white"
                          >
                            采用推荐安全色
                          </Button>
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </CardContent>
            </Card>

            {/* Section 5: Font Size */}
            <Card className="border border-border/40 shadow-sm">
              <CardHeader>
                <CardTitle className="text-section-title">字体大小</CardTitle>
                <CardDescription className="text-caption">按比例缩放系统中的所有标题、正文及指标的字号大小</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                  {fontOptions.map(opt => {
                    const isSelected = settings.fontScale === opt.value;
                    return (
                      <button
                        key={opt.value}
                        onClick={() => setFontScale(opt.value)}
                        className={`flex flex-col items-center justify-center p-3 rounded-xl border transition-all text-center gap-1 ${
                          isSelected 
                            ? "border-primary bg-primary-soft/20 text-primary shadow-sm" 
                            : "border-border/60 hover:bg-muted text-secondary-foreground hover:text-foreground"
                        }`}
                      >
                        <span className="text-base font-bold">{opt.label}</span>
                        <span className="text-caption text-secondary-foreground">{opt.scale}</span>
                      </button>
                    );
                  })}
                </div>
              </CardContent>
            </Card>

            {/* Section 6: Interface Density */}
            <Card className="border border-border/40 shadow-sm">
              <CardHeader>
                <CardTitle className="text-section-title">界面密度</CardTitle>
                <CardDescription className="text-caption">调整主要卡片边距、输入框高度和表格行高</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  {densityOptions.map(opt => {
                    const isSelected = settings.density === opt.value;
                    return (
                      <button
                        key={opt.value}
                        onClick={() => setDensity(opt.value)}
                        className={`flex flex-col items-center justify-center p-4 rounded-xl border transition-all text-center gap-1 ${
                          isSelected 
                            ? "border-primary bg-primary-soft/20 text-primary shadow-sm" 
                            : "border-border/60 hover:bg-muted text-secondary-foreground hover:text-foreground"
                        }`}
                      >
                        <span className="text-sm font-semibold">{opt.label}</span>
                        <span className="text-caption text-secondary-foreground mt-0.5">{opt.desc}</span>
                      </button>
                    );
                  })}
                </div>
              </CardContent>
            </Card>

            {/* Section 7: AI Floating Launcher */}
            <Card className="border border-border/40 shadow-sm">
              <CardHeader>
                <CardTitle className="text-section-title">AI 助手入口</CardTitle>
                <CardDescription className="text-caption">控制右下角 AI 助手悬浮图标是否显示</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="flex items-center justify-between gap-4 rounded-xl bg-muted/40 p-4">
                  <div className="min-w-0 space-y-1">
                    <div className="flex items-center gap-2 text-sm font-semibold text-foreground">
                      <AppIcon name="ai" size={16} className="text-primary" />
                      右下角悬浮图标
                    </div>
                    <p className="text-caption text-secondary-foreground">
                      {settings.aiLauncherVisible ? "当前会在主窗口显示 AI 助手入口。" : "当前已隐藏主窗口中的 AI 助手入口。"}
                    </p>
                  </div>
                  <button
                    type="button"
                    role="switch"
                    aria-checked={settings.aiLauncherVisible}
                    onClick={() => setAiLauncherVisible(!settings.aiLauncherVisible)}
                    className={`relative inline-flex h-7 w-12 shrink-0 items-center rounded-full p-1 transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/30 ${
                      settings.aiLauncherVisible ? "bg-primary" : "bg-muted"
                    }`}
                  >
                    <span className="sr-only">显示 AI 助手悬浮图标</span>
                    <span
                      className={`h-5 w-5 rounded-full bg-card shadow-sm transition-transform ${
                        settings.aiLauncherVisible ? "translate-x-5" : "translate-x-0"
                      }`}
                    />
                  </button>
                </div>
              </CardContent>
            </Card>

            {/* Section 8: 测算行为 */}
            <Card className="border border-border/40 shadow-sm">
              <CardHeader>
                <CardTitle className="text-section-title">测算行为</CardTitle>
                <CardDescription className="text-caption">财务口径相关的自动处理开关</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="flex items-center justify-between gap-4 rounded-xl bg-muted/40 p-4">
                  <div className="min-w-0 space-y-1">
                    <div className="flex items-center gap-2 text-sm font-semibold text-foreground">
                      <AppIcon name="quickAction" size={16} className="text-primary" />
                      财务口径自动修正含税金额
                    </div>
                    <p className="text-caption text-secondary-foreground">
                      业务系统以不含税为准反推含税。录入的含税价不可精确表示时（如 6% 税率下 1038 元对应系统口径 1038.01 元），
                      {taxInclAutoFix
                        ? "当前会在输入完成后自动改为财务口径反推值。"
                        : "当前仅提示不改数；智能反算、甄选限价回填遇到此类金额将被拒绝写入。"}
                    </p>
                  </div>
                  <button
                    type="button"
                    role="switch"
                    aria-checked={taxInclAutoFix}
                    onClick={() => setTaxInclAutoFix(!taxInclAutoFix)}
                    className={`relative inline-flex h-7 w-12 shrink-0 items-center rounded-full p-1 transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/30 ${
                      taxInclAutoFix ? "bg-primary" : "bg-muted"
                    }`}
                  >
                    <span className="sr-only">财务口径自动修正含税金额</span>
                    <span
                      className={`h-5 w-5 rounded-full bg-card shadow-sm transition-transform ${
                        taxInclAutoFix ? "translate-x-5" : "translate-x-0"
                      }`}
                    />
                  </button>
                </div>
              </CardContent>
            </Card>

            {/* Default Restoration */}
            <div className="flex justify-end pt-2">
              <Button 
                variant="outline"
                onClick={resetAppearance}
                className="flex items-center gap-2 text-destructive border-destructive/20 hover:bg-destructive-soft"
              >
                <AppIcon name="reverse" size={16} />
                恢复默认外观
              </Button>
            </div>

          </div>

          {/* Right Column - Preview Area */}
          <div className="space-y-6">
            <div className="sticky top-6">
              <Card className="border border-border/40 shadow-md">
                <CardHeader className="border-b border-border/30 bg-muted/20">
                  <CardTitle className="text-section-title flex items-center gap-2">
                    <AppIcon name="ai" size={18} className="text-primary animate-pulse" />
                    外观实时预览
                  </CardTitle>
                  <CardDescription className="text-caption">此处卡片展示了所选外观和密度的真实渲染效果</CardDescription>
                </CardHeader>
                <CardContent className="p-6 space-y-5">
                  
                  {/* Preview Section Header */}
                  <div className="space-y-1">
                    <div className="flex items-center gap-2">
                      <span className="text-label font-bold text-foreground">项目名称</span>
                      <span className="shrink-0 bg-success-soft text-success-foreground text-[10px] font-bold px-2 py-0.5 rounded-full">
                        正常
                      </span>
                    </div>
                    <p className="text-caption text-secondary-foreground">
                      当前外观主题: <span className="font-semibold text-primary">{settings.themePreset}</span>
                      {settings.customAccent.enabled && <span className="ml-1.5 text-xs text-primary-foreground bg-primary px-1.5 py-0.25 rounded-md">自定义</span>}
                      {settings.contrastPreference === "high" && <span className="ml-1.5 text-xs bg-muted text-foreground border border-border px-1.5 py-0.25 rounded-md font-bold">高对比度</span>}
                    </p>
                  </div>

                  {/* Text sizes preview */}
                  <div className="space-y-2 rounded-lg bg-muted/40 p-4 border border-border/30">
                    <div className="text-body font-semibold">文字排版预览</div>
                    <p className="text-caption text-secondary-foreground leading-relaxed">
                      Lamber 极简美学强调无边框设计与表面色阶的平滑移转。
                    </p>
                  </div>

                  {/* Form input and buttons */}
                  <div className="space-y-3">
                    <div className="space-y-1.5">
                      <Label htmlFor="preview-input" className="text-label">预算总金额 (含税)</Label>
                      <Input 
                        id="preview-input" 
                        type="text" 
                        value="1,245,600.00" 
                        readOnly 
                        className="font-mono text-sm ring-2 ring-primary/20 border-primary"
                      />
                    </div>

                    <div className="flex gap-2">
                      <Button variant="default" className="flex-1 text-xs">主按钮</Button>
                      <Button variant="outline" className="flex-1 text-xs">次按钮</Button>
                    </div>
                  </div>

                  {/* Financial indicators with Tabular Nums */}
                  <div className="pt-2 border-t border-border/40 flex items-center justify-between">
                    <div>
                      <div className="text-caption text-secondary-foreground font-semibold">测算毛利率</div>
                      <div className="text-metric font-extrabold text-foreground numeric-value">23.85%</div>
                    </div>
                    <div>
                      <div className="text-caption text-secondary-foreground font-semibold">投资净现值 (NPV)</div>
                      <div className="text-metric font-extrabold text-foreground numeric-value">¥128,456.00</div>
                    </div>
                  </div>

                  {/* Success / Warning notifications */}
                  <div className="space-y-2">
                    <div className="flex items-center gap-2 p-2.5 rounded-lg bg-success-soft/60 border border-success/10 text-[11px] font-medium text-success-foreground">
                      <AppIcon name="success" size={14} className="text-success shrink-0" />
                      财务数据校验已通过，支持一键导出 Excel 效益表。
                    </div>
                    <div className="flex items-center gap-2 p-2.5 rounded-lg bg-warning-soft/60 border border-warning/10 text-[11px] font-medium text-warning-foreground">
                      <AppIcon name="warning" size={14} className="text-warning shrink-0" />
                      当前 1% 自有产品保底规则提示：自身 CT 收入未达标。
                    </div>
                  </div>

                  {/* Simulated Chat Copilot */}
                  <div className="space-y-3 pt-4 border-t border-border/40">
                    <div className="text-caption text-secondary-foreground font-semibold flex items-center gap-1.5">
                      <AppIcon name="aiMessage" size={14} className="text-primary" />
                      AI 助手对话预览
                    </div>
                    <div className="space-y-2.5 text-xs">
                      {/* User Bubble */}
                      <div className="flex justify-end">
                        <div className="bg-primary text-primary-foreground max-w-[85%] rounded-2xl rounded-tr-none px-3.5 py-2 shadow-sm font-medium leading-relaxed">
                          如何优化该项目的 CT 专线带宽利润？
                        </div>
                      </div>
                      {/* Assistant Bubble */}
                      <div className="flex justify-start">
                        <div className="bg-muted text-foreground max-w-[90%] rounded-2xl rounded-tl-none px-3.5 py-2.5 border border-border/40 leading-relaxed space-y-1">
                          <p className="font-semibold text-primary">Lamber 智能顾问：</p>
                          <p className="text-secondary-foreground">
                            建议通过<strong>锁定收入总额结构反算</strong>，在保持投资额不变的前提下，适当下调 CT 带宽成本占比。
                          </p>
                        </div>
                      </div>
                    </div>
                  </div>

                </CardContent>
              </Card>
            </div>
          </div>

        </div>
      </div>
    </div>
  );
}
