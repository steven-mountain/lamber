# Design Specification: 效益测算工具

This document outlines the design system extracted from the "效益测算工具" Stitch project.

## 1. Visual Identity & Creative Direction
**Creative North Star: "The Architectural Ledger"**
The design focuses on transforming complex financial calculations into a clear, authoritative, and trustworthy experience. It avoids rigid outlines in favor of tonal shifts and sophisticated surface layering.

---

## 2. Lamber Global Visual Specification v1

This specification outlines the semantic tokens, typography scales, and visual layout parameters implemented in Lamber to guarantee visual consistency.

### 2.1 Color Palette & Semantic Roles
The palette is rooted in clinical greys, pale blues, and soft feedback color tokens. Hardcoded HSL values are stored in CSS variables and mapped into Tailwind configurations.

| Category | Token | CSS Variable Mapping | Hex / HSL Target | Usage |
| :--- | :--- | :--- | :--- | :--- |
| **Primary** | `primary` | `hsl(var(--primary))` | `#2563eb` / `221 83% 53%` | Primary actions and branding |
| | `primary-soft` | `hsl(var(--primary-soft))` | `221 83% 95%` | User chat bubbles, highlighted backgrounds |
| **Surfaces** | `background` | `hsl(var(--background))` | `#f8fafc` / `210 40% 98%` | Main viewport background |
| | `card` | `hsl(var(--card))` | `0 0% 100%` | Layout cards and panels |
| | `muted` | `hsl(var(--muted))` | `#f1f5f9` / `210 40% 96%` | Muted containers, inactive panels |
| **Borders** | `border` | `hsl(var(--border))` | `#e2e8f0` / `214 32% 91%` | Subtle boundaries |
| | `input` | `hsl(var(--input))` | `#cbd5e1` / `213 27% 84%` | Interactive field borders |
| **Feedback**| `success` | `hsl(var(--success))` | `142 70% 45%` | Positive statuses, healthy metrics |
| | `success-soft`| `hsl(var(--success-soft))` | `142 70% 96%` | Success badges, confirmation banners |
| | `warning` | `hsl(var(--warning))` | `38 92% 50%` | Mild alerts, warnings |
| | `warning-soft`| `hsl(var(--warning-soft))` | `38 92% 96%` | Warning cards, alert backgrounds |
| | `destructive` | `hsl(var(--destructive))` | `#9f403d` / `2 45% 43%` | Critical errors, destructive actions |
| | `destructive-soft`| `hsl(var(--destructive-soft))`| `2 45% 95%` | Error badges, alert panels |

### 2.2 The "No-Line" Rule
Traditional 1px solid dark borders are replaced by:
- **Background Shifts**: Using different surface tokens (e.g., nesting `bg-card` inside `bg-background` or `bg-muted/30`) to define sections.
- **Tonal Boundaries**: When outlines are necessary, always utilize `border-border` or `border-border/30` rather than arbitrary slate/zinc colors.
- **Shadow Depth**: Using subtle soft shadows (`shadow-sm` or `shadow-md`) to elevate active elements.

### 2.3 Typography & Sizing Scale
**Primary Font Stack**: `Inter, "Microsoft YaHei", "PingFang SC", "Noto Sans SC", system-ui, -apple-system, BlinkMacSystemFont, sans-serif`
Legibility, structure, and perfect alignment of numerical metrics are enforced by typographic scales linked to the base variable `--font-scale` (defaults to `1`).

| Typographic Role | CSS Size Variable | Computed Default Value | Weight | Line Height | Usage |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Display** | `--text-display` | `calc(28px * var(--font-scale))` | `600` | `36px` | Main headers |
| **Page Title** | `--text-page-title` | `calc(22px * var(--font-scale))` | `600` | `30px` | Primary view titles |
| **Section Title** | `--text-section-title` | `calc(16px * var(--font-scale))` | `600` | `24px` | Card / Sidebar headers |
| **Body** | `--text-body` | `calc(14px * var(--font-scale))` | `400` | `22px` | Paragraphs and text blocks |
| **Body Strong** | `--text-body-strong` | `calc(14px * var(--font-scale))` | `600` | `22px` | Bold body items |
| **Label** | `--text-label` | `calc(13px * var(--font-scale))` | `500` | `20px` | Form and field labels |
| **Caption** | `--text-caption` | `calc(12px * var(--font-scale))` | `400` | `18px` | Metadata and helper tips |
| **Metric** | `--text-metric` | `calc(22px * var(--font-scale))` | `600` | `30px` | Dashboard indicators |

#### Numerical Presentation
All financial columns, cashflow tables, NPV graphs, percentage badges, and calculator inputs must incorporate the `.numeric-value` utility to enforce:
```css
font-variant-numeric: tabular-nums;
```
This ensures numbers align vertically, facilitating visual audits of financial ledgers.

### 2.4 Radii & Corner Standards
- **Standard Corners (`ROUND_FOUR`)**: Base border-radius uses `var(--radius)` (mapped to Tailwind `rounded-lg` / `rounded-xl`).
- **Inner elements** use computed radii `calc(var(--radius) - 2px)` (`rounded-md`) or `calc(var(--radius) - 4px)` (`rounded-sm`) to avoid corner overlap.
- **Glassmorphism**: Floating panels, modals, and backdrop drop-shadow menus utilize a backdrop blur filter (`backdrop-blur-sm` or `backdrop-filter: blur(12px)`) combined with a semi-transparent surface background (`bg-background/80` or `bg-card/95`).

### 2.5 Theme Presets
Lamber supports five light theme presets and a unified dark theme framework:

1. **Lamber 默认 (`lamber`)**: Slate light grey background (`#f8fafc`) + Royal blue accent (`#2563eb`). Traditional corporate engineering style.
2. **石墨灰 (`graphite`)**: Pure neutral grey background (`#f1f3f5`) + Deep charcoal accent (`#343a40`). Low-contrast professional view.
3. **海军蓝 (`navy`)**: Deep blue-tinted grey background (`#f4f6f9`) + Navy blue accent (`#1e3a8a`). Government / enterprise project layout.
4. **森林绿 (`forest`)**: Sage-green-tinted background (`#f4f9f4`) + Forest green accent (`#15803d`). Distinguished branding.
5. **暖石色 (`warmStone`)**: Warm brown-grey background (`#f7f5f2`) + Warm clay accent (`#5c5346`). High-readability low-fatigue layout.

#### Dark Mode Base
To maintain design boundaries and contrast, all presets in dark mode utilize a single shared background base (`#0f172a` / `#1e293b`) with slight primary/accent tone adjustments matched to their theme IDs (e.g., `#3b82f6` for Lamber vs `#15803d`-like accents for Forest).

### 2.6 Font Size Scaling Presets
To facilitate screen compatibility and low-vision accessibility, typographic styles scale dynamically through `--font-scale`:

- **紧凑 (`compact`)**: `0.93` scaling factor. Ideal for small displays or dense financial logs.
- **标准 (`standard`)**: `1.00` scaling factor. Default.
- **舒适 (`comfortable`)**: `1.08` scaling factor. For comfortable daily reading.
- **大字号 (`large`)**: `1.16` scaling factor. Optimized for high-DPI viewports.

Typographic line-heights are dynamically scaled (`lh * var(--font-scale)`) proportionally to prevent character overlaps.

### 2.7 Interface Density Presets
Spacing values adapt dynamically based on three interface densities via CSS variables:

- **紧凑 (`compact`)**: Card padding `1rem` (p-4), form control height `2rem` (h-8), table cell vertical padding `0.5rem` (8px). Designed for maximum data-density.
- **标准 (`standard`)**: Card padding `1.5rem` (p-6), form control height `2.25rem` (h-9), table cell vertical padding `0.75rem` (12px). Default.
- **宽松 (`comfortable`)**: Card padding `2rem` (p-8), form control height `2.5rem` (h-10), table cell vertical padding `1rem` (16px). Maximized breathing room.

### 2.8 Runtime DOM Application
The active settings are applied via `document.documentElement` data attributes and CSS variables:
- Attributes: `data-theme`, `data-color-mode`, `data-density`.
- Toggled Classes: `.dark` for Tailwind theme compilers.
- Variables: Writes variables directly into `style` properties, which guarantees immediate styling changes without app reloading.

### 2.9 Advanced Customization & Contrast Validation (Phase 3)

Lamber Phase 3 introduces safe user-customizable accent colors, WCAG-compliant contrast checking, and a dedicated high-contrast preference mode:

#### Custom Accent Color Selection & Boundaries
- Users can choose from a set of pre-calculated high-contrast recommended palettes (Business Blue, Navy, Teal, Forest, Amber, Graphite) or input a custom hex value through a native HTML5 color picker.
- **Scope Limit**: The custom color only overrides the theme's accent elements (primary, primary-foreground, primary-soft, ring, accent, accent-foreground). Other layout tokens (background, foreground, card, borders, popover) and state colors (success, warning, destructive) are not editable to prevent theme fragmentation.

#### Contrast Checking & Automatic HSL Derivation
- All custom accent selections are validated against WCAG AA standards.
- In **Light Mode**, the primary accent is checked against white background (`#FFFFFF`). If the contrast ratio is below the minimum threshold (4.5:1 for standard, 7.0:1 for high contrast), the system automatically darkens the HSL lightness (`L`) until it meets the target.
- In **Dark Mode**, the primary accent is checked against dark background (`#0F172A`). If the contrast ratio is below the minimum threshold, the system automatically lightens the HSL lightness (`L`) until it meets the target.
- `primary-foreground` is dynamically chosen between light slate (`210 40% 98%`) and dark slate (`222 47% 11%`) based on which text color provides higher contrast against the adjusted primary background.
- Warning banners in the settings panel alert the user when their chosen custom color has been auto-adjusted to comply with accessibility rules.

#### High Contrast Preference Mode
- When `contrastPreference === "high"`, the application layers high contrast overrides onto standard presets:
  - **Light High Contrast**: Sets pure white backgrounds (`#FFFFFF`), near-black text (`#0F172A` / `#000000`), darker secondary and muted foregrounds, and thick high-visibility borders (`#4B5563`).
  - **Dark High Contrast**: Sets pure black backgrounds (`#000000`), pure white text (`#FFFFFF`), bright visible borders, and highly distinct success/warning/destructive badges.
  - Accent contrast ratio threshold is raised to a strict 7.0:1.

#### Preset Dark Mode Refinements
- The five presets (`lamber`, `graphite`, `navy`, `forest`, `warmStone`) feature distinct, non-overlapping dark mode surface values (e.g. graphite uses pure neutral dark greys, navy uses deep dark blue greys, forest uses dark green greys, and warmStone uses dark brown-tinted greys). This ensures proper layout hierarchy and visual identity is preserved even in dark modes.

### 2.10 Functional Bounds
The appearance system represents user preferences. It is strictly forbidden to:
- Write preferences to workspace project SQLite tables.
- Support free-form font file uploads.
- Allow changes to financial calculations, cashflows, NPV, or document content generation based on appearance parameters.
- Direct theme imports/exports are currently not implemented.


