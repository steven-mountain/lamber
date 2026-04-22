# Design Specification: 效益测算工具

This document outlines the design system extracted from the "效益测算工具" Stitch project.

## 1. Visual Identity & Creative Direction
**Creative North Star: "The Architectural Ledger"**
The design focuses on transforming complex financial calculations into a clear, authoritative, and trustworthy experience. It avoids rigid outlines in favor of tonal shifts and sophisticated surface layering.

## 2. Color Palette
The palette is rooted in deep, trustworthy blues and clinical greys.

| Category | Token | Hex Color | Usage |
| :--- | :--- | :--- | :--- |
| **Brand** | `primary` | `#285ab9` | Core actions and branding |
| | `primary_container` | `#d9e2ff` | High-impact moments |
| **Surfaces** | `surface` | `#f9f9ff` | Main background |
| | `surface_container` | `#e8edff` | Nested data layers |
| | `surface_container_highest`| `#d7e2ff` | Critical user inputs |
| **Typography** | `on_surface` | `#00316b` | Main text color |
| | `on_surface_variant` | `#005cbf` | Secondary metadata |
| **Feedback**| `error` | `#9f403d` | Critical warnings |
| | `outline` | `#2d78e4` | subtle boundaries |

### The "No-Line" Rule
Traditional 1px solid borders are replaced by:
- **Background Shifts**: Using different surface tokens to define sections.
- **Surface Nesting**: Creating depth by placing lighter surfaces inside darker ones.

## 3. Typography
**Primary Font Family**: `Inter`

- **Authority**: Inter is used for its legibility and perfect alignment of numerical data.
- **Data Display**: All numerical data uses `font-variant-numeric: tabular-nums` to ensure columns of figures align accurately.
- **Hierarchy**:
    - **Headlines**: Anchor the page without consuming excessive space.
    - **Labels**: High tracking and small scale for metadata hierarchy.

## 4. UI Components & Layout
- **Roundness**: `ROUND_FOUR` (Subtle 4px-8px corners) for a professional financial aesthetic.
- **Buttons**: Gradient fill from `primary` to `primary_dim`. Tightly clipped corners.
- **Glassmorphism**: Floating elements (modals/dropdowns) use `backdrop-filter: blur(12px)`.
- **Metrics**: End-of-journey results use `primary_container` with `on_primary_container` text.
