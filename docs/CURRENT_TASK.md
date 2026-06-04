# 采购甄选费测算固定锚点优化

- **Status:** Done
- **Objective:** 在 ICT 投入侧“采购甄选费测算”中增加供应商报价与甄选最高限价的互斥固定锚点，使代理服务费上浮变化时按用户固定的金额方向联动另一个金额。

## Completed

1. [x] **固定锚点状态**
   - 新增 `selectionFeeAnchor`，默认固定“供应商报价”，保持原有“报价 + 上浮 -> 限价”行为。
   - 点击“固定供应商报价”或“固定甄选最高限价”时互斥切换。
   - 直接编辑供应商报价或最高限价时，对应金额自动成为当前固定锚点。
2. [x] **联动方向收敛**
   - 固定供应商报价时，`markup` 变化走 `calculate_selection_fee` 正向计算并更新最高限价。
   - 固定最高限价时，`markup` 变化走 `reverse_calculate_selection_fee` 反向计算并更新供应商报价。
   - 增加请求序号保护，避免连续输入时旧异步计算结果覆盖最新输入。
3. [x] **简单交互与可访问性**
   - 在“供应商报价”和“甄选最高限价”标题旁增加圆点按钮表示固定状态。
   - 为甄选测算输入补充 `aria-label`，便于键盘/辅助技术和自动化验证识别。

## Validation

- `npx tsc --noEmit` in `src-ui`: passed.
- `npm run build` in `src-ui`: passed.
- `cargo test` in `src-tauri`: passed (11/11; existing warnings only).
- Browser smoke test at `http://127.0.0.1:5173/`: fixed-dot default, mutual exclusion, and edit-to-anchor state passed. Plain Vite browser has no Tauri IPC, so backend invoke numeric calculation was not executed there.

## Remaining

- 在 Tauri 桌面运行态中手工输入一组报价、上浮和限价，确认正向/反向金额联动与“填入集成服务”工作流。

## Next

- 若后续需要更清晰的视觉提示，可在不增加说明文案的前提下微调固定圆点的颜色或 hover 状态。
