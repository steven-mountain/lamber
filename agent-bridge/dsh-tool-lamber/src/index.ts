/**
 * dsh-tool-lamber — a deepseek-harness plugin that exposes lamber's Rust
 * business capabilities to the agent as tools.
 *
 * Every tool is a thin client over the loopback bridge server the lamber Tauri
 * backend hosts (`LAMBER_BRIDGE_URL`); no business math lives in this package.
 *
 * The plugin also owns the approval *guard*: which of its tools need a human.
 * Tools and their gating policy ship together on purpose — a tool whose risk is
 * declared in another package drifts from the tool itself. The decision half
 * lives in lamber, which answers ACP's `session/requestPermission` directly;
 * see `dsh-tool-lamber/src/approval.ts` for why the split falls there.
 */
import type { Context } from '@deepseek-ai/cordis';
import { applyApproval } from './approval.js';
import { runBenefitCalculation } from './runBenefitCalculation.js';
import { writeTestMarker } from './writeTestMarker.js';

export const name = 'dsh-tool-lamber';

export const inject = ['tools'] as const;

/**
 * Register lamber's tools and its approval guard on the harness runtime.
 *
 * @param ctx - the plugin context, with `tools` injected.
 */
export function apply(ctx: Context): void {
  ctx.tools.register(runBenefitCalculation);
  ctx.tools.register(writeTestMarker);
  applyApproval(ctx);
}

export { runBenefitCalculation, CALCULATE_ROUTE } from './runBenefitCalculation.js';
export { writeTestMarker, WRITE_TEST_MARKER } from './writeTestMarker.js';
export { applyApproval, isGatedTool } from './approval.js';
export {
  BRIDGE_URL_ENV,
  BRIDGE_TOKEN_ENV,
  BRIDGE_TOKEN_HEADER_ENV,
  LamberBridgeError,
  postBridge,
} from './bridge.js';
