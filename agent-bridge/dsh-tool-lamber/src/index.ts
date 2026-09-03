/**
 * dsh-tool-lamber — a deepseek-harness plugin that exposes lamber's Rust
 * business capabilities to the agent as tools.
 *
 * Every tool is a thin client over the loopback bridge server the lamber Tauri
 * backend hosts (`LAMBER_BRIDGE_URL`); no business math lives in this package.
 *
 * The plugin also owns lamber's approval channel. Tools and their approval
 * policy ship together on purpose: the guard correlates a gated call's
 * arguments to its approval question through an in-process map, and splitting
 * the halves across two npm packages would risk two module instances holding
 * two separate maps.
 */
import type { Context } from '@deepseek-ai/cordis';
import { applyApproval } from './approval.js';
import { runBenefitCalculation } from './runBenefitCalculation.js';
import { writeTestMarker } from './writeTestMarker.js';

export const name = 'dsh-tool-lamber';

export const inject = ['tools'] as const;

/**
 * Register lamber's tools and its approval channel on the harness runtime.
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
export {
  applyApproval,
  askLamber,
  isGatedTool,
  APPROVAL_ROUTE,
  ANSWERER_TIMEOUT_MS,
} from './approval.js';
export {
  BRIDGE_URL_ENV,
  BRIDGE_TOKEN_ENV,
  BRIDGE_TOKEN_HEADER_ENV,
  LamberBridgeError,
  postBridge,
} from './bridge.js';
