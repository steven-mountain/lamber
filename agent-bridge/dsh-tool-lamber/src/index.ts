/**
 * dsh-tool-lamber — a deepseek-harness plugin that exposes lamber's Rust
 * business capabilities to the agent as tools.
 *
 * Every tool is a thin client over the loopback bridge server the lamber Tauri
 * backend hosts (`LAMBER_BRIDGE_URL`); no business math lives in this package.
 */
import type { Context } from '@deepseek-ai/cordis';
import { runBenefitCalculation } from './runBenefitCalculation.js';

export const name = 'dsh-tool-lamber';

export const inject = ['tools'] as const;

/**
 * Register lamber's tools on the harness tool runtime.
 *
 * @param ctx - the plugin context, with `tools` injected.
 */
export function apply(ctx: Context): void {
  ctx.tools.register(runBenefitCalculation);
}

export { runBenefitCalculation, CALCULATE_ROUTE } from './runBenefitCalculation.js';
export {
  BRIDGE_URL_ENV,
  BRIDGE_TOKEN_ENV,
  BRIDGE_TOKEN_HEADER_ENV,
  LamberBridgeError,
  postBridge,
} from './bridge.js';
