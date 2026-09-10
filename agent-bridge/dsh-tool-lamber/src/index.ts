import { calculateSelectionFee, reverseCalculateSelectionFee } from './selectionFee.js';
import { simulateBenefitCalculation } from './simulateBenefitCalculation.js';
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
import { handshakeBridge, reportStartup, MISMATCH_MESSAGE } from './contract.js';
import type { Context } from '@deepseek-ai/cordis';
import { readBenefitInputs } from './readBenefitInputs.js';
import { readTemplateFields } from './readTemplateFields.js';
import { fillTemplateFields } from './fillTemplateFields.js';
import { queryProjects } from './queryProjects.js';
import { applyProjectScope } from './projectScope.js';
import { applyStreaming } from './stream.js';
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
export async function apply(ctx: Context): Promise<void> {
  try {
    await handshakeBridge();
  } catch (error) {
    reportStartup(error instanceof Error && error.message === MISMATCH_MESSAGE ? 'mismatch' : 'unreachable');
    throw error;
  }
  ctx.tools.register(runBenefitCalculation);
  ctx.tools.register(readBenefitInputs);
  ctx.tools.register(simulateBenefitCalculation);
  ctx.tools.register(calculateSelectionFee);
  ctx.tools.register(reverseCalculateSelectionFee);
  ctx.tools.register(queryProjects);
  ctx.tools.register(fillTemplateFields);
  ctx.tools.register(readTemplateFields);
  ctx.tools.register(writeTestMarker);
  applyProjectScope(ctx);
  applyApproval(ctx);
  applyStreaming(ctx);
  reportStartup('ready');
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

export { queryProjects } from './queryProjects.js';

export { fillTemplateFields, FILL_TEMPLATE_FIELDS } from './fillTemplateFields.js';

export { readTemplateFields } from './readTemplateFields.js';

export { readBenefitInputs } from './readBenefitInputs.js';

export { simulateBenefitCalculation } from './simulateBenefitCalculation.js';

export { calculateSelectionFee, reverseCalculateSelectionFee } from './selectionFee.js';
