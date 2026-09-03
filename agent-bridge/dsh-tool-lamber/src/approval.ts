/**
 * lamber's human-approval channel for dsh tools.
 *
 * Two halves, both required:
 *
 * * **The guard** (`tools/pre-execute`) decides *which* calls need a human. dsh
 *   has no declarative "risk level" on `defineTool`; a gated tool is one a
 *   pre-execute listener answers `{kind: 'ask'}` for. The waterfall's terminal
 *   default is `{kind: 'allow'}`, so anything this guard does not name — such as
 *   the read-only `run_benefit_calculation` — proceeds untouched.
 * * **The answerer** (`approval/request`) decides *what the human said*. It
 *   forwards the question to lamber over the same authenticated bridge the tools
 *   use, and blocks until lamber returns a decision.
 *
 * Failure is closed at every step. dsh maps `rejected` / `cancelled` /
 * `unavailable` to a denial with distinct reasons; a bridge error, a malformed
 * reply, or a timeout all resolve `'rejected'` here rather than hanging or
 * defaulting to a grant.
 */
import type { Context } from '@deepseek-ai/cordis';
import type { ApprovalOutcome } from '@deepseek-ai/dsh-user-approval/types';
import { postBridge } from './bridge.js';
import { rememberCall, takeCall } from './pendingCalls.js';
import { WRITE_TEST_MARKER } from './writeTestMarker.js';

/** Bridge route serving approval questions. */
export const APPROVAL_ROUTE = '/lamber-bridge/approval';

/**
 * How long the answerer waits for lamber before failing closed.
 *
 * Deliberately longer than lamber's own wait so the normal path is lamber
 * returning an explicit `rejected` on its timeout; this bound only catches a
 * bridge that stops answering entirely.
 */
export const ANSWERER_TIMEOUT_MS = 180_000;

/** Tools that require a human decision, with the reason shown to the user. */
const GATED_TOOLS = new Map<string, string>([
  [WRITE_TEST_MARKER, '该工具会写入文件（测试标记文件，位于系统临时目录），需要你确认后才执行。'],
]);

/**
 * Whether a tool requires a human decision.
 *
 * The single source of truth the guard consults, exported so the gating policy
 * can be asserted directly instead of inferred from an agent run.
 *
 * @param toolName - the tool name dsh is about to dispatch.
 * @returns `true` only for tools this plugin gates.
 */
export function isGatedTool(toolName: string): boolean {
  return GATED_TOOLS.has(toolName);
}

/** Decision contract of `POST /lamber-bridge/approval`; mirrors the Rust `ApprovalDecision`. */
interface ApprovalDecisionReply {
  /** `true` only for an explicit human grant. */
  approved: boolean;
  /** Optional human-readable explanation, surfaced in logs. */
  reason?: string;
}

/**
 * Register the approval guard and answerer on one plugin context.
 *
 * @param ctx - the plugin context; both registrations are disposed with it.
 */
export function applyApproval(ctx: Context): void {
  ctx.on('tools/pre-execute', async (exec, next) => {
    const reason = GATED_TOOLS.get(exec.name);
    if (reason === undefined) return next();
    // The approval question carries no arguments, so stash them for the answerer.
    if (exec.callId !== undefined) {
      rememberCall(exec.callId, exec.name, exec.arguments);
    }
    return { kind: 'ask', reason };
  });

  ctx.on('approval/request', async (request, next) => {
    if (!GATED_TOOLS.has(request.toolName)) return next();
    return askLamber(request.toolName, request.callId, request.reason, request.signal);
  });
}

/**
 * Put one question to lamber and normalize every failure to a denial.
 *
 * Exported so `scripts/check-approval.mjs` can exercise this exact path without
 * booting dsh or spending an LLM call.
 *
 * @param toolName - the tool awaiting a decision.
 * @param callId - the exact tool call, when dsh supplied one.
 * @param reason - the guard's human-readable explanation.
 * @param signal - dsh's cancellation signal for the pending question.
 * @returns the closed outcome; only an explicit grant returns `allowed-once`.
 */
export async function askLamber(
  toolName: string,
  callId: string | undefined,
  reason: string | undefined,
  signal: AbortSignal | undefined,
): Promise<ApprovalOutcome> {
  const recorded = takeCall(callId);
  const controller = new AbortController();
  const onAbort = () => controller.abort();
  signal?.addEventListener('abort', onAbort, { once: true });
  const timer = setTimeout(() => controller.abort(), ANSWERER_TIMEOUT_MS);

  try {
    const reply = await postBridge<ApprovalDecisionReply>(
      APPROVAL_ROUTE,
      {
        toolName,
        callId: callId ?? null,
        reason: reason ?? null,
        args: recorded?.args ?? null,
      },
      controller.signal,
    );
    return reply.approved === true ? 'allowed-once' : 'rejected';
  } catch (error) {
    // dsh withdrew the question: report the cancellation it is expecting.
    if (signal?.aborted === true) return 'cancelled';
    console.error(
      `[dsh-tool-lamber] 审批请求失败，按拒绝处理: ${
        error instanceof Error ? error.message : String(error)
      }`,
    );
    return 'rejected';
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener('abort', onAbort);
  }
}
