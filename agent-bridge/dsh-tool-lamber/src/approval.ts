/**
 * lamber's approval *guard* for dsh tools.
 *
 * dsh has no declarative "risk level" on `defineTool`; a gated tool is one a
 * `tools/pre-execute` listener answers `{kind: 'ask'}` for. The waterfall's
 * terminal default is `{kind: 'allow'}`, so anything this guard does not name —
 * such as the read-only `run_benefit_calculation` — proceeds untouched.
 *
 * **This file used to have a second half.** An `approval/request` listener
 * posted the question to lamber over the loopback bridge and blocked on the
 * answer. Under `--profile acp` that listener is unreachable: `dsh-acp`
 * registers its own `approval/request` handler ahead of this plugin's and
 * forwards the question to the ACP client as `session/requestPermission`
 * (`dsh-acp/lib/index.js:1118-1140`), so the waterfall never reaches here. The
 * answerer, its `pendingCalls` correlation map, and lamber's
 * `/lamber-bridge/approval` route were removed rather than left as a second,
 * dead path to the same decision. lamber now answers over ACP; see
 * `src-tauri/src/agent_bridge/dsh_session.rs`.
 *
 * What survives is exactly the half ACP does not replace: deciding *which*
 * calls need a human. dsh-acp only forwards a question once something has asked
 * one, and for lamber's tools this guard is what asks.
 */
import type { Context } from '@deepseek-ai/cordis';
import { WRITE_TEST_MARKER } from './writeTestMarker.js';

/**
 * Tools that require a human decision, with the reason recorded alongside.
 *
 * The reason no longer reaches the user: `dsh-acp` builds its
 * `session/requestPermission` params from the tool call id alone and drops it.
 * lamber keeps the text the dialog shows in its own mirror table
 * (`src-tauri/src/agent_bridge/approval.rs`), and a Rust test fails if the two
 * tables stop naming the same tools. It stays here because the guard's own
 * contract takes a reason, and because this is where a reader looks to learn
 * why a tool is gated.
 */
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

/**
 * Register the approval guard on one plugin context.
 *
 * @param ctx - the plugin context; the registration is disposed with it.
 */
export function applyApproval(ctx: Context): void {
  ctx.on('tools/pre-execute', async (exec, next) => {
    const reason = GATED_TOOLS.get(exec.name);
    if (reason === undefined) return next();
    return { kind: 'ask', reason };
  });
}
