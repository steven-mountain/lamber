import type { Context } from '@deepseek-ai/cordis';
import type { ToolExecutionInput } from '@deepseek-ai/dsh-tools';
import { postBridge } from './bridge.js';

export const AUTHORIZE_ROUTE = '/lamber-bridge/authorize';
/** Identity is injected by the agent loop; never use a model argument. */
export function trustedSessionId(exec: Pick<ToolExecutionInput, 'agent'>): string {
  const id = exec.agent?.session.id;
  if (!id?.trim()) throw new Error('缺少可信 AI 会话身份，拒绝工具调用');
  return id;
}
export async function authorizeTool(exec: ToolExecutionInput): Promise<void> {
  const args = exec.arguments as { projectId?: unknown } | null;
  await postBridge(AUTHORIZE_ROUTE, {
    sessionId: trustedSessionId(exec), tool: exec.name,
    projectId: typeof args?.projectId === 'string' ? args.projectId : null,
  }, exec.signal);
}
/** Hard permission precedes the unchanged human approval guard. Unknown tools deny. */
export function applyProjectScope(ctx: Context): void {
  // Monotonic guard cannot be bypassed by an earlier waterfall listener.
  ctx.tools.guard(exec => {
    try { trustedSessionId(exec); } catch (error) { return String(error); }
    if (!['run_benefit_calculation', 'write_test_marker', 'query_projects', 'fill_template_fields', 'read_template_fields', 'read_benefit_inputs', 'simulate_benefit_calculation', 'calculate_selection_fee', 'reverse_calculate_selection_fee'].includes(exec.name)) return '该工具尚未定义项目权限，已拒绝调用';
  });
  ctx.on('tools/pre-execute', async (exec, next) => {
    try { await authorizeTool(exec); }
    catch (error) { return { kind: 'deny', reason: String(error) }; }
    return next();
  });
}
