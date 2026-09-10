import type { ToolExecutionInput } from '@deepseek-ai/dsh-tools';
import { postBridge } from './bridge.js';
import { trustedSessionId } from './projectScope.js';
export async function reviewedArguments(exec: ToolExecutionInput): Promise<{ note?: string }> {
  if (!exec.callId) throw new Error('缺少可信工具调用身份，拒绝执行');
  const result = await postBridge<{args:{note?:string}}>('/lamber-bridge/reviewed-arguments', {
    sessionId: trustedSessionId(exec), callId: exec.callId, tool: exec.name, originalArgs: exec.arguments,
  }, exec.signal);
  return result.args;
}
