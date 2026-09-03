/**
 * In-process correlation between a gated tool call and its approval question.
 *
 * `ApprovalRequestEvent` deliberately carries no arguments — its doc comment
 * says `callId` "links to an already presented tool call, so arguments are not
 * duplicated here". lamber's approval dialog still needs to show the user what
 * the tool would actually do, so the pre-execute guard records the parsed
 * arguments here and the answerer reads them back by `callId`.
 *
 * The map is bounded and self-cleaning: the answerer removes the entry it
 * consumes, and a guard that never reaches an answerer (denied earlier in the
 * waterfall, or a cancelled turn) is evicted once the map exceeds its cap.
 */

/** Most recent gated calls kept for correlation; older entries are dropped first. */
const MAX_PENDING = 64;

/** One gated call awaiting (or having received) an approval decision. */
export interface PendingCall {
  readonly toolName: string;
  readonly args: unknown;
  readonly recordedAt: number;
}

const pending = new Map<string, PendingCall>();

/**
 * Record a gated call's arguments so the answerer can present them.
 *
 * @param callId - the tool call id the approval question will carry.
 * @param toolName - the gated tool's name.
 * @param args - the validated arguments the tool would run with.
 */
export function rememberCall(callId: string, toolName: string, args: unknown): void {
  pending.set(callId, { toolName, args, recordedAt: Date.now() });
  while (pending.size > MAX_PENDING) {
    const oldest = pending.keys().next();
    if (oldest.done) break;
    pending.delete(oldest.value);
  }
}

/**
 * Consume a recorded call.
 *
 * @param callId - the id from the approval request.
 * @returns the recorded call, or `undefined` when the guard did not record one.
 */
export function takeCall(callId: string | undefined): PendingCall | undefined {
  if (callId === undefined) return undefined;
  const found = pending.get(callId);
  if (found !== undefined) pending.delete(callId);
  return found;
}

/** Drop every recorded call. Test-only; production entries expire by cap. */
export function resetPendingCalls(): void {
  pending.clear();
}
