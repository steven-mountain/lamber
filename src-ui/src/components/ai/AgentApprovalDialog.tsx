import { useCallback, useEffect, useRef, useState } from "react";
import { listen } from "@tauri-apps/api/event";
import { invoke } from "@tauri-apps/api/core";

/** Backend event carrying one pending approval question. */
export const AI_APPROVAL_REQUEST_EVENT = "ai://approval-request";

/** Mirrors `agent_bridge::approval::ApprovalPrompt`. */
interface ApprovalPrompt {
  requestId: string;
  toolName: string;
  callId: string | null;
  reason: string | null;
  args: unknown;
  timeoutSeconds: number;
}

/**
 * Confirmation dialog for agent tool calls that require a human decision.
 *
 * Mounted at the app root rather than inside the AI panel on purpose: the
 * backend parks a dsh tool call waiting for this answer, so a listener that
 * disappeared when the panel closed would turn every approval into a timeout.
 *
 * Deliberately plain — this exists to make the approval channel usable and
 * demonstrable, not to be the final UI.
 */
export default function AgentApprovalDialog() {
  const [queue, setQueue] = useState<ApprovalPrompt[]>([]);
  const [busy, setBusy] = useState(false);
  const [remaining, setRemaining] = useState<number | null>(null);
  const current = queue[0] ?? null;
  // Read inside the countdown effect without making it a dependency.
  const currentRef = useRef<ApprovalPrompt | null>(null);
  currentRef.current = current;

  useEffect(() => {
    let unlisten: (() => void) | undefined;
    let disposed = false;
    listen<ApprovalPrompt>(AI_APPROVAL_REQUEST_EVENT, event => {
      setQueue(prev =>
        prev.some(item => item.requestId === event.payload.requestId)
          ? prev
          : [...prev, event.payload],
      );
    })
      .then(handler => {
        if (disposed) handler();
        else unlisten = handler;
      })
      .catch(error => console.warn("Failed to listen for approval requests:", error));
    return () => {
      disposed = true;
      unlisten?.();
    };
  }, []);

  // Mirror the backend's own deadline so the dialog cannot look actionable
  // after the call has already been auto-rejected.
  useEffect(() => {
    if (!current) {
      setRemaining(null);
      return;
    }
    setRemaining(current.timeoutSeconds);
    const requestId = current.requestId;
    const timer = window.setInterval(() => {
      setRemaining(prev => {
        if (prev === null) return null;
        if (prev > 1) return prev - 1;
        if (currentRef.current?.requestId === requestId) {
          setQueue(items => items.filter(item => item.requestId !== requestId));
        }
        return null;
      });
    }, 1000);
    return () => window.clearInterval(timer);
  }, [current]);

  const respond = useCallback(
    async (approved: boolean) => {
      if (!current || busy) return;
      setBusy(true);
      try {
        await invoke("ai_resolve_approval", { requestId: current.requestId, approved });
      } catch (error) {
        // A timed-out request is already denied on the backend; drop it either way.
        console.warn("Failed to resolve approval request:", error);
      } finally {
        setQueue(items => items.filter(item => item.requestId !== current.requestId));
        setBusy(false);
      }
    },
    [busy, current],
  );

  if (!current) return null;

  const argsText =
    current.args === null || current.args === undefined
      ? "（无参数）"
      : JSON.stringify(current.args, null, 2);

  return (
    <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/40 p-4">
      <div className="w-full max-w-md rounded-2xl bg-background p-5 shadow-xl">
        <h2 className="text-base font-semibold text-foreground">AI 请求执行操作</h2>
        <p className="mt-2 text-sm text-muted-foreground">
          {current.reason ?? "该操作需要你确认后才会执行。"}
        </p>

        <div className="mt-4 rounded-xl bg-muted p-3">
          <div className="text-xs text-muted-foreground">工具</div>
          <div className="font-mono text-sm text-foreground">{current.toolName}</div>
          <div className="mt-3 text-xs text-muted-foreground">参数</div>
          <pre className="mt-1 max-h-40 overflow-auto whitespace-pre-wrap break-all font-mono text-xs text-foreground">
            {argsText}
          </pre>
        </div>

        <div className="mt-4 flex items-center justify-between gap-3">
          <span className="text-xs tabular-nums text-muted-foreground">
            {remaining !== null ? `${remaining} 秒后自动拒绝` : "已超时"}
          </span>
          <div className="flex gap-2">
            <button
              type="button"
              disabled={busy}
              onClick={() => respond(false)}
              className="rounded-lg bg-muted px-4 py-2 text-sm text-foreground disabled:opacity-50"
            >
              拒绝
            </button>
            <button
              type="button"
              disabled={busy}
              onClick={() => respond(true)}
              className="rounded-lg bg-primary px-4 py-2 text-sm text-primary-foreground disabled:opacity-50"
            >
              确认执行
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
