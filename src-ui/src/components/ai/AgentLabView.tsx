import { useCallback, useEffect, useRef, useState } from "react";
import { listen } from "@tauri-apps/api/event";
import { invoke } from "@tauri-apps/api/core";

/** Backend event carrying every ACP notification and turn outcome. */
const AI_SESSION_EVENT = "ai://session-event";

/** Method the backend labels one ACP `session/update` with. */
const UPDATE_METHOD = "session/update";

/** Method the backend labels the end of one turn with. */
const TURN_ENDED_METHOD = "session/turn-ended";

/**
 * Label one backend event for the log.
 *
 * ACP tags every notification kind inside the payload rather than giving each
 * its own method, so the interesting name lives in `params.update.sessionUpdate`
 * and every update would otherwise read as the same line.
 */
function describeEvent(method: string, params: unknown): string {
  if (method === TURN_ENDED_METHOD) {
    // Fire-and-forget prompts return before the turn runs, so this event is the
    // only thing that says the agent has stopped working.
    const outcome = params as { stopReason?: string; error?: string };
    return outcome.error ? "本轮结束（出错）" : `本轮结束 · ${outcome.stopReason ?? "?"}`;
  }
  if (method !== UPDATE_METHOD) return method;
  const kind = (params as { update?: { sessionUpdate?: string } })?.update?.sessionUpdate;
  return kind ? `${method} · ${kind}` : method;
}

/** One line in the event log. */
interface LogLine {
  id: number;
  label: string;
  detail: string;
}

interface ApprovalLogEntry {
  requestId: string;
  toolName: string;
  approved: boolean;
  decidedBy: string;
  decisionReason: string;
  decidedAt: string;
}

/**
 * Minimal harness for driving the dsh agent from the real application.
 *
 * It exists because nothing else in the UI calls `ai_send_prompt`: without a
 * trigger, the approval dialog can never appear in a running app and the
 * approval channel cannot be exercised by hand. Deliberately plain — this is a
 * lab bench, not the product surface. Reachable at `#/agent-lab`.
 */
export default function AgentLabView() {
  const [text, setText] = useState("请调用 write_test_marker 工具，note 参数填「真实点击联调」。");
  const [lines, setLines] = useState<LogLine[]>([]);
  const [sending, setSending] = useState(false);
  const [auditLog, setAuditLog] = useState<ApprovalLogEntry[]>([]);
  const nextId = useRef(0);

  const append = useCallback((label: string, detail: string) => {
    nextId.current += 1;
    const id = nextId.current;
    setLines(prev => [...prev.slice(-200), { id, label, detail }]);
  }, []);

  useEffect(() => {
    let unlisten: (() => void) | undefined;
    let disposed = false;
    listen<{ method: string; params: unknown }>(AI_SESSION_EVENT, event => {
      const { method, params } = event.payload;
      append(describeEvent(method, params), JSON.stringify(params));
    })
      .then(handler => {
        if (disposed) handler();
        else unlisten = handler;
      })
      .catch(error => append("listen-error", String(error)));
    return () => {
      disposed = true;
      unlisten?.();
    };
  }, [append]);

  const refreshAudit = useCallback(async () => {
    try {
      setAuditLog(await invoke<ApprovalLogEntry[]>("ai_list_approval_log", { limit: 20 }));
    } catch (error) {
      append("audit-error", String(error));
    }
  }, [append]);

  useEffect(() => {
    void refreshAudit();
  }, [refreshAudit]);

  // `#/agent-lab?autorun=1` fires one prompt on mount. Lets the approval dialog
  // be reached (and screenshotted) without a click, for verifying rendering and
  // the timeout path; the confirm/reject paths still need a real click.
  const autorun = window.location.hash.includes("autorun=1");
  const autorunFired = useRef(false);

  const send = useCallback(async () => {
    if (sending || !text.trim()) return;
    setSending(true);
    const sessionId = `lab-${Date.now()}`;
    append("prompt", `${sessionId} · ${text}`);
    try {
      // ACP names sessions itself, so this is the agent's id, not the one sent.
      const acpSession = await invoke<string>("ai_send_prompt", { sessionId, text });
      append("accepted", `ACP 会话 ${acpSession}`);
    } catch (error) {
      append("send-error", String(error));
    } finally {
      setSending(false);
    }
  }, [append, sending, text]);

  useEffect(() => {
    if (!autorun || autorunFired.current) return;
    autorunFired.current = true;
    const timer = window.setTimeout(() => void send(), 400);
    return () => window.clearTimeout(timer);
  }, [autorun, send]);

  return (
    <div className="flex h-screen flex-col gap-3 bg-background p-5 text-foreground">
      <h1 className="text-lg font-semibold">Agent 联调台（实验）</h1>
      <p className="text-xs text-muted-foreground">
        用于人工验证 dsh 工具调用与审批通道（ACP 协议，
        <code className="px-1">dsh --profile acp</code>）。需要设置 DEEPSEEK_API_KEY
        环境变量后启动应用。
      </p>

      <div className="flex gap-2">
        <input
          value={text}
          onChange={event => setText(event.target.value)}
          onKeyDown={event => {
            if (event.key === "Enter") void send();
          }}
          className="flex-1 rounded-lg bg-muted px-3 py-2 text-sm"
          placeholder="给 Agent 的指令"
        />
        <button
          type="button"
          data-testid="agent-lab-send"
          disabled={sending}
          onClick={() => void send()}
          className="rounded-lg bg-primary px-4 py-2 text-sm text-primary-foreground disabled:opacity-50"
        >
          {sending ? "发送中…" : "发送"}
        </button>
        <button
          type="button"
          onClick={() => void refreshAudit()}
          className="rounded-lg bg-muted px-4 py-2 text-sm"
        >
          刷新审批日志
        </button>
      </div>

      <div className="grid min-h-0 flex-1 grid-cols-2 gap-3">
        <div className="flex min-h-0 flex-col rounded-xl bg-muted p-3">
          <div className="mb-2 text-xs font-medium text-muted-foreground">会话事件</div>
          <div className="min-h-0 flex-1 overflow-auto font-mono text-[11px] leading-relaxed">
            {lines.map(line => (
              <div key={line.id} className="mb-1 break-all">
                <span className="text-primary">{line.label}</span>{" "}
                <span className="text-muted-foreground">{line.detail.slice(0, 400)}</span>
              </div>
            ))}
          </div>
        </div>

        <div className="flex min-h-0 flex-col rounded-xl bg-muted p-3">
          <div className="mb-2 text-xs font-medium text-muted-foreground">
            审批审计日志（持久化）
          </div>
          <div className="min-h-0 flex-1 overflow-auto font-mono text-[11px] leading-relaxed">
            {auditLog.length === 0 ? (
              <div className="text-muted-foreground">（暂无记录）</div>
            ) : (
              auditLog.map(entry => (
                <div key={entry.requestId} className="mb-1 break-all">
                  <span className={entry.approved ? "text-primary" : "text-muted-foreground"}>
                    {entry.approved ? "已批准" : "已拒绝"}
                  </span>{" "}
                  {entry.toolName} · {entry.decidedBy} · {entry.decisionReason} ·{" "}
                  <span className="tabular-nums">{entry.decidedAt}</span>
                </div>
              ))
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
