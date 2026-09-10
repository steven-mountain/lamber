import { readableArguments } from '../../ai/approvalReview';
import AiSessionProjectPicker from './AiSessionProjectPicker';
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
  argsJson: string;
  requestId: string;
  toolName: string;
  approved: boolean;
  decidedBy: string;
  decisionReason: string;
  decidedAt: string;
}

const REVIEW_SAMPLE = `本项目拟为客户现有办公园区提供统一网络接入与业务协同支撑，解决多个楼栋网络覆盖不均、业务系统访问不稳定、故障处理依赖人工排查等问题。建设范围覆盖办公区、公共服务区及必要的设备机房，具体点位以双方确认的现场勘察记录为准。

方案应兼顾现有系统平稳运行与后续扩容，按照分区实施、逐项验证的方式安排施工。实施前需确认接入条件、设备安装位置、供电与布线路由，避免影响客户日常办公；涉及业务切换的工作应提前沟通窗口，并保留回退方案。

交付内容包括现场勘察、设备安装调试、网络联通验证、使用培训和交付资料整理。验收应依据双方确认的需求清单逐项核对，重点检查覆盖范围、访问稳定性、故障告警与运维交接情况。项目完成后应明确日常维护联系人、问题反馈渠道与升级处理流程，并对尚未完成的事项形成书面清单。

本段为审批界面演练样例，不代表任何真实项目的已确认需求。请替换为待审核正文，核对后再批准。`;

function AuditVersions({ argsJson }: { argsJson: string }) {
  let value;
  try { value = JSON.parse(argsJson); } catch { return <p>审计正文无法解析</p>; }
  const versions = value?.auditVersion === 2
    ? [{title:'模型原文',args:value.modelArgs},{title:'用户修改值',args:value.userArgs ?? '未修改'},{title:'最终批准值',args:value.approvedArgs ?? '未批准'}]
    : [{title:'历史参数',args:value}];
  return <details className="mt-1 font-sans"><summary className="cursor-pointer">查看原文与修改记录</summary>{versions.map(version => <div key={version.title} className="mt-2 rounded-lg bg-card p-2"><strong>{version.title}</strong>{readableArguments(version.args).map((item,i) => <p key={i} className="whitespace-pre-wrap break-words leading-6">{item.text}</p>)}</div>)}</details>;
}

/** Diagnostics only. Product chat uses the same dsh runtime and approval gate. */
export default function AgentLabView() {
  const [text, setText] = useState("请调用 write_test_marker 工具，note 参数填「真实点击联调」。");
  const [lines, setLines] = useState<LogLine[]>([]);
  const [sending, setSending] = useState(false);
  const [auditLog, setAuditLog] = useState<ApprovalLogEntry[]>([]);
  const [reviewText, setReviewText] = useState(REVIEW_SAMPLE);
  const [reviewing, setReviewing] = useState(false);
  const [receipt, setReceipt] = useState<{approved:boolean;path:string;savedText:string|null;reason:string} | null>(null);
  const nextId = useRef(0);
  const [projectChoice, setProjectChoice] = useState<{ projectId: string | null } | null>(null);

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

  const leaveLab = () => {
    window.location.hash = "";
    window.location.reload();
  };

  const send = useCallback(async () => {
    if (sending || !text.trim() || !projectChoice) return;
    setSending(true);
    const sessionId = `lab-${Date.now()}`;
    append("prompt", `${sessionId} · ${text}`);
    try {
      await invoke("ai_bind_session_to_project", { sessionId, projectId: projectChoice.projectId });
      // ACP names sessions itself, so this is the agent's id, not the one sent.
      const acpSession = await invoke<string>("ai_send_prompt", { sessionId, text });
      append("accepted", `ACP 会话 ${acpSession}`);
    } catch (error) {
      append("send-error", String(error));
    } finally {
      setSending(false);
    }
  }, [append, sending, text, projectChoice]);

  useEffect(() => {
    if (!autorun || autorunFired.current || !projectChoice) return;
    autorunFired.current = true;
    const timer = window.setTimeout(() => void send(), 400);
    return () => window.clearTimeout(timer);
  }, [autorun, send, projectChoice]);

  const rehearse = async () => {
    if (reviewing) return;
    setReviewing(true); setReceipt(null);
    try {
      setReceipt(await invoke("ai_rehearse_text_approval", { proposedText: reviewText }));
      await refreshAudit();
    } catch (error) { append('审批演练失败', String(error)); }
    finally { setReviewing(false); }
  };

  return (
    <div className="flex h-screen flex-col gap-3 bg-background p-5 text-foreground">
      <div className="flex items-center justify-between gap-3">
        <h1 className="text-lg font-semibold">Agent 联调台（实验）</h1>
        <button
          type="button"
          onClick={leaveLab}
          className="rounded-lg bg-muted px-3 py-2 text-sm text-secondary-foreground transition-colors hover:bg-secondary"
        >
          返回主界面
        </button>
      </div>
      <p className="text-xs text-muted-foreground">
        用于人工验证 dsh 工具调用与审批通道（ACP 协议，
        <code className="px-1">dsh --profile acp</code>）。请先在设置中心保存 API Key；
        debug 构建仍允许使用 DEEPSEEK_API_KEY 环境变量兜底。
      </p>

      <details className="shrink-0 rounded-lg bg-muted/50 p-3 text-sm">
        <summary className="cursor-pointer font-medium">长文本审批演练（不写项目，无需模型）</summary>
        <p className="my-2 text-xs text-muted-foreground">首次为空；再次演练显示上次保存正文，便于核对覆盖。仅保存独立演练文本，审批日志保留原文与修改值。</p>
        <textarea aria-label="审批演练正文" value={reviewText} onChange={e => setReviewText(e.target.value)} rows={5} maxLength={20000} className="w-full rounded-lg bg-card p-3 leading-6" />
        <button type="button" onClick={() => void rehearse()} disabled={reviewing || !reviewText.trim()} className="mt-2 rounded-lg bg-secondary px-4 py-2 disabled:opacity-50">{reviewing ? '等待审批…' : '打开审批演练'}</button>
        {receipt && <details className="mt-2"><summary className="cursor-pointer">{receipt.approved ? '已保存批准正文' : '已拒绝，文件未改变'} · 查看实际文件回执</summary><p className="my-1 break-all text-xs">{receipt.path}</p><p className="max-h-40 overflow-auto whitespace-pre-wrap leading-6">{receipt.savedText ?? '（文件尚未创建）'}</p></details>}
      </details>

      {!projectChoice && <AiSessionProjectPicker legacy={false} onChoose={async projectId => setProjectChoice({ projectId })} />}
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
          disabled={sending || !projectChoice}
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
                  <AuditVersions argsJson={entry.argsJson} />
                </div>
              ))
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
