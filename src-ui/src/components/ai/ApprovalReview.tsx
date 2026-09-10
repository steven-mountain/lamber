import { useEffect, useState } from "react";
import { amendedArguments, readableArguments, remainingSeconds, type ApprovalPrompt } from "../../ai/approvalReview";

export function ApprovalReview({ current, onSettled, resolve, onStop }: { onStop?: () => Promise<unknown>; current: ApprovalPrompt; onSettled: (id: string) => void; resolve: (args: { requestId: string; approved: boolean; modifiedArgs: unknown }) => Promise<unknown> }) {
  const [values, setValues] = useState<Record<string, string>>(() => Object.fromEntries(current.intent?.fields.map(field => [field.key, field.proposedValue]) ?? []));
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [remaining, setRemaining] = useState(() => remainingSeconds(current.expiresAt));
  useEffect(() => {
    const tick = () => {
      const seconds = remainingSeconds(current.expiresAt);
      setRemaining(seconds);
      if (!seconds) onSettled(current.requestId);
    };
    tick();
    const timer = window.setInterval(tick, 500);
    return () => window.clearInterval(timer);
  }, [current, onSettled]);
  const modifiedArgs = amendedArguments(current, values);
  const respond = async (approved: boolean) => {
    if (busy || !remainingSeconds(current.expiresAt)) return;
    setBusy(true); setError('');
    try {
      await resolve({ requestId: current.requestId, approved, modifiedArgs: approved ? modifiedArgs ?? null : null });
      onSettled(current.requestId);
    } catch (cause) {
      setError(String(cause)); // Keep the review visible; a rejected edit must never silently approve.
    } finally { setBusy(false); }
  };
  return <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-foreground/40 p-4 backdrop-blur-sm">
    <section role="dialog" aria-modal="true" aria-labelledby="approval-heading" className="flex max-h-[90vh] w-full max-w-4xl flex-col rounded-xl bg-background p-5 shadow-xl">
      <header className="shrink-0">
        <h2 id="approval-heading" className="text-lg font-semibold text-foreground">审核 AI 写入内容</h2>
        <p className="mt-1 whitespace-pre-wrap text-sm text-muted-foreground">{current.reason ?? '该操作需要你确认后才会执行。'}</p>
        {current.intent ? <div className="mt-3 rounded-lg bg-muted/60 p-3 text-sm">
          <p>项目：<strong>{current.intent.projectName}</strong></p>
          <p>表单：{current.intent.templateName}</p>
          <p className="mt-1 text-xs text-muted-foreground">写入位置：{current.intent.targetDescription}</p>
        </div> : <p className="mt-2 text-sm">操作：{current.toolName}</p>}
      </header>
      <div className="mt-4 min-h-0 space-y-4 overflow-y-auto">
        {current.intent ? current.intent.fields.map(field => <section key={field.key} className="rounded-lg bg-muted/40 p-3">
          <div className="mb-2 flex items-center justify-between gap-3">
            <h3 className="font-semibold">{field.label}</h3>
            <span className="rounded-md bg-card px-2 py-1 text-xs text-secondary-foreground">{field.previousValue ? '覆盖已有内容' : '填入空白字段'}</span>
          </div>
          <div className="grid gap-3 md:grid-cols-2">
            <div><p className="mb-1 text-xs text-muted-foreground">当前保存内容</p>
              <div className="min-h-40 whitespace-pre-wrap break-words rounded-lg bg-card p-3 text-sm leading-7">{field.previousValue || '（尚未填写）'}</div>
            </div>
            <div><label className="mb-1 block text-xs text-muted-foreground" htmlFor={`review-${field.key}`}>拟写入内容（可修改）</label>
              <textarea id={`review-${field.key}`} value={values[field.key] ?? ''} disabled={busy || remaining === 0} rows={12} maxLength={20000}
                onChange={event => setValues(previous => ({...previous,[field.key]:event.target.value}))}
                className="w-full resize-y rounded-lg bg-card p-3 text-sm leading-7 outline-none focus:ring-2 focus:ring-ring" />
              <p className="text-right text-xs tabular-nums text-muted-foreground">{(values[field.key] ?? '').length} 字</p>
              {values[field.key] !== field.proposedValue && <details className="mt-1 text-xs text-muted-foreground"><summary className="cursor-pointer">查看 AI 原始建议</summary><p className="whitespace-pre-wrap p-2 leading-6">{field.proposedValue}</p></details>}
            </div>
          </div>
        </section>) : readableArguments(current.args).map((field,index) => <div key={index} className="rounded-lg bg-muted/50 p-3"><p className="text-xs text-muted-foreground">{field.label}</p><p className="mt-1 whitespace-pre-wrap break-words text-sm leading-7">{field.text}</p></div>)}
      </div>
      <footer className="mt-4 shrink-0">
        {error && <p role="alert" className="mb-2 whitespace-pre-wrap text-sm text-destructive">{error}</p>}
        <div className="flex flex-wrap items-end justify-between gap-3">
          <div className="text-xs text-muted-foreground"><p>最长审核时间 {Math.ceil(current.timeoutSeconds / 60)} 分钟，可先阅读、对照再修改。</p><p className="mt-1 tabular-nums">{remaining} 秒后自动拒绝；未确认不会写入。</p></div>
          <div className="flex flex-wrap gap-2">{onStop && <button type="button" disabled={busy || !remaining} onClick={() => { setBusy(true); void onStop().catch(error => setError(String(error))).finally(() => setBusy(false)); }} className="rounded-lg bg-muted px-4 py-2 text-sm disabled:opacity-50">停止本轮并拒绝</button>}<button type="button" disabled={busy || !remaining} onClick={() => void respond(false)} className="rounded-lg bg-muted px-4 py-2 text-sm disabled:opacity-50">拒绝</button>
            <button type="button" disabled={busy || !remaining} onClick={() => void respond(true)} className="rounded-lg bg-primary px-4 py-2 text-sm text-primary-foreground disabled:opacity-50">{modifiedArgs ? '修改后批准' : '批准写入'}</button></div>
        </div>
      </footer>
    </section>
  </div>;
}
