import { useEffect, useRef, useState } from 'react';
import { listen } from '@tauri-apps/api/event';
import { subscribeDemandAssetsChanged } from '../../services/demandTemplateAssets';
import { completionSummary, documentReceipt, loadDocumentTargets, requestDocumentGeneration, type DocumentTarget } from '../../services/chatDocumentGeneration';

export default function DocumentGenerationCards({ sessionId, templateIds, refreshToken, disabled, onReceipt }: {
  sessionId: string; templateIds: string[]; refreshToken: number; disabled: boolean; onReceipt: (content: string) => void;
}) {
  const [targets, setTargets] = useState<DocumentTarget[]>([]);
  const [error, setError] = useState('');
  const [revision, setRevision] = useState(0);
  const [busy, setBusy] = useState(false);
  const busyRef = useRef(false);
  const ids = templateIds.join(',');
  useEffect(() => {
    const refresh = () => setRevision(value => value + 1);
    const invalidate = () => { setTargets([]); refresh(); };
    const unsubscribe = subscribeDemandAssetsChanged(refresh);
    const subscriptions = [listen('lamber-template-text-changed', refresh), listen('lamber-workspace-state-changed', invalidate)];
    window.addEventListener('focus', refresh);
    return () => { unsubscribe(); window.removeEventListener('focus', refresh); subscriptions.forEach(promise => { void promise.then(stop => stop()).catch(console.warn); }); };
  }, []);
  useEffect(() => {
    let active = true;
    void loadDocumentTargets(sessionId, ids.split(',')).then(next => { if (active) { setTargets(next); setError(''); } })
      .catch(reason => { if (active) { setTargets([]); setError(String(reason)); } });
    return () => { active = false; };
  }, [sessionId, ids, refreshToken, revision]);
  const generate = async (target: DocumentTarget) => {
    if (busyRef.current) return;
    busyRef.current = true; setBusy(true); setError('');
    try { onReceipt(documentReceipt(target, await requestDocumentGeneration(target))); }
    catch (reason) { const message = String(reason); setError(message); onReceipt(documentReceipt(target, { status: 'error', message })); }
    finally { busyRef.current = false; setBusy(false); setRevision(value => value + 1); }
  };
  return <>
    {targets.map(target => {
      const summary = completionSummary(target.completion);
      return <div key={target.templateName} className="rounded-lg bg-muted/50 p-4 text-body space-y-3">
        <div className="font-semibold">生成文档 · {target.templateName}</div>
        <div className="text-caption text-secondary-foreground">{target.projectName} · 按已保存表单检查</div>
        <div className="rounded-md bg-card p-3 space-y-2">
          <div className="tabular-nums">完成度 {summary.filledCount}/{summary.totalCount}</div>
          {summary.missing.length > 0 && <p>仍缺：{summary.missing.map(item => item.label).join('、')}。仍可生成，空缺内容请自行核对。</p>}
          {summary.unknown.length > 0 && <p>尚有 {summary.unknown.length} 项未能判断：{summary.unknown.map(item => item.label).join('、')}。不能确认整表已齐备。</p>}
          {!summary.missing.length && !summary.unknown.length && <p>目录检查项已齐备，生成时仍执行产品原有核验。</p>}
        </div>
        <p className="text-caption text-secondary-foreground">点击后在主窗口调用原有生成流程；如有覆盖或核验提示，请在主窗口处理。</p>
        <button type="button" disabled={disabled || busy} className="rounded-md bg-primary px-4 py-2 text-primary-foreground disabled:opacity-50" onClick={() => void generate(target)}>{busy ? '正在生成，请查看主窗口…' : '生成文档'}</button>
      </div>;
    })}
    {error && <div role="alert" className="rounded-lg bg-muted/50 p-4 text-caption">生成卡片：{error}<button className="ml-2 rounded-md bg-card px-3 py-2" onClick={() => setRevision(value => value + 1)}>重试读取</button></div>}
  </>;
}
