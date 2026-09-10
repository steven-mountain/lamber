import { useEffect, useState } from 'react';
import { listenBusinessEvent as listen } from '../../services/businessEvents';
import { loadDemandUploadTargets, type DemandUploadTarget } from '../../services/demandUploadTargets';
import { subscribeDemandAssetsChanged } from '../../services/demandTemplateAssets';
import DemandImageCompletionCard from './DemandImageCompletionCard';

export default function DemandImageCompletionCards({ sessionId, refreshToken, disabled, onReceipt }: {
  sessionId: string; refreshToken: number; disabled: boolean; onReceipt: (content: string) => void;
}) {
  const [targets, setTargets] = useState<DemandUploadTarget[]>([]);
  const [error, setError] = useState('');
  const [revision, setRevision] = useState(0);
  useEffect(() => {
    const refresh = () => setRevision(value => value + 1);
    const invalidate = () => { setTargets([]); refresh(); };
    const unsubscribe = subscribeDemandAssetsChanged(refresh);
    const subscriptions = [listen('lamber-template-text-changed', refresh), listen('lamber-workspace-state-changed', invalidate)]
      .map(promise => promise.catch(reason => { console.warn(reason); return () => {}; }));
    window.addEventListener('focus', refresh);
    return () => {
      unsubscribe(); window.removeEventListener('focus', refresh);
      subscriptions.forEach(promise => { void promise.then(stop => stop()).catch(console.warn); });
    };
  }, []);
  useEffect(() => {
    let active = true;
    loadDemandUploadTargets(sessionId).then(next => {
      if (active) { setTargets(next); setError(''); }
    }).catch(reason => { if (active) { setTargets([]); setError(String(reason)); } });
    return () => { active = false; };
  }, [sessionId, refreshToken, revision]);
  return <>
    {targets.map(target => <DemandImageCompletionCard
      key={`${target.sessionId}-${target.workspaceId}-${target.projectId}-${target.templateName}-${target.usage}`}
      target={target} disabled={disabled} onReceipt={onReceipt} />)}
    {error && <div role="alert" className="rounded-lg bg-muted/50 p-4 text-caption">
      附件卡片读取失败：{error}
      <button className="ml-2 rounded-md bg-card px-3 py-2" onClick={() => setRevision(value => value + 1)}>重试</button>
    </div>}
  </>;
}
