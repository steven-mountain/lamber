import { useEffect, useRef, useState } from 'react';
import { listen } from '@tauri-apps/api/event';
import { listChatTemplateImages, readChatTemplateImage, prepareReplacement, replaceChatTemplateImage,
  type ImageTarget, type ReplacementImage } from '../../services/chatTemplateImages';
import { subscribeDemandAssetsChanged } from '../../services/demandTemplateAssets';
import type { AiImageAttachment } from '../../ai/types';

const slot = (usage: string) => usage === 'attach1' ? '附件1 · 客户确认材料' : '附件2 · 公示截图';

export default function TemplateImageCard({ sessionId, disabled, onReceipt, onAnalyze }: {
  sessionId: string; disabled: boolean; onReceipt: (content: string) => void;
  onAnalyze: (image: AiImageAttachment) => void;
}) {
  const [targets, setTargets] = useState<ImageTarget[]>([]);
  const [selected, setSelected] = useState('');
  const [original, setOriginal] = useState<Awaited<ReturnType<typeof readChatTemplateImage>> | null>(null);
  const [candidate, setCandidate] = useState<ReplacementImage | null>(null);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [receipt, setReceipt] = useState('');
  const [comparison, setComparison] = useState<{ before: string; after: string } | null>(null);
  const [revision, setRevision] = useState(0);
  const active = useRef(true);
  const operation = useRef(0);
  const writing = useRef(false);
  const target = targets.find(item => item.assetId === selected);
  const targetRef = useRef(target); targetRef.current = target;

  useEffect(() => {
    active.current = true;
    const stop = subscribeDemandAssetsChanged(() => setRevision(value => value + 1));
    const ws = listen('lamber-workspace-state-changed', () => {
      active.current = false;
      operation.current++; setTargets([]); setSelected(''); setOriginal(null); setCandidate(null);
      setComparison(null); setReceipt(''); setRevision(value => value + 1);
    });
    return () => { active.current = false; stop(); void ws.then(fn => fn()).catch(console.warn); };
  }, []);

  useEffect(() => {
    let current = true;
    setLoading(true);
    listChatTemplateImages(sessionId).then(next => {
      if (!current) return;
      setTargets(next); setSelected(previous => next.some(t => t.assetId === previous) ? previous : next[0]?.assetId || '');
      setError('');
    }).catch(reason => { if (current) { setTargets([]); setSelected(''); setError(String(reason)); } })
      .finally(() => { if (current) setLoading(false); });
    return () => { current = false; };
  }, [sessionId, revision]);

  useEffect(() => {
    let current = true;
    operation.current++; setOriginal(null); setCandidate(null); setBusy(false);
    const requested = targetRef.current;
    if (requested) void readChatTemplateImage(requested).then(image => { if (current) setOriginal(image); })
      .catch(reason => { if (current) setError(`图片无法读取：${String(reason)}`); });
    return () => { current = false; };
    // Asset IDs are immutable; metadata/list refresh does not discard a prepared replacement.
  }, [selected, sessionId]);

  const prepare = async (file?: File) => {
    if (!file || disabled || busy || writing.current) return;
    const id = ++operation.current;
    setBusy(true); setCandidate(null); setError(''); setComparison(null); setReceipt('');
    try {
      const next = await prepareReplacement(file);
      if (active.current && operation.current === id) setCandidate(next);
    } catch (reason) { if (active.current && operation.current === id) setError(String(reason)); }
    finally { if (active.current && operation.current === id) setBusy(false); }
  };

  const replace = async () => {
    if (!target || !candidate || !original || disabled || writing.current) return;
    const id = operation.current;
    writing.current = true; setBusy(true); setError('');
    try {
      const result = await replaceChatTemplateImage(target, candidate,
        () => active.current && operation.current === id && targetRef.current?.assetId === target.assetId);
      const message = `图片已替换并保存：${target.projectName} / ${target.templateName} / ${slot(target.usage)}。\n${target.name}（${original.width ?? '?'}×${original.height ?? '?'}，${target.assetId}） → ${candidate.name}（${candidate.width}×${candidate.height}，${result.assetId}）。\n仅替换选中的这一张图片，其余附件保留。${result.refreshWarning}`;
      onReceipt(message);
      if (active.current) {
        setReceipt(message); setComparison({ before: original.dataUrl, after: candidate.dataUrl });
        setCandidate(null); setRevision(value => value + 1);
      }
    } catch (reason) { if (active.current) setError(`未收到替换成功回执：${String(reason)} 请刷新核对当前图片。`); }
    finally { writing.current = false; if (active.current) setBusy(false); }
  };

  return <section aria-label="项目图片" className="space-y-3 rounded-lg bg-muted/50 p-4 text-caption">
    <p className="font-semibold">项目图片 · 需求导入表已保存附件</p>
    {loading && <p role="status">正在读取图片列表…</p>}
    {!loading && !targets.length && !error && <p>此项目没有已保存的需求表附件图片。</p>}
    <div className="flex flex-wrap gap-2">{targets.map(item => <button key={item.assetId}
      disabled={busy} aria-pressed={selected === item.assetId}
      className={`rounded-md px-3 py-2 text-left ${selected === item.assetId ? 'bg-card shadow-sm' : 'bg-muted'}`}
      onClick={() => { setSelected(item.assetId); setError(''); setComparison(null); setReceipt(''); }}>
      {slot(item.usage)} · {item.name}
    </button>)}</div>
    {target && <>
      <p className="break-all text-muted-foreground">{target.projectName} / {target.templateName} / {slot(target.usage)}</p>
      <div className={`grid gap-3 ${candidate ? 'sm:grid-cols-2' : ''}`}>
        <figure className="rounded-md bg-card p-3"><figcaption className="mb-2">{candidate ? '变更前' : '当前图片'} · {target.name}</figcaption>
          {original ? <a href={original.dataUrl} target="_blank" rel="noreferrer" title="查看完整图片"><img src={original.dataUrl} alt={`原图：${target.name}`} className="max-h-72 w-full object-contain" /></a> : <p>图片加载中或不可用，请重试。</p>}
        </figure>
        {candidate && <figure className="rounded-md bg-card p-3"><figcaption className="mb-2">变更后（尚未保存） · {candidate.name}</figcaption>
          <img src={candidate.dataUrl} alt={`待替换：${candidate.name}`} className="max-h-72 w-full object-contain" />
        </figure>}
      </div>
      <div className="flex flex-wrap gap-2" tabIndex={0} onPaste={event => {
        const file = event.clipboardData.files[0]; if (file) { event.preventDefault(); void prepare(file); }
      }}>
        <button className="rounded-md bg-card px-3 py-2" disabled={disabled || busy || !original}
          onClick={async () => {
            try {
              const image = await readChatTemplateImage(target);
              if (!active.current || targetRef.current?.assetId !== target.assetId) return;
              onAnalyze({ id: image.id, name: image.name, mimeType: image.mimeType, size: image.size, dataUrl: image.dataUrl,
                source: 'template_asset', projectId: target.projectId, templateId: target.templateName, assetId: target.assetId, fieldKey: target.usage });
            } catch (reason) { if (active.current) setError(String(reason)); }
          }}>加入输入框供 AI 分析</button>
        <label className="cursor-pointer rounded-md bg-card px-3 py-2">选择替换图片
          <input aria-label="选择替换图片" type="file" accept="image/png,image/jpeg,image/webp" className="sr-only" disabled={disabled || busy || !original}
            onChange={event => { const file = event.target.files?.[0]; event.target.value = ''; void prepare(file); }} />
        </label>
        {candidate && <><button className="rounded-md bg-primary-soft px-3 py-2 font-semibold" disabled={disabled || busy || !original} onClick={() => void replace()}>确认替换这张图片</button>
          <button className="rounded-md bg-card px-3 py-2" disabled={busy} onClick={() => setCandidate(null)}>取消替换</button></>}
      </div>
      <p className="text-muted-foreground">可在操作区粘贴新图。选择后先对照，确认才保存；PNG / JPEG / WEBP，最大20MB。</p>
    </>}
    {busy && <p role="status">正在处理图片…</p>}
    {error && <p role="alert" className="text-destructive">{error}</p>}
    <button className="rounded-md bg-card px-3 py-2" disabled={busy} onClick={() => { setSelected(''); setRevision(value => value + 1); }}>刷新图片</button>
    {receipt && <p role="status" className="whitespace-pre-wrap">{receipt}</p>}
    {comparison && <div className="grid grid-cols-2 gap-3">
      <figure><figcaption>已替换 · 变更前</figcaption><img src={comparison.before} alt="保存结果：变更前" className="max-h-48 w-full object-contain" /></figure>
      <figure><figcaption>已替换 · 变更后</figcaption><img src={comparison.after} alt="保存结果：变更后" className="max-h-48 w-full object-contain" /></figure>
    </div>}
  </section>;
}
