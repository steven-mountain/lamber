import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { useEffect, useState } from 'react';
import DocumentGenerationCards from '../../components/ai/DocumentGenerationCards';
import TechItemsCard from '../../components/ai/TechItemsCard';
import InquiryCard from '../../components/ai/InquiryCard';
import TemplateImageCard from '../../components/ai/TemplateImageCard';
import DemandImageCompletionCards from '../../components/ai/DemandImageCompletionCards';
import StructureReverseCard from '../../components/ai/StructureReverseCard';
import type { BusinessPresentation } from './businessPresentation';
import type { AiImageAttachment } from '../types';
import type { BusinessCall } from '../../services/webBusinessTransport';
const sections = [['documents','生成文档'],['technical','技术清单'],['inquiry','询价清单'],['images','项目图片'],['reverse','结构反算'],['receipts','操作记录']] as const;
interface Receipt { requestId: string; method: string; state: string; createdAt: string; result?: { ok: boolean; summary?: string; error?: string; value?: unknown } }
type Section = typeof sections[number][0];
export function BusinessWorkspace({ sessionId, call, attach, disabled }: { sessionId: string; call: BusinessCall; disabled: boolean; attach: (image: AiImageAttachment) => void }) {
  const [section, setSection] = useState<Section | null>(null);
  const [message, setMessage] = useState('');
  const [presentation, setPresentation] = useState<BusinessPresentation>();
  const [receipts, setReceipts] = useState<Receipt[]>([]);
  const businessId = `web:${sessionId}`;
  useEffect(() => { setSection(null); setMessage(''); setPresentation(undefined); setReceipts([]); }, [sessionId]);
  useEffect(() => {
    let active = true; let timer: ReturnType<typeof setTimeout>;
    const refresh = async () => {
      try { const next = await call<BusinessPresentation>('business-presentation', { sessionId }); if (active) setPresentation(next); }
      catch { /* Read-only invitation lookup does not replace business action errors. */ }
      finally { if (active) timer = setTimeout(() => void refresh(), 1500); }
    };
    void refresh(); return () => { active = false; clearTimeout(timer); };
  }, [sessionId, call]);

  useEffect(() => {
    if (section !== 'receipts') return;
    let active = true;
    void call<Receipt[]>('receipts', { sessionId: businessId }).then(value => { if (active) setReceipts(value); })
      .catch(error => { if (active) setMessage(String(error)); });
    return () => { active = false; };
  }, [businessId, section, call]);
  return <div className="lamber-business">
    {presentation && (presentation.lists.tech || presentation.lists.inquiry || presentation.documents.length > 0 || presentation.reverse.requested || presentation.demandImages || presentation.savedImages) && <p className="text-caption">本轮相关操作：{[presentation.lists.tech && '技术清单（可采用本轮建议）', presentation.lists.inquiry && '询价清单', presentation.documents.length > 0 && '生成文档', presentation.reverse.requested && '结构反算', (presentation.demandImages || presentation.savedImages) && '项目图片'].filter(Boolean).join('、')}。请点击下方对应入口核对。</p>}
    <div className="lamber-business-actions">{sections.map(([id, label]) => <button type="button" key={id} onClick={() => { setMessage(''); setSection(id); }}>{label}</button>)}</div>
    {section && <div className="lamber-business-backdrop"><section className="lamber-business-panel" role="dialog" aria-modal="true" aria-label={sections.find(([id]) => id === section)?.[1]}>
      <header><strong>{sections.find(([id]) => id === section)?.[1]}</strong><button type="button" aria-label="关闭业务面板" onClick={() => setSection(null)}>关闭</button></header>
      <div className="lamber-business-content">
        {section === 'documents' && <DocumentGenerationCards sessionId={businessId} templateIds={['demand','meeting','approval','selection','presales','benefit','decision']} refreshToken={0} disabled={disabled} onReceipt={setMessage} />}
        {section === 'technical' && <TechItemsCard sessionId={businessId} proposal={presentation?.proposal ?? []} disabled={disabled} onReceipt={setMessage} />}
        {section === 'inquiry' && <InquiryCard sessionId={businessId} disabled={disabled} onReceipt={setMessage} />}
        {section === 'images' && <><TemplateImageCard sessionId={businessId} disabled={disabled} onReceipt={setMessage} onAnalyze={image => { try { attach(image); setMessage('图片已加入当前会话输入框；请核对预览并选择视觉模型后发送。'); } catch (error) { setMessage(String(error)); } }} /><DemandImageCompletionCards sessionId={businessId} refreshToken={0} disabled={disabled} onReceipt={setMessage} /></>}
        {section === 'reverse' && <StructureReverseCard sessionId={businessId} initialMetric={presentation?.reverse.metricType ?? "margin"} initialTarget={presentation?.reverse.targetPercent ?? ""} initialScenario={presentation?.reverse.scenario ?? ""} disabled={disabled} onReceipt={setMessage} />}
        {section === 'receipts' && <><p>以下回执描述操作发生时的结果；后续编辑和正式保存请以主窗口当前状态为准。历史预览不能再次批准或执行。</p>{receipts.length ? receipts.map(receipt => <article key={receipt.requestId} className="rounded-lg bg-muted p-3"><p className="text-caption">{receipt.createdAt} UTC · {receipt.state === 'queued' ? '等待主窗口接收' : receipt.state === 'running' ? '正在处理，请勿重复操作' : '历史结果'}</p>{receipt.result?.summary || receipt.result?.error ? <div className="lamber-business-receipt"><ReactMarkdown remarkPlugins={[remarkGfm]}>{receipt.result.summary || receipt.result.error}</ReactMarkdown></div> : receipt.result?.value ? <details><summary>查看早期记录的原始结果</summary><pre>{JSON.stringify(receipt.result.value, null, 2)}</pre></details> : null}</article>) : <p>此会话暂无业务操作记录。</p>}</>}
        {message && (message.length > 300 ? <details><summary>查看本次操作回执</summary><pre>{message}</pre></details> : <p className="whitespace-pre-wrap" role="status">{message}</p>)}
      </div>
    </section></div>}
  </div>;
}
