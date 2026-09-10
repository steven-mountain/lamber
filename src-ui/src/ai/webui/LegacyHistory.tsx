import { useState } from 'react';
import type { BusinessCall } from '../../services/webBusinessTransport';
interface LegacySummary { id: string; title: string; messageCount: number; status: string; harnessSessionId: string | null }
interface LegacyRecord { id: string; title: string; messages: { role: string; content: string; appReceipt?: boolean; images?: { id: string; name: string }[] }[] }
export function LegacyHistory({ call, openSession }: { call: BusinessCall; openSession: (id: string) => Promise<void> }) {
  const [opened, setOpened] = useState(false);
  const [rows, setRows] = useState<LegacySummary[]>([]);
  const [record, setRecord] = useState<LegacyRecord>();
  const [error, setError] = useState('');
  const show = async () => {
    setOpened(true); setError('');
    try {
      const [rows, status] = await Promise.all([call<LegacySummary[]>('legacy-list', {}), call<{ warning?: string }>('legacy-status', {})]);
      setRows(rows); setRecord(undefined);
      if (status.warning) setError(`旧记录迁移未完成：${status.warning}。原始记录保留，以下仅显示已成功迁移的记录。`);
    }
    catch (error) { setError(String(error)); }
  };
  const read = async (id: string) => {
    setError('');
    try { setRecord(await call<LegacyRecord>('legacy-read', { id })); }
    catch (error) { setError(String(error)); }
  };
  const restore = async (id: string) => {
    setError('');
    try { const next = await call<{ sessionId: string }>('legacy-restore', { id }); await openSession(next.sessionId); setOpened(false); }
    catch (error) { setError(`原会话未恢复：${String(error)}。旧记录仍可阅读，请勿当作上下文已恢复。`); }
  };
  return <div className="lamber-business"><button type="button" title="旧会话记录" aria-label="旧会话记录" onClick={() => void show()}>历史</button>
    {opened && <div className="lamber-business-backdrop"><section className="lamber-business-panel" role="dialog" aria-modal="true" aria-label="旧会话记录">
      <header><strong>旧会话记录</strong><button type="button" onClick={() => setOpened(false)}>关闭历史</button></header>
      <div className="lamber-business-content"><p>原记录完整保留。这里仅供阅读；恢复必须同时找到后端绑定和原 dsh 历史，旧业务动作不会重放。</p>
        {error && <p role="alert">{error}</p>}
        {record ? <><button type="button" onClick={() => setRecord(undefined)}>返回记录列表</button><h3>{record.title}</h3>{record.messages.map((message, index) => <article key={index}>
          <strong>{message.appReceipt ? '历史业务回执' : message.role === 'user' ? '用户' : '助手'}</strong><pre>{message.content}</pre>
          {message.images?.map(image => <p key={image.id}>附件：{image.name}（历史元数据，未自动发送；继续分析时请重新附加图片）</p>)}
        </article>)}</> : rows.length ? rows.map(row => <article key={row.id}><h3>{row.title}</h3><p>{row.messageCount} 条消息 · {row.status}</p><div className="lamber-business-actions"><button type="button" onClick={() => void read(row.id)}>读取旧记录</button>{row.harnessSessionId && <button type="button" onClick={() => void restore(row.id)}>恢复原会话</button>}</div></article>) : <p>当前工作区没有待迁移旧记录。其他工作区的已绑定记录请在原工作区查看。</p>}
      </div>
    </section></div>}
  </div>;
}
