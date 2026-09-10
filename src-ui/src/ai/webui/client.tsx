import { LegacyHistory } from './LegacyHistory';
import AiAgentSettingsCard from '../../components/settings/AiAgentSettingsCard';
/* eslint-disable react-refresh/only-export-components -- Official plugin entry owns slot registration. */
import { createElement, useCallback, useEffect, useState, useSyncExternalStore } from 'react';
import { ApprovalReview } from '../../components/ai/ApprovalReview';
import type { ApprovalPrompt } from '../approvalReview';
import { installWebBusinessTransport, type BusinessCall } from '../../services/webBusinessTransport';
import { publishWebBusinessEvent } from '../../services/businessEvents';
import { BusinessWorkspace } from './BusinessWorkspace';
import type { AiImageAttachment } from '../types';

type Binding = { projectId: string | null; projectName: string | null };
type Bootstrap = { cwd: string; projects: { id: string; name: string }[] };
type Rpc = <T>(method: string, payload?: unknown) => Promise<T>;
// Official packages own slot scope and transport. Only this small deployment face is structural.
interface DeploymentContext {
  connection: { rpc: { call(channel: string, endpoint: string, payload: unknown): Promise<{ ok: boolean; value?: unknown; error?: { message: string } }> } };
  slots: { inject(name: string, setup: () => unknown): unknown; register(descriptor: { name: string; id?: string; order?: number; label?: string }, component: unknown): () => void };
  theme: { overrideTokens(id: string, values: Record<string, { light: string; dark: string }>): () => void };
  sessions: { list: { getSnapshot(): { current?: string; byId?: Record<string, { running: boolean }> }; subscribe(callback: () => void): () => void }; create(options: { cwd: string }): Promise<string>; refresh(): Promise<unknown>; open(id: string): void; scope(id: string): unknown };
  conversation: { input: { for(scope: unknown): { addImages(ids: string[]): boolean } }; createDraftImages(files: File[]): { id: string }[]; releaseDraftImages(images: { id: string }[]): void };
  effect(setup: () => unknown): void;
}
export const inject = ['slots', 'theme', 'connection', 'conversation', 'sessions'];

function ProjectBinding({ sessionId, call, attach, sessionStore }: { sessionId: string; call: Rpc; attach: (image: AiImageAttachment) => void; sessionStore: DeploymentContext['sessions']['list'] }) {
  const running = useSyncExternalStore(sessionStore.subscribe, () => sessionStore.getSnapshot().byId?.[sessionId]?.running ?? false);
  const [binding, setBinding] = useState<Binding | null | undefined>();
  const [projects, setProjects] = useState<Bootstrap['projects']>([]);
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  useEffect(() => {
    let disposed = false;
    setBinding(undefined);
    void Promise.all([call<Binding | null>('binding', { sessionId }), call<Bootstrap>('bootstrap')])
      .then(([next, info]) => { if (!disposed) { setBinding(next); setProjects(info.projects); setError(''); } })
      .catch(reason => { if (!disposed) setError(String(reason)); });
    return () => { disposed = true; };
  }, [sessionId, call]);
  const bind = async (projectId: string | null) => {
    setBusy(true); setError('');
    try { setBinding(await call<Binding>('bind', { sessionId, projectId })); }
    catch (cause) { setError(String(cause)); }
    finally { setBusy(false); }
  };
  return <div className="lamber-business-surface">
    {binding ? <span>{binding.projectId ? `已绑定：${binding.projectName}` : '通用聊天 · 聚合只读及甄选费计算'}</span>
      : <label>会话权限 <select aria-label="选择会话项目" disabled={busy || binding === undefined} value="" onChange={event => void bind(event.target.value === '__general' ? null : event.target.value)}>
        <option value="">请先选择项目或通用聊天</option><option value="__general">通用聊天</option>
        {projects.map(project => <option key={project.id} value={project.id}>{project.name}</option>)}
      </select></label>}
    {error && <p role="alert">{error}</p>}
    {binding?.projectId && <BusinessWorkspace sessionId={sessionId} call={call} attach={attach} disabled={running} />}
  </div>;
}
function StartSession({ call, sessions }: { call: Rpc; sessions: DeploymentContext['sessions'] }) {
  const [info, setInfo] = useState<Bootstrap>();
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  useEffect(() => { let active = true; void call<Bootstrap>('bootstrap').then(value => { if (active) setInfo(value); }).catch(error => { if (active) setError(String(error)); }); return () => { active = false; }; }, [call]);
  const start = async (project: string) => {
    if (!info || busy) return;
    setBusy(true); setError('');
    try {
      const created = await call<{ sessionId: string }>('new-session', { projectId: project === '__general' ? null : project });
      await sessions.refresh();
      sessions.open(created.sessionId);
    } catch (error) { setError(String(error)); }
    finally { setBusy(false); }
  };
  return <div className="lamber-business-surface"><label>选择项目并开始 <select aria-label="新会话项目" value="" disabled={busy || !info} onChange={event => void start(event.target.value)}>
    <option value="">选择项目或通用聊天</option><option value="__general">通用聊天</option>{info?.projects.map(project => <option key={project.id} value={project.id}>{project.name}</option>)}
  </select></label>{busy && <span>正在创建会话…</span>}{error && <p role="alert">{error}</p>}</div>;
}
function ApprovalQueue({ call }: { call: Rpc }) {
  const [queue, setQueue] = useState<ApprovalPrompt[]>([]);
  const [error, setError] = useState('');
  useEffect(() => {
    let disposed = false;
    let timer: ReturnType<typeof setTimeout>;
    const refresh = async () => {
      try { const next = await call<ApprovalPrompt[]>('pending'); if (!disposed) { setQueue(next); setError(''); } }
      catch (cause) { if (!disposed) setError(String(cause)); }
      finally { if (!disposed) timer = setTimeout(() => void refresh(), 750); }
    };
    void refresh();
    return () => { disposed = true; clearTimeout(timer); };
  }, [call]);
  const settled = useCallback((id: string) => setQueue(items => items.filter(item => item.requestId !== id)), []);
  return <div className="lamber-business">{error && <span title={error}>业务连接暂时中断</span>}{queue[0] && <ApprovalReview key={queue[0].requestId} current={queue[0]} onSettled={settled} onStop={() => call('stop-session', { sessionId: queue[0].sessionId })} resolve={args => call('resolve', args)} />}</div>;
}

export function apply(ctx: DeploymentContext) {
  const call: Rpc = async (method, payload = {}) => {
    const response = await ctx.connection.rpc.call('/lamber', method, payload);
    if (!response.ok) throw new Error(response.error?.message || '业务操作未完成');
    return response.value as never;
  };
  const business: BusinessCall = async (method, payload) => {
    if (method === 'settings' || method === 'save-settings') {
      const result = await call(method, payload);
      if (method === 'save-settings') void call('restart').catch(() => {});
      return result as never;
    }
    const sessionId = (payload as { sessionId: string }).sessionId;
    const requestId = crypto.randomUUID();
    await call('business', { sessionId, requestId, method, payload });
    const deadline = Date.now() + 20 * 60_000;
    while (Date.now() < deadline) {
      const next = await call<{ state: string; result?: { ok: boolean; value?: unknown; error?: string } }>('business-result', { sessionId, requestId });
      if (next.result) {
        if (!next.result.ok) throw new Error(next.result.error || '业务操作未完成');
        void call('release-read', { sessionId, requestId }).catch(() => {});
        if (['template-action', 'replace-image', 'upload-demand'].includes(method)) {
          publishWebBusinessEvent('lamber-template-text-changed');
          publishWebBusinessEvent('lamber-demand-assets-changed');
        }
        return next.result.value as never;
      }
      await new Promise(resolve => setTimeout(resolve, 500));
    }
    throw new Error('操作仍未收到最终回执，请在操作记录和主窗口核对。系统不会自动重做。');
  };
  ctx.effect(() => installWebBusinessTransport(business));
  ctx.effect(() => {
    let disposed = false;
    let unlisten: (() => void) | undefined;
    // The loopback port is ephemeral. Selection belongs to the desktop workspace,
    // not to browser localStorage for one transient origin.
    void (async () => {
      const saved = await call<{ sessionId: string | null }>('selected-session');
      await ctx.sessions.refresh();
      if (disposed) return;
      if (saved.sessionId) {
        try {
          ctx.sessions.open(saved.sessionId);
        }
        catch (error) { console.warn('原会话未自动打开，请从历史中恢复。', error); }
      }
      let previous = saved.sessionId;
      const changed = () => {
        const current = ctx.sessions.list.getSnapshot().current;
        if (!current || current === previous) return;
        previous = current;
        void call('select-session', { sessionId: current }).catch(console.error);
      };
      unlisten = ctx.sessions.list.subscribe(changed);
      changed();
    })().catch(console.error);
    return () => { disposed = true; unlisten?.(); };
  });
  const attachToSession = (sessionId: string, image: AiImageAttachment) => {
    const scope = ctx.sessions.scope(sessionId);
    if (!scope || !image.dataUrl) throw new Error('当前会话不能添加图片，请稍后重试。');
    const [metadata, encoded] = image.dataUrl.split(',');
    if (!metadata.endsWith(';base64') || !encoded) throw new Error('图片数据不可用，请重新附加。');
    const bytes = Uint8Array.from(atob(encoded), char => char.charCodeAt(0));
    const images = ctx.conversation.createDraftImages([new File([bytes], image.name, { type: image.mimeType })]);
    if (!ctx.conversation.input.for(scope).addImages(images.map(item => item.id))) {
      ctx.conversation.releaseDraftImages(images);
      throw new Error('输入框正在发送，请稍后重新附加。');
    }
  };
  ctx.slots.inject('conversation.hero.agentPreset', function* () {
    yield ctx.slots.register({ name: 'conversation.hero.agentPreset' }, ({ useSessions }: { useSessions: <T>(selector: (state: { current?: string }) => T) => T }) => {
      const current = useSessions(state => state.current);
      return current ? <ProjectBinding sessionStore={ctx.sessions.list} sessionId={current} call={call} attach={image => attachToSession(current, image)} /> : <StartSession call={call} sessions={ctx.sessions} />;
    });
  });
  ctx.slots.inject('settings.section', function* () {
    yield ctx.slots.register({ name: 'settings.section', id: 'lamber-model', order: 5, label: 'Lamber 模型与服务' }, () => <div className="lamber-business"><AiAgentSettingsCard showDiagnostics={false} /></div>);
  });
  ctx.slots.inject('conversation.hero.brand.mark', function* () {
    yield ctx.slots.register({ name: 'conversation.hero.brand.mark' }, () => <strong aria-label="Lamber">L</strong>);
  });
  ctx.slots.inject('sidebar.brand.mark', () => ctx.slots.inject('sidebar.brand.name', function* () {
    yield ctx.slots.register({ name: 'sidebar.brand.mark' }, ({ size }: { size: number }) => createElement('span', { style: { fontSize: size, fontWeight: 600 } }, 'L'));
    yield ctx.slots.register({ name: 'sidebar.brand.name' }, () => <strong>Lamber AI</strong>);
  }));
  ctx.slots.inject('conversation.composer.dock', function* () {
    yield ctx.slots.register({ name: 'conversation.composer.dock', id: 'lamber-project' }, ({ sessionId }: { sessionId: string }) => <ProjectBinding sessionStore={ctx.sessions.list} sessionId={sessionId} call={call} attach={image => attachToSession(sessionId, image)} />);
  });
  ctx.slots.inject('sidebar.footer.action', function* () {
    yield ctx.slots.register({ name: 'sidebar.footer.action', id: 'lamber-history' }, () => <LegacyHistory call={call} openSession={async id => { await ctx.sessions.refresh(); ctx.sessions.open(id); }} />);
    yield ctx.slots.register({ name: 'sidebar.footer.action', id: 'lamber-approval' }, () => <ApprovalQueue call={call} />);
  });
  const font = 'Inter, "Microsoft YaHei", "PingFang SC", system-ui, sans-serif';
  ctx.effect(() => ctx.theme.overrideTokens('lamber', { '--dsw-font-family': { light: font, dark: font } }));
}
