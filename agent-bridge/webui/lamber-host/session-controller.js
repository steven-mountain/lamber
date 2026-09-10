import { SessionController } from '@deepseek-ai/dsh-api-session-controller';
import { postBridge } from 'dsh-tool-lamber/lib/bridge.js';
const call = (method, args) => postBridge(`/lamber-webui/${method}`, args, AbortSignal.timeout(15000));
/** Bounded deployment adapter; official history, admission and streaming remain upstream. */
export default class LamberSessionController extends SessionController {
  lamberWebDeployment = true;
  async inspect(sessionId, signal) {
    const value = await super.inspect(sessionId, signal);
    if (value.meta.cwd !== process.cwd()) throw new Error('会话不属于当前工作区');
    return value;
  }
  async list(_request, signal) {
    const value = await super.list(_request, signal);
    return { ...value, items: value.items.filter(item => item.cwd === process.cwd()) };
  }
  async search(request, signal) {
    const allowed = new Set((await this.list({}, signal)).items.map(item => item.sessionId));
    const value = await super.search(request, signal);
    return { ...value, items: value.items.filter(item => allowed.has(item.sessionId)) };
  }
  async *follow(request, signal) {
    if (request.address?.kind !== 'session') throw new Error('此部署仅开放当前工作区的业务会话');
    await this.inspect(request.address.sessionId, signal);
    yield* super.follow(request, signal);
  }
  async *control(signal) {
    for await (const frame of super.control(signal)) {
      if (frame.type === 'baseline') {
        const ids = new Set((await this.list({}, signal)).items.map(item => item.sessionId));
        const own = values => Object.fromEntries(Object.entries(values).filter(([id]) => ids.has(id)));
        yield { ...frame, value: { queues: own(frame.value.queues), jobs: own(frame.value.jobs), projections: own(frame.value.projections) } };
      } else if (this.ctx.sessions.get(frame.sessionId)?.header.cwd === process.cwd()) yield frame;
    }
  }
  async create(request) {
    const cwd = request.workspaceId ? this.ctx.workspaceRegistry.get(request.workspaceId)?.path : request.cwd ?? process.cwd();
    if (cwd !== process.cwd()) throw new Error('只能在当前 Lamber 工作区创建或恢复会话');
    if (request.agentPreset && request.agentPreset !== 'lamber') throw new Error('此部署仅使用 Lamber 业务助手');
    return super.create(request);
  }
  async prompt(request, signal) {
    await call('admit', { sessionId: request.sessionId, cwd: process.cwd() });
    return super.prompt(request, signal);
  }
  async fork(request) {
    await call('admit', { sessionId: request.sessionId, cwd: process.cwd() });
    const child = await super.fork(request);
    await call('inherit', { sessionId: child.sessionId, parentSessionId: request.sessionId, cwd: process.cwd() });
    return child;
  }
}
