import { WorkspaceController } from '@deepseek-ai/dsh-api-workspace-controller';
/** Preserve the official follow protocol while exposing only this desktop workspace. */
export function scopedWorkspaceFrame(frame, cwd, known, sessionIds) {
  if (frame.type === 'baseline') {
    const items = frame.value.items.filter(item => item.path === cwd);
    known.clear(); items.forEach(item => known.add(item.workspaceId));
    return { ...frame, value: { items, archivedSessionIds: frame.value.archivedSessionIds.filter(id => sessionIds.has(id)) } };
  }
  if (frame.type === 'upsert') {
    if (frame.workspace.path === cwd) { known.add(frame.workspace.workspaceId); return frame; }
    return known.delete(frame.workspace.workspaceId) ? { type: 'remove', workspaceId: frame.workspace.workspaceId } : null;
  }
  if (frame.type === 'remove') return known.delete(frame.workspaceId) ? frame : null;
  if (frame.type === 'order') return { ...frame, workspaceIds: frame.workspaceIds.filter(id => known.has(id)) };
  if (frame.type === 'archived') return { ...frame, archivedSessionIds: frame.archivedSessionIds.filter(id => sessionIds.has(id)) };
  return null;
}
export default class LamberWorkspaceController extends WorkspaceController {
  static inject = [...WorkspaceController.inject, 'sessionController'];
  lamberWebDeployment = true;
  async create(request) {
    if (request.path !== process.cwd()) throw new Error('请通过 Lamber 主窗口切换工作区');
    return super.create(request);
  }
  async delete() { throw new Error('请通过 Lamber 主窗口管理工作区'); }
  async rename() { throw new Error('请通过 Lamber 主窗口管理工作区名称'); }
  async insertBefore() { throw new Error('此窗口只显示当前 Lamber 工作区'); }
  async insertSessionBefore(request) {
    const workspace = this.ctx.workspaceRegistry.get(request.workspaceId);
    if (workspace?.path !== process.cwd()) throw new Error('会话不属于当前工作区');
    for (const id of [request.sessionId, request.beforeSessionId].filter(Boolean)) await this.ctx.sessionController.inspect(id);
    return super.insertSessionBefore(request);
  }
  async archiveSession(request) {
    await this.ctx.sessionController.inspect(request.sessionId);
    const result = await super.archiveSession(request);
    const own = new Set((await this.ctx.sessionController.list({})).items.map(item => item.sessionId));
    return { ...result, archivedSessionIds: result.archivedSessionIds.filter(id => own.has(id)) };
  }
  async *follow(signal) {
    const known = new Set();
    for await (const frame of super.follow(signal)) {
      const own = new Set(['baseline', 'archived'].includes(frame.type) ? (await this.ctx.sessionController.list({}, signal)).items.map(item => item.sessionId) : []);
      const scoped = scopedWorkspaceFrame(frame, process.cwd(), known, own);
      if (scoped) yield scoped;
    }
  }
}
