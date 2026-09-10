import { TypertGatewayService } from '@deepseek-ai/dsh-api-gateway';
import { RemoteError } from '@deepseek-ai/dsh-typert-protocol';
const restricted = new Set(['fileReferences', 'directoryPicker', 'credentials', 'cordis', 'plugins', 'goals', 'jobs', 'skills', 'subagents', 'workflows', 'sessionReferenceResolver']);
export function permitted({ namespace, method, args = {} }) {
  if (restricted.has(namespace)) return false;
  if (namespace === 'settings') return method === 'describe' || (['update', 'replace', 'mutate'].includes(method) && ['ui-theme', 'locale'].includes(args.ns));
  if (namespace === 'agentPresets') return ['list', 'read'].includes(method);
  if (namespace === 'commands' && method === 'execute') return /^\/compact(?:\s|$)/.test(args.line ?? '');
  if (['permissionPresets', 'sandbox', 'approvalPolicy'].includes(namespace)) return !['select', 'set', 'update'].includes(method);
  return true;
}
/** Deployment authorization over public invocation seams; upstream owns decoding and transport. */
export default class LamberGateway extends TypertGatewayService {
  static inject = [...TypertGatewayService.inject, 'sessionController'];
  lamberWebDeployment = true;
  async authorizeSession(request) {
    const args = request.args?.request ?? request.args;
    const id = args?.sessionId ?? args?.agentId ?? (typeof args?.agent === 'string' ? args.agent : undefined);
    if (id && ['session', 'commands', 'approval', 'userQuestions'].includes(request.namespace)) await this.ctx.sessionController.inspect(id, request.signal);
  }
  async invoke(request) {
    if (!permitted(request)) throw new RemoteError('gateway/bad-request', '此功能由 Lamber 主窗口和业务权限管理。', {});
    await this.authorizeSession(request);
    const value = await super.invoke(request);
    if (request.namespace === 'settings' && request.method === 'describe') {
      return { ...value, hasDocument: false, namespaces: value.namespaces.filter(row => ['ui-theme', 'locale'].includes(row.ns)) };
    }
    if (request.namespace === 'commands' && request.method === 'list') return value.filter(command => command.name === 'compact');
    return value;
  }
  async stream(request) {
    if (!permitted(request)) throw new RemoteError('gateway/bad-request', '此功能由 Lamber 主窗口和业务权限管理。', {});
    await this.authorizeSession(request);
    return super.stream(request);
  }
}
