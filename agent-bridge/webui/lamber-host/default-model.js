import { Service } from '@deepseek-ai/cordis';
import { postBridge } from 'dsh-tool-lamber/lib/bridge.js';
// Official SessionController persists model selection through this deployment-owned service.
// Lamber config is the only writable default, while each live agent keeps its selected model.
export default class LamberDefaultModel extends Service {
  lamberWebDeployment = true;
  constructor(ctx, config) {
    super(ctx, 'agentDefaultModel');
    this.selection = { provider: 'deepseek-official', model: config.model };
  }
  currentSelection() { return { ...this.selection }; }
  async saveSelection(next) {
    if (next.provider !== 'deepseek-official') throw new Error('此部署尚未支持该服务商');
    await postBridge('/lamber-webui/select-model', { model: next.model }, AbortSignal.timeout(10000));
    this.selection = { ...next };
  }
}
