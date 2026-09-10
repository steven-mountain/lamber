import { buildAiChatContext } from '../ai/context/buildAiChatContext';
import { invoke } from '@tauri-apps/api/core';
import { listen } from '@tauri-apps/api/event';
import { documentReceipt, loadDocumentTargets, requestTemplateAction } from './chatDocumentGeneration';
import { loadListTargets, readQuoteImage } from './chatTemplateLists';
import { loadReverseProject, requestStructureReverse, structureReceipt } from './chatStructureReverse';
import { listChatTemplateImages, readChatTemplateImage, replaceChatTemplateImage } from './chatTemplateImages';
import { loadDemandUploadTargets } from './demandUploadTargets';
import { saveChatDemandUpload } from './chatDemandUpload';

interface Request { requestId: string; method: string; payload: Record<string, unknown>; expiresAt: number }
/** Main window reuses the same generators, normalization and editor request listeners. */
async function execute({ method, payload: p }: Request): Promise<unknown> {
  const sessionId = String(p.sessionId);
  switch (method) {
    case 'context': {
      const binding = await invoke<{ workspaceId: string; projectId: string | null } | null>('ai_get_session_binding', { sessionId });
      if (!binding) throw new Error('当前会话没有可信绑定，无法读取业务上下文。');
      const context = await buildAiChatContext({ currentView: localStorage.getItem('lamber_ai_current_view') || 'hub', userMessage: String(p.userMessage || ''), boundProjectId: binding.projectId });
      const after = await invoke<{ workspaceId: string; projectId: string | null } | null>('ai_get_session_binding', { sessionId });
      if (!after || after.workspaceId !== binding.workspaceId || after.projectId !== binding.projectId) throw new Error('读取期间工作区或绑定已变更。');
      return context.contextNodes;
    }
    case 'documents': return loadDocumentTargets(sessionId, p.templateIds as string[]);
    case 'lists': return loadListTargets(sessionId, p.templateIds as string[]);
    case 'reverse-project': return loadReverseProject(sessionId);
    case 'images': return listChatTemplateImages(sessionId);
    case 'read-image': return readChatTemplateImage(p.target as Parameters<typeof readChatTemplateImage>[0]);
    case 'quote-image': return readQuoteImage(p.target as Parameters<typeof readQuoteImage>[0], String(p.assetId));
    case 'demand-targets': return loadDemandUploadTargets(sessionId);
    case 'template-action': return requestTemplateAction(p.target as Parameters<typeof requestTemplateAction>[0], p.action as Parameters<typeof requestTemplateAction>[1]);
    case 'reverse-action': return requestStructureReverse(p.input as Parameters<typeof requestStructureReverse>[0]);
    case 'replace-image': return replaceChatTemplateImage(p.target as Parameters<typeof replaceChatTemplateImage>[0], p.next as Parameters<typeof replaceChatTemplateImage>[1], () => true);
    case 'upload-demand': return saveChatDemandUpload(p.target as Parameters<typeof saveChatDemandUpload>[0], p.image as Parameters<typeof saveChatDemandUpload>[1], () => true);
    default: throw new Error('未开放此业务操作');
  }
}
function receiptSummary(request: Request, value: unknown): string | undefined {
  const p = request.payload;
  if (request.method === 'reverse-action') return structureReceipt(String((p.input as { projectName: string }).projectName), value as Parameters<typeof structureReceipt>[1]);
  if (request.method === 'template-action') {
    const target = p.target as Parameters<typeof documentReceipt>[0];
    const result = value as Awaited<ReturnType<typeof requestTemplateAction>>;
    const type = (p.action as { type?: string } | undefined)?.type;
    if (result.status !== 'list') {
      if (type === 'generate' || result.status === 'success') return documentReceipt(target, result);
      return `${type === 'saveTech' ? '技术清单' : '询价'}操作未完成：${target.projectName} · ${target.templateName}\n${result.message}`;
    }
    if (type === 'saveTech') return `技术清单已由用户保存：${target.projectName}，共 ${result.snapshot.techItems.length} 行，同时更新需求导入表与会审纪要。`;
    if (type === 'generateInquiry') return `询价清单已由用户确认生成：${target.projectName}，共 ${result.snapshot.inqVendors.length} 家报价。具体报价及证据请在卡片核对。`;
    if (type === 'saveInquiry') return `询价清单已由用户保存：${target.projectName}，共 ${result.snapshot.inqVendors.length} 家报价。`;
    return `已读取清单：${target.projectName} · ${target.templateName}`;
  }
  if (request.method === 'replace-image' || request.method === 'upload-demand') {
    const target = p.target as { projectName: string; templateName: string; name?: string };
    const result = value as { refreshWarning?: string; warning?: string };
    return `图片已由用户${request.method === 'replace-image' ? '确认替换' : '上传保存'}：${target.projectName} · ${target.templateName}\n${result.refreshWarning || result.warning || ''}`;
  }
}
export function listenWebBusinessRequests() {
  return listen<Request>('lamber-webui-business-request', event => {
    void (async () => {
      const request = event.payload;
      // Claim in Rust before any side effect. Duplicate native events never execute twice.
      try { await invoke('ai_webui_claim_action', { requestId: request.requestId }); }
      catch { return; }
      let result: { ok: boolean; value?: unknown; error?: string; summary?: string };
      try {
        if (request.expiresAt <= Date.now()) throw new Error('主窗口未及时接收，操作已取消。');
        const value = await execute(request);
        result = { ok: true, value, summary: receiptSummary(request, value) };
      } catch (error) { result = { ok: false, error: String(error) }; }
      // A missing acknowledgement never asks the user to repeat an already committed operation.
      await invoke('ai_webui_complete_action', { requestId: request.requestId, result });
    })().catch(console.error);
  });
}
