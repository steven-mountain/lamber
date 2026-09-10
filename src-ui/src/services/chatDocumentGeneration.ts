import type { TemplateListAction, ListSnapshot } from './templateListTypes';
import { invoke } from '@tauri-apps/api/core';
import { emitTo, listen } from '@tauri-apps/api/event';
import { getCurrentWindow } from '@tauri-apps/api/window';
import { create } from 'zustand';
import { workspaceService } from '../utils/workspaceService';
import { domainSaveService } from './domainSaveService';
import { projectService } from '../utils/projectService';
import { getCatalogCompletion, getCatalogTemplate, type CatalogCompletionItem } from '../lib/templateCompletion/catalog';

export interface DocumentTarget {
  sessionId: string; workspaceId: string; projectId: string; projectName: string; templateName: string;
  completion: CatalogCompletionItem[];
}
export interface DocumentRequest extends Omit<DocumentTarget, 'completion'> {
  action?: TemplateListAction; requestId: string; replyWindow: string; expiresAt: number;
}
export type GenerationResult = { status: 'success'; outputDir: string } | { status: 'error' | 'cancelled'; message: string };
export type TemplateActionResult = GenerationResult | { status: 'list'; snapshot: ListSnapshot };
const REQUEST = 'lamber-chat-document-request';
const RESULT = 'lamber-chat-document-result';
export const useDocumentRequest = create<{ request: DocumentRequest | null; phase: 'loading' | 'opening' | 'generating' | 'running' }>(() => ({ request: null, phase: 'opening' }));

export async function assertDocumentBinding(target: Pick<DocumentTarget, 'sessionId' | 'workspaceId' | 'projectId'>) {
  const binding = await invoke<{workspaceId: string; projectId: string | null} | null>('ai_get_session_binding', { sessionId: target.sessionId });
  const workspace = await workspaceService.getState();
  if (!binding || binding.projectId !== target.projectId || binding.workspaceId !== target.workspaceId
    || workspace.currentWorkspace?.workspaceId !== target.workspaceId) throw new Error('工作区或会话绑定已变更，请重新发起生成。');
}

/** Reads only the trusted bound project, without page context or a model invocation. */
export async function loadDocumentTargets(sessionId: string, templateIds: readonly string[]): Promise<DocumentTarget[]> {
  const binding = await invoke<{workspaceId: string; projectId: string | null; projectName?: string} | null>('ai_get_session_binding', { sessionId });
  if (!binding?.projectId) return [];
  const base = { sessionId, workspaceId: binding.workspaceId, projectId: binding.projectId, projectName: binding.projectName || binding.projectId };
  await assertDocumentBinding(base);
  const available = await invoke<string[]>('get_available_templates', { moduleId: 'ict_lifecycle' });
  const names = available.filter(name => templateIds.includes(getCatalogTemplate(name)?.id ?? ''));
  const targets = await Promise.all(names.map(async templateName => {
    const state = await domainSaveService.loadTemplateState(base.projectId, templateName);
    const assets = await domainSaveService.loadTemplateAssets(base.projectId, templateName);
    const existence = await Promise.all(assets.map(async asset => ({ fieldKey: asset.usage,
      exists: await projectService.getTemplateAssetPath(asset.id).then(() => true, () => false) })));
    return { ...base, templateName, completion: getCatalogCompletion(templateName, state?.filledDataJson ?? {}, existence) };
  }));
  await assertDocumentBinding(base);
  if (!targets.length) throw new Error('当前模板目录中没有找到目标模板，请在设置中检查模板目录。');
  return targets;
}

export function completionSummary(items: readonly CatalogCompletionItem[]) {
  return { totalCount: items.length, filledCount: items.filter(item => item.evaluated && item.filled).length,
    missing: items.filter(item => item.evaluated && !item.filled), unknown: items.filter(item => !item.evaluated) };
}
export function documentReceipt(target: Pick<DocumentTarget, 'projectName' | 'templateName'>, result: GenerationResult) {
  const label = `${target.projectName} · ${target.templateName}`;
  return result.status === 'success' ? `文档已生成：${label}\n实际输出目录：${result.outputDir}`
    : `${result.status === 'cancelled' ? '已取消生成' : '文档未生成'}：${label}\n${result.message}`;
}

/** Only called by a user click. Subscribe before dispatch; never retry a generation implicitly. */
export async function requestDocumentGeneration(target: DocumentTarget): Promise<GenerationResult> {
  const result = await requestTemplateAction(target);
  if (result.status === 'list') throw new Error('生成请求收到错误类型的回执');
  return result;
}
export async function requestTemplateAction(target: DocumentTarget, action?: TemplateListAction): Promise<TemplateActionResult> {
  await assertDocumentBinding(target);
  const { completion: _, ...identity } = target;
  const request: DocumentRequest = { ...identity, action, requestId: crypto.randomUUID(), replyWindow: getCurrentWindow().label, expiresAt: Date.now() + 15_000 };
  return new Promise((resolve, reject) => {
    let stop: (() => void) | undefined;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const finish = (result: TemplateActionResult) => { stop?.(); clearTimeout(timer); resolve(result); };
    void listen<{requestId: string; result?: TemplateActionResult}>(RESULT, event => {
      if (event.payload.requestId !== request.requestId) return;
      clearTimeout(timer); // Accepted requests have no generation timeout; a dialog may await the user.
      if (event.payload.result) finish(event.payload.result);
    }).then(unlisten => {
      stop = unlisten;
      timer = setTimeout(() => finish({ status: 'error', message: '主窗口未接收生成请求，请打开主窗口后重试。' }), 15_000);
      return emitTo('main', REQUEST, request);
    }).catch(error => { stop?.(); clearTimeout(timer); reject(error); });
  });
}
export async function finishDocumentRequest(request: DocumentRequest, result: TemplateActionResult) {
  if (useDocumentRequest.getState().request?.requestId === request.requestId) useDocumentRequest.setState({ request: null, phase: 'opening' });
  await emitTo(request.replyWindow, RESULT, { requestId: request.requestId, result });
}
export async function listenDocumentRequests(open: (request: DocumentRequest) => Promise<void>) {
  return listen<DocumentRequest>(REQUEST, event => {
    const request = event.payload;
    // Expired events cannot generate a file after the sender reports a delivery failure.
    if (request.expiresAt <= Date.now()) return;
    if (useDocumentRequest.getState().request) {
      void emitTo(request.replyWindow, RESULT, { requestId: request.requestId, result: { status: 'error', message: '已有文档正在生成，请完成主窗口中的操作后重试。' } });
      return;
    }
    useDocumentRequest.setState({ request, phase: 'loading' });
    void (async () => {
      await emitTo(request.replyWindow, RESULT, { requestId: request.requestId });
      await assertDocumentBinding(request);
      await open(request);
      if (useDocumentRequest.getState().request?.requestId === request.requestId) useDocumentRequest.setState({ phase: 'opening' });
    })().catch(error => finishDocumentRequest(request, { status: 'error', message: String(error) }).catch(console.error));
  });
}
