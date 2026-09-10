import { invoke } from '@tauri-apps/api/core';
import { domainSaveService } from './domainSaveService';
import { workspaceService } from '../utils/workspaceService';
import { getCatalogTemplate } from '../lib/templateCompletion/catalog';
import { loadAiTemplateAsset } from './aiProjectContextService';
import { publishDemandAssetsChanged } from './demandTemplateAssets';

export interface ImageTarget {
  sessionId: string; workspaceId: string; projectId: string; projectName: string;
  templateName: string; assetId: string; usage: string; name: string;
  width?: number; height?: number;
}
export interface ReplacementImage { name: string; dataUrl: string; width: number; height: number; size: number; mimeType: string }
interface Binding { workspaceId: string; projectId: string | null; projectName?: string }

export async function assertImageTarget(target: Pick<ImageTarget, 'sessionId' | 'workspaceId' | 'projectId'>) {
  const binding = await invoke<Binding | null>('ai_get_session_binding', { sessionId: target.sessionId });
  const ws = await workspaceService.getState();
  if (!binding?.projectId || binding.projectId !== target.projectId || binding.workspaceId !== target.workspaceId
    || ws.currentWorkspace?.workspaceId !== target.workspaceId) throw new Error('会话或工作区已切换，请重新打开项目图片。');
}

export async function listChatTemplateImages(sessionId: string): Promise<ImageTarget[]> {
  const binding = await invoke<Binding | null>('ai_get_session_binding', { sessionId });
  if (!binding?.projectId) throw new Error('请使用绑定项目的会话。');
  const target = { sessionId, workspaceId: binding.workspaceId, projectId: binding.projectId };
  await assertImageTarget(target);
  // Assets can exist before the template's first form-state save.
  const assets = await domainSaveService.loadTemplateAssets(binding.projectId);
  const result: ImageTarget[] = assets.filter(a => getCatalogTemplate(a.templateId || a.templateName)?.id === 'demand'
    && a.assetType === 'image' && ['attach1', 'attach2'].includes(a.usage))
    .slice().reverse().map(a => ({ ...target, projectName: binding.projectName || binding.projectId!,
      templateName: a.templateId || a.templateName, assetId: a.id, usage: a.usage,
      name: a.originalFileName || a.id, width: a.width, height: a.height }));
  await assertImageTarget(target);
  return result;
}

export async function readChatTemplateImage(target: ImageTarget) {
  await assertImageTarget(target);
  const image = await loadAiTemplateAsset(target.projectId, target.assetId);
  await assertImageTarget(target);
  return image;
}

export async function prepareReplacement(file: File): Promise<ReplacementImage> {
  if (!['image/png', 'image/jpeg', 'image/webp'].includes(file.type)) throw new Error('仅支持 PNG、JPEG、WEBP 图片。');
  if (!file.size || file.size > 20 * 1024 * 1024) throw new Error('图片须非空且不超过20MB。');
  const dataUrl = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader(); reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error('图片读取失败')); reader.readAsDataURL(file);
  });
  const size = await new Promise<{ width: number; height: number }>((resolve, reject) => {
    const image = new Image(); image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight });
    image.onerror = () => reject(new Error('图片无法解码')); image.src = dataUrl;
  });
  return { name: file.name, dataUrl, ...size, size: file.size, mimeType: file.type };
}

export async function replaceChatTemplateImage(target: ImageTarget, next: ReplacementImage, isActive: () => boolean) {
  await assertImageTarget(target);
  if (!isActive()) throw new Error('当前图片卡片已失效，请重新选择。');
  const assetId = await invoke<string>('ai_replace_template_image', { request: {
    sessionId: target.sessionId, workspaceId: target.workspaceId, projectId: target.projectId,
    templateName: target.templateName, assetId: target.assetId,
    fileName: next.name, dataUrl: next.dataUrl, width: next.width, height: next.height,
  } });
  // The transaction is committed: notification failure must not suggest retrying the write.
  let refreshWarning = '';
  try { await publishDemandAssetsChanged(target); } catch { refreshWarning = '已保存；页面刷新通知失败，请重新打开模板查看。'; }
  return { assetId, refreshWarning };
}
