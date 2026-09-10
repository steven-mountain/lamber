import { invoke } from '@tauri-apps/api/core';
import { getCatalogCompletion, getCatalogTemplate } from '../lib/templateCompletion/catalog';
import { domainSaveService } from './domainSaveService';
import { workspaceService } from '../utils/workspaceService';
import { projectService } from '../utils/projectService';
import type { DemandAssetTarget } from './demandTemplateAssets';

export interface DemandUploadTarget extends DemandAssetTarget {
  sessionId: string; projectName: string; key: string; label: string; usage: 'attach1' | 'attach2';
}
interface Binding { workspaceId: string; projectId: string | null; projectName?: string | null }

export async function assertDemandUploadBinding(target: DemandUploadTarget) {
  const binding = await invoke<Binding | null>('ai_get_session_binding', { sessionId: target.sessionId });
  const workspace = await workspaceService.getState();
  if (!binding || binding.projectId !== target.projectId || binding.workspaceId !== target.workspaceId
    || workspace.currentWorkspace?.workspaceId !== target.workspaceId) {
    throw new Error('工作区或会话绑定已变更，请重新选择图片。');
  }
}

/** Product UI read channel. Never attach these states to a model prompt. */
export async function loadDemandUploadTargets(sessionId: string): Promise<DemandUploadTarget[]> {
  const binding = await invoke<Binding | null>('ai_get_session_binding', { sessionId });
  if (!binding?.projectId) return [];
  const workspace = await workspaceService.getState();
  if (workspace.currentWorkspace?.workspaceId !== binding.workspaceId) throw new Error('请打开会话绑定的工作区。');
  const projectId = binding.projectId;
  const saved = await domainSaveService.loadTemplateStates(projectId);
  const isDemand = (name: string) => getCatalogTemplate(name)?.id === 'demand';
  let names = saved.map(state => state.templateId).filter(isDemand);
  if (!names.length) {
    names = (await invoke<string[]>('get_available_templates', { moduleId: 'ict_lifecycle' })).filter(isDemand);
    if (names.length > 1) throw new Error('找到多张需求导入表，请先在模板页选择并保存目标模板，再回聊天补图。');
  }
  const targets: DemandUploadTarget[] = [];
  for (const templateName of [...new Set(names)]) {
    const state = saved.find(item => item.templateId === templateName)
      ?? await domainSaveService.loadTemplateState(projectId, templateName);
    const assets = await domainSaveService.loadTemplateAssets(projectId, templateName);
    const existence = await Promise.all(assets.map(async asset => ({
      fieldKey: asset.usage,
      exists: await projectService.getTemplateAssetPath(asset.id).then(() => true, () => false),
    })));
    for (const item of getCatalogCompletion(templateName, state?.filledDataJson ?? {}, existence)) {
      if (item.kind === 'image' && item.evaluated && !item.filled && item.usage) {
        targets.push({ sessionId, workspaceId: binding.workspaceId, projectId,
          projectName: binding.projectName || projectId, templateName: item.templateName,
          key: item.key, label: item.label, usage: item.usage });
      }
    }
  }
  // Reads may span a workspace switch or a session reset. Never expose a stale target.
  const current = await invoke<Binding | null>('ai_get_session_binding', { sessionId });
  const after = await workspaceService.getState();
  if (current?.projectId !== projectId || current?.workspaceId !== binding.workspaceId
    || after.currentWorkspace?.workspaceId !== binding.workspaceId) throw new Error('工作区或会话已切换，请刷新附件卡片。');
  return targets;
}
