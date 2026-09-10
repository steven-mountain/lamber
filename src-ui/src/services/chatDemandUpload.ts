import { assertDemandUploadBinding, type DemandUploadTarget } from './demandUploadTargets';
import { domainSaveService } from './domainSaveService';
import { publishDemandAssetsChanged } from './demandTemplateAssets';
import { projectService } from '../utils/projectService';
import { webBusinessTransport } from './webBusinessTransport';
export interface DemandUpload { name: string; dataUrl: string; width: number; height: number }
/** Called only after the existing card's explicit file choice or paste. */
export async function saveChatDemandUpload(target: DemandUploadTarget, image: DemandUpload, isActive: () => boolean) {
  if (!isActive()) throw new Error('图片卡片已失效，请重新选择。');
  const remote = webBusinessTransport();
  if (remote) return remote<{ assetId: string; path: string; warning: string }>('upload-demand', { sessionId: target.sessionId, target, image });
  await assertDemandUploadBinding(target);
  if (!isActive()) throw new Error('工作区或会话已切换，请重新选择图片。');
  const assetId = await domainSaveService.saveTemplateAsset(target.projectId, target.templateName, {
    assetType: 'image', usage: target.usage, originalFileName: image.name,
    base64Data: image.dataUrl, width: image.width, height: image.height,
  });
  let warning = '';
  try { await publishDemandAssetsChanged(target); } catch { warning = '已入库，请重新打开模板查看。'; }
  const path = await projectService.getTemplateAssetPath(assetId).catch(() => '图片已入库，请在模板页查看。');
  return { assetId, path, warning };
}
