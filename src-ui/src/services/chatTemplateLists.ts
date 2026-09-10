import { loadDocumentTargets, assertDocumentBinding, requestTemplateAction, type DocumentTarget } from './chatDocumentGeneration';
import { domainSaveService } from './domainSaveService';
import { listSnapshot, type InquiryVendor, type TechItem, type TemplateListAction } from './templateListTypes';
export interface ListTarget extends DocumentTarget { snapshot: ReturnType<typeof listSnapshot> }
export async function loadListTargets(sessionId: string, templateIds: string[]): Promise<ListTarget[]> {
  const targets = await loadDocumentTargets(sessionId,templateIds);
  const result = await Promise.all(targets.map(async target => {
    const saved = await domainSaveService.loadTemplateState(target.projectId,target.templateName);
    const state = saved?.filledDataJson ?? {};
    return {...target,snapshot:listSnapshot(Array.isArray(state.techItems) ? state.techItems as TechItem[] : [], Array.isArray(state.inqVendors) ? state.inqVendors as InquiryVendor[] : [])};
  }));
  if (targets[0]) await assertDocumentBinding(targets[0]);
  return result;
}
export async function runListAction(target: ListTarget, action: TemplateListAction) {
  const result = await requestTemplateAction(target,action);
  if (result.status === 'list') return {...target,snapshot:result.snapshot};
  if (result.status === 'success') throw new Error('清单请求收到文档回执，请重新核对');
  throw new Error(result.message);
}
