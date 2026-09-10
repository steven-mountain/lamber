import { invoke } from '@tauri-apps/api/core';
import { emitTo, listen } from '@tauri-apps/api/event';
import { getCurrentWindow } from '@tauri-apps/api/window';
import { create } from 'zustand';
import { workspaceService } from '../utils/workspaceService';
import { projectService, type BenefitAnalysisScheme, type Project } from '../utils/projectService';
import { reverseChangesText, reverseMetricDetail, reversePercent, type ReverseMetric, type ReverseSubjectChange } from '../lib/structureReverseResult';

export interface ReverseIdentity { sessionId: string; workspaceId: string; projectId: string; projectName: string; schemeId: string }
export interface ReverseRequest extends ReverseIdentity {
  requestId: string; replyWindow: string; expiresAt: number; action: 'preview' | 'apply';
  subjectCode: string; metricType: ReverseMetric; target?: number; token?: string;
}
export interface ReversePreview {
  status: 'preview'; token: string; schemeName: string; stage: string | null; snapshotVersion: number | null;
  subjectName: string; balancingName: string; side: 'revenue' | 'cost'; totalIncl: number;
  minMetric: number; maxMetric: number;
}
export type ReverseReply = ReversePreview | { status: 'error'; message: string }
  | { status: 'success' | 'applied_warning'; message: string; target: number; achieved: number; targetReached: boolean;
      metricType: ReverseMetric; changes: ReverseSubjectChange[]; schemeName: string; stage: string | null; snapshotVersion: number | null };
const REQUEST = 'lamber-chat-structure-reverse-request';
const RESULT = 'lamber-chat-structure-reverse-result';
const previews = new Map<string, { identity: string; stamp: string; expiresAt: number }>();
const previewIdentity = (request: ReverseRequest) => JSON.stringify([request.sessionId, request.workspaceId,
  request.projectId, request.schemeId, request.subjectCode, request.metricType]);
export function issueStructurePreview(request: ReverseRequest, stamp: string) {
  for (const [token, value] of previews) if (value.expiresAt <= Date.now()) previews.delete(token);
  const token = crypto.randomUUID();
  previews.set(token, { identity: previewIdentity(request), stamp, expiresAt: Date.now() + 15 * 60_000 });
  return token;
}
export function consumeStructurePreview(request: ReverseRequest, stamp: string) {
  const preview = request.token ? previews.get(request.token) : undefined;
  if (request.token) previews.delete(request.token);
  if (!preview || preview.expiresAt <= Date.now() || preview.identity !== previewIdentity(request) || preview.stamp !== stamp) {
    throw new Error('范围预览已失效，或项目、方案、科目、指标、输入已变化，请重新读取范围。');
  }
}
export const useStructureRequest = create<{ request: ReverseRequest | null; phase: 'loading' | 'opening' | 'running' }>(() => ({ request: null, phase: 'loading' }));

export async function assertStructureBinding(target: Pick<ReverseIdentity, 'sessionId' | 'workspaceId' | 'projectId'>) {
  const binding = await invoke<{ workspaceId: string; projectId: string | null } | null>('ai_get_session_binding', { sessionId: target.sessionId });
  const workspace = await workspaceService.getState();
  if (!binding || binding.projectId !== target.projectId || binding.workspaceId !== target.workspaceId
    || workspace.currentWorkspace?.workspaceId !== target.workspaceId) throw new Error('工作区或会话绑定已变更，请重新发起结构反算。');
}
export async function loadReverseProject(sessionId: string) {
  const binding = await invoke<{ workspaceId: string; projectId: string | null; projectName?: string } | null>('ai_get_session_binding', { sessionId });
  if (!binding?.projectId) return null;
  const identity = { sessionId, workspaceId: binding.workspaceId, projectId: binding.projectId };
  await assertStructureBinding(identity);
  const [project, schemes] = await Promise.all([projectService.getProject(identity.projectId), projectService.getSchemes(identity.projectId)]);
  if (!project) throw new Error('未找到会话绑定项目。');
  await assertStructureBinding(identity);
  return { ...identity, project, schemes };
}
/** Same stage → exact id → newest exact name/default semantics as calculation.rs. No unmatched fallback. */
export function selectReverseScheme(project: Project, schemes: BenefitAnalysisScheme[], scenario: string) {
  const selector = scenario.trim();
  const newest = (rows: BenefitAnalysisScheme[]) => rows.reduce<BenefitAnalysisScheme | undefined>((best, row) =>
    !best || row.updated_at > best.updated_at || (row.updated_at === best.updated_at && row.created_at >= best.created_at) ? row : best, undefined);
  const found = !selector ? schemes.find(row => row.id === project.default_scheme_id) ?? newest(schemes)
    : ['pre_selection', 'post_selection'].includes(selector) ? newest(schemes.filter(row => row.stage === selector))
    : schemes.find(row => row.id === selector) ?? newest(schemes.filter(row => row.name === selector));
  if (!found) throw new Error(`项目「${project.name}」中没有匹配 ${selector || '默认方案'} 的测算方案；请明确选择已有方案。`);
  return found;
}

/** Dispatched only by a card interaction. Apply is never automatically retried. */
export async function requestStructureReverse(input: Omit<ReverseRequest, 'requestId' | 'replyWindow' | 'expiresAt'>): Promise<ReverseReply> {
  await assertStructureBinding(input);
  const request: ReverseRequest = { ...input, requestId: crypto.randomUUID(), replyWindow: getCurrentWindow().label, expiresAt: Date.now() + 15000 };
  return new Promise((resolve, reject) => {
    let stop: (() => void) | undefined;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const finish = (result: ReverseReply) => { stop?.(); clearTimeout(timer); resolve(result); };
    void listen<{ requestId: string; result?: ReverseReply }>(RESULT, event => {
      if (event.payload.requestId !== request.requestId) return;
      clearTimeout(timer);
      if (event.payload.result) finish(event.payload.result);
    }).then(unlisten => {
      stop = unlisten;
      timer = setTimeout(() => finish({ status: 'error', message: '主窗口未接收请求，请打开主窗口后重试。' }), 15000);
      return emitTo('main', REQUEST, request);
    }).catch(error => { stop?.(); clearTimeout(timer); reject(error); });
  });
}
export async function finishStructureRequest(request: ReverseRequest, result: ReverseReply) {
  if (useStructureRequest.getState().request?.requestId === request.requestId) useStructureRequest.setState({ request: null, phase: 'loading' });
  await emitTo(request.replyWindow, RESULT, { requestId: request.requestId, result });
}
export async function listenStructureRequests(open: (request: ReverseRequest) => Promise<void>) {
  const seen = new Set<string>();
  return listen<ReverseRequest>(REQUEST, event => {
    const request = event.payload;
    if (request.expiresAt <= Date.now() || seen.has(request.requestId)) return;
    seen.add(request.requestId);
    if (seen.size > 1000) seen.delete(seen.values().next().value!);
    if (useStructureRequest.getState().request) {
      void emitTo(request.replyWindow, RESULT, { requestId: request.requestId, result: { status: 'error', message: '已有结构反算操作正在进行，请等待其结果。' } });
      return;
    }
    useStructureRequest.setState({ request, phase: 'loading' });
    void (async () => {
      await emitTo(request.replyWindow, RESULT, { requestId: request.requestId });
      await assertStructureBinding(request);
      const schemes = await projectService.getSchemes(request.projectId);
      if (!schemes.some(scheme => scheme.id === request.schemeId)) throw new Error('所选方案不存在或不属于绑定项目，已停止结构反算。');
      await open(request);
      if (useStructureRequest.getState().request?.requestId === request.requestId) useStructureRequest.setState({ phase: 'opening' });
    })().catch(error => finishStructureRequest(request, { status: 'error', message: error instanceof Error ? error.message : String(error) }).catch(console.error));
  });
}
export function structureReceipt(projectName: string, result: ReverseReply) {
  if (result.status === 'error') return `结构反算未执行成功：${projectName}\n${result.message}`;
  const scheme = `${projectName} · ${result.schemeName} · ${result.stage || '未标注阶段'} · 原快照${result.snapshotVersion === null ? '无' : `v${result.snapshotVersion}`}`;
  if (result.status === 'preview') return `结构反算范围预览（未写金额）：${scheme}\n目标科目：${result.subjectName}；差额承接：${result.balancingName}；锁定总额：${result.totalIncl.toFixed(2)}元。\n当前结构下可达范围约为 ${reversePercent(result.minMetric)} - ${reversePercent(result.maxMetric)}（采样范围不代表其中每个值都能达到）。请用户自行确认目标值。`;
  return `${result.status === 'success' ? '结构反算已写入当前编辑器，尚未保存' : '结构反算已修改金额，请核对'}：${scheme}\n\n${result.metricType === 'margin' ? '毛利润率' : '净现值率'}\n\n| 目标值 | 实际达成值 |\n| --- | --- |\n| ${reverseMetricDetail(result.target)} | ${reverseMetricDetail(result.achieved)} |\n\n${result.message}\n\n逐科目含税金额与科目收付款计划（变更前 → 变更后）：\n${reverseChangesText(result.changes) || '科目金额和计划均未变化。'}`;
}
