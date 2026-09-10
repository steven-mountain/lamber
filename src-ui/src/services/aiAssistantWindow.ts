import { AI_SESSION_STORAGE_KEY } from '../ai/sessionTypes';
import { invoke } from '@tauri-apps/api/core';
import { emit, emitTo } from '@tauri-apps/api/event';
import { WebviewWindow } from '@tauri-apps/api/webviewWindow';
import { availableMonitors, currentMonitor, PhysicalPosition, PhysicalSize } from '@tauri-apps/api/window';
import { AI_CONTEXT_REFRESH_REQUEST_EVENT } from '../store/useAiContextStore';
import { AI_WINDOW_POSITION_KEY, parseWindowPosition, placeAiWindow } from '../lib/aiWindowPlacement';

const LABEL = 'ai-assistant';
let opening: Promise<void> | null = null;
let trackedWindow: WebviewWindow | null = null;

async function openWindow() {
  let target = await WebviewWindow.getByLabel(LABEL);
  const created = !target;
  const legacy = localStorage.getItem(AI_SESSION_STORAGE_KEY);
  if (legacy) await invoke('ai_webui_import_history', { raw: legacy });
  await invoke('ai_open_webui');
  target = await WebviewWindow.getByLabel(LABEL);
  if (!target) throw new Error('AI 窗口未创建，请重试。');
  if (trackedWindow !== target && created) {
    trackedWindow = target;
    const currentTarget = target;
    const stopMoved = await target.onMoved(({ payload }) => {
      try { localStorage.setItem(AI_WINDOW_POSITION_KEY, JSON.stringify({ version: 2, x: payload.x, y: payload.y })); }
      catch { /* Geometry preferences never block native events. */ }
    });
    const stopDestroyed = await target.once('tauri://destroyed', () => {
      stopMoved(); stopDestroyed();
      if (trackedWindow === currentTarget) trackedWindow = null;
    });
  }
  await target.unminimize();
  const [screens, preferred, size, current, scale] = await Promise.all([
    availableMonitors(), currentMonitor(), target.outerSize(), target.outerPosition(), target.scaleFactor(),
  ]);
  let saved = null;
  try { saved = created ? parseWindowPosition(localStorage.getItem(AI_WINDOW_POSITION_KEY)) : null; }
  catch { /* Position preferences cannot prevent opening. */ }
  const candidate = saved
    ? { x: saved.x * (saved.version === 2 ? 1 : scale), y: saved.y * (saved.version === 2 ? 1 : scale) }
    : current;
  const placement = placeAiWindow(candidate, size, screens, preferred);
  if (placement.size.width !== size.width || placement.size.height !== size.height) {
    await target.setSize(new PhysicalSize(placement.size.width, placement.size.height));
  }
  await target.setPosition(new PhysicalPosition(placement.position.x, placement.position.y));
  await target.show();
  await target.setFocus();
  try { localStorage.setItem(AI_WINDOW_POSITION_KEY, JSON.stringify({ version: 2, ...placement.position })); }
  catch { /* Best-effort UI preference only. */ }
  const view = localStorage.getItem('lamber_ai_current_view') || 'hub';
  await Promise.all([
    emitTo(LABEL, 'lamber-ai-view-changed', { view }),
    emit(AI_CONTEXT_REFRESH_REQUEST_EVENT, { view }),
  ]);
}

/** Serialize creation until the native window is ready; repeated clicks reuse it. */
export function openAiAssistantWindow(view: string): Promise<void> {
  localStorage.setItem('lamber_ai_current_view', view);
  if (!('__TAURI_INTERNALS__' in window)) {
    return Promise.reject(new Error('AI 工作区需要桌面应用，请通过 npm run tauri dev 启动。'));
  }
  if (!opening) opening = openWindow().finally(() => { opening = null; });
  return opening;
}
