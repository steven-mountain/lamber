import { emit, emitTo } from '@tauri-apps/api/event';
import { WebviewWindow } from '@tauri-apps/api/webviewWindow';
import { availableMonitors, currentMonitor, PhysicalPosition, PhysicalSize } from '@tauri-apps/api/window';
import { AI_CONTEXT_REFRESH_REQUEST_EVENT } from '../store/useAiContextStore';
import { AI_WINDOW_POSITION_KEY, parseWindowPosition, placeAiWindow } from '../lib/aiWindowPlacement';

const LABEL = 'ai-assistant';
let opening: Promise<void> | null = null;

function waitForCreation(target: WebviewWindow): Promise<void> {
  const listeners: Promise<() => void>[] = [];
  return new Promise<void>((resolve, reject) => {
    listeners.push(
      target.once('tauri://created', () => resolve()),
      target.once('tauri://error', event => reject(new Error(String(event.payload)))),
    );
    // Registration failures must also reach the launcher, not become unhandled rejections.
    listeners.forEach(listener => { void listener.catch(reject); });
  }).finally(() => {
    listeners.forEach(listener => { void listener.then(stop => stop(), () => {}); });
  });
}

async function openWindow() {
  let target = await WebviewWindow.getByLabel(LABEL);
  const created = !target;
  if (!target) {
    target = new WebviewWindow(LABEL, {
      url: `/#/ai-assistant?view=${encodeURIComponent(localStorage.getItem('lamber_ai_current_view') || 'hub')}`,
      title: 'Lamber AI 助手', width: 780, height: 680, minWidth: 360, minHeight: 480,
      decorations: false, transparent: true, backgroundColor: [0, 0, 0, 0],
      alwaysOnTop: true, resizable: true, shadow: false, skipTaskbar: false,
      center: true, visible: false, preventOverflow: true,
    });
    await waitForCreation(target);
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
    window.location.hash = `#/ai-assistant?view=${encodeURIComponent(view)}`;
    return Promise.resolve();
  }
  if (!opening) opening = openWindow().finally(() => { opening = null; });
  return opening;
}
