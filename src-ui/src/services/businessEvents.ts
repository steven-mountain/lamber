import { listen as nativeListen, type EventCallback, type UnlistenFn } from '@tauri-apps/api/event';
import { webBusinessTransport } from './webBusinessTransport';
const listeners = new Map<string, Set<(payload: unknown) => void>>();
export function publishWebBusinessEvent(name: string, payload: unknown = {}) {
  for (const callback of listeners.get(name) ?? []) callback(payload);
}
export function listenBusinessEvent<T>(name: string, callback: EventCallback<T>): Promise<UnlistenFn> {
  if (!webBusinessTransport()) return nativeListen(name, callback);
  const listener = (payload: unknown) => callback({ event: name, id: 0, payload: payload as T });
  const group = listeners.get(name) ?? new Set(); listeners.set(name, group); group.add(listener);
  return Promise.resolve(() => { group.delete(listener); if (!group.size) listeners.delete(name); });
}
