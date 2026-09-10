import { BRIDGE_CONTRACT } from './contract.generated.js';
import { BRIDGE_URL_ENV, BRIDGE_TOKEN_ENV, BRIDGE_TOKEN_HEADER_ENV } from './bridge.js';

export const HANDSHAKE_ROUTE = '/lamber-bridge/handshake';
export const STARTUP_PREFIX = 'LAMBER_BRIDGE_STARTUP:';
export const MISMATCH_MESSAGE = 'AI 组件版本不匹配，无法启动。请完整重新构建 Lamber 和 AI 组件，或重新安装最新版本。';
export const UNREACHABLE_MESSAGE = '无法连接 AI 本地服务，无法启动。请关闭并重新打开 Lamber；若仍失败，请重新安装。';

/** Runs once before any tools or hooks are registered. Never retry or degrade. */
export async function handshakeBridge(): Promise<void> {
  const origin = process.env[BRIDGE_URL_ENV]?.trim();
  const token = process.env[BRIDGE_TOKEN_ENV]?.trim();
  let response: Response;
  try {
    if (!origin || !token) throw new Error('missing launch credentials');
    response = await fetch(`${origin.replace(/\/+$/, '')}${HANDSHAKE_ROUTE}`, {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        [process.env[BRIDGE_TOKEN_HEADER_ENV]?.trim() || 'x-lamber-bridge-token']: token,
      },
      body: JSON.stringify(BRIDGE_CONTRACT),
      signal: AbortSignal.timeout(5000),
    });
  } catch {
    throw new Error(UNREACHABLE_MESSAGE);
  }
  // An old backend has no handshake endpoint. Its normal unknown-route behavior
  // stays intact; only this startup request translates it to a version error.
  if (response.status === 404 || response.status === 409) throw new Error(MISMATCH_MESSAGE);
  if (!response.ok) throw new Error(UNREACHABLE_MESSAGE);
  let reply: unknown;
  try { reply = await response.json(); } catch { throw new Error(MISMATCH_MESSAGE); }
  const result = reply as { version?: unknown; routes?: unknown } | null;
  const routes = result?.routes;
  if (!result || result.version !== BRIDGE_CONTRACT.version || !Array.isArray(routes)
      || !BRIDGE_CONTRACT.routes.every(route => routes.includes(route))) {
    throw new Error(MISMATCH_MESSAGE);
  }
}

/** Explicit lifecycle receipt consumed by Rust before ACP initialize. */
export function reportStartup(outcome: 'ready' | 'mismatch' | 'unreachable'): void {
  process.stderr.write(`${STARTUP_PREFIX}${outcome}\n`);
}
