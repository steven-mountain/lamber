/**
 * Thin HTTP client for the lamber bridge server hosted by the Tauri backend.
 *
 * The bridge is the only way this plugin reaches lamber business logic: every
 * tool body posts JSON to a loopback endpoint and the Rust side dispatches into
 * `benefit::calculator` / `docfill`. Keeping the transport in one module means a
 * second tool never re-invents URL resolution or error normalization.
 */

/** Environment variable carrying the bridge origin, e.g. `http://127.0.0.1:41293`. */
export const BRIDGE_URL_ENV = 'LAMBER_BRIDGE_URL';

/**
 * Environment variable carrying the per-launch bridge token. Loopback is not an
 * authorization boundary — any local process could otherwise read a customer's
 * project financials — so lamber mints a token per launch and requires it.
 */
export const BRIDGE_TOKEN_ENV = 'LAMBER_BRIDGE_TOKEN';

/** Header the token travels in; lamber may override the name it expects. */
export const BRIDGE_TOKEN_HEADER_ENV = 'LAMBER_BRIDGE_TOKEN_HEADER';

const DEFAULT_TOKEN_HEADER = 'x-lamber-bridge-token';

/** Failure raised when the bridge is unreachable, misconfigured, or returns a non-2xx body. */
export class LamberBridgeError extends Error {
  constructor(message: string, readonly cause?: unknown) {
    super(message);
    this.name = 'LamberBridgeError';
  }
}

function requireEnv(key: string): string {
  const raw = process.env[key];
  if (!raw || raw.trim() === '') {
    throw new LamberBridgeError(
      `${key} is not set; the lamber backend must export it when spawning dsh.`,
    );
  }
  return raw.trim();
}

function resolveBridgeOrigin(): string {
  return requireEnv(BRIDGE_URL_ENV).replace(/\/+$/, '');
}

function bridgeHeaders(): Record<string, string> {
  const header = process.env[BRIDGE_TOKEN_HEADER_ENV]?.trim() || DEFAULT_TOKEN_HEADER;
  return {
    'content-type': 'application/json',
    [header]: requireEnv(BRIDGE_TOKEN_ENV),
  };
}

/**
 * POST a JSON payload to one bridge route and return its parsed JSON body.
 *
 * @param path - route below the bridge origin, e.g. `/lamber-bridge/calculate`.
 * @param payload - request body, serialized as JSON.
 * @param signal - the tool execution's cancellation signal.
 * @returns the parsed response body.
 */
export async function postBridge<T>(
  path: string,
  payload: unknown,
  signal: AbortSignal,
): Promise<T> {
  const url = `${resolveBridgeOrigin()}${path}`;
  let response: Response;
  try {
    response = await fetch(url, {
      method: 'POST',
      headers: bridgeHeaders(),
      body: JSON.stringify(payload),
      signal,
    });
  } catch (error) {
    if (signal.aborted) throw error;
    throw new LamberBridgeError(`lamber bridge request to ${url} failed`, error);
  }

  const text = await response.text();
  if (!response.ok) {
    throw new LamberBridgeError(
      `lamber bridge ${path} returned ${response.status}: ${text.slice(0, 500)}`,
    );
  }
  try {
    return JSON.parse(text) as T;
  } catch (error) {
    throw new LamberBridgeError(
      `lamber bridge ${path} returned a non-JSON body: ${text.slice(0, 200)}`,
      error,
    );
  }
}
