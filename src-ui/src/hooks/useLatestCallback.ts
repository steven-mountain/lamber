import { useCallback, useLayoutEffect, useRef } from "react";

/**
 * Keeps a callback identity stable while always invoking its latest implementation.
 * Useful for effects and registrations whose trigger lifecycle must stay independent
 * from the changing state read by the callback.
 */
export function useLatestCallback<TArgs extends unknown[], TResult>(
  callback: (...args: TArgs) => TResult,
): (...args: TArgs) => TResult {
  const callbackRef = useRef(callback);
  useLayoutEffect(() => {
    callbackRef.current = callback;
  }, [callback]);

  return useCallback((...args: TArgs) => callbackRef.current(...args), []);
}
