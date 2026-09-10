/** Deployment boundary: only the official WebUI installs this authenticated caller.
 * Native services remain the sole business implementation in the main window. */
export type BusinessCall = <T>(method: string, payload: unknown) => Promise<T>;
let caller: BusinessCall | undefined;
export function installWebBusinessTransport(next: BusinessCall) {
  caller = next;
  return () => { if (caller === next) caller = undefined; };
}
export function webBusinessTransport() { return caller; }
