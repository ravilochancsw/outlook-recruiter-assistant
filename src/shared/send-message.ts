/**
 * Fire-and-forget messaging that cannot throw.
 *
 * `chrome.runtime.sendMessage` rejects when there is no listener, but it also
 * throws **synchronously** when the extension context has been invalidated —
 * which happens to every already-open tab the moment the extension is reloaded.
 * A `.catch()` does not catch that, so the content script would throw uncaught
 * errors from a tab that simply needs a refresh.
 */
export function sendToBackground(message: unknown): void {
  try {
    const result = chrome.runtime.sendMessage(message) as unknown;
    if (result && typeof (result as Promise<unknown>).catch === 'function') {
      void (result as Promise<unknown>).catch(() => undefined);
    }
  } catch {
    // Extension reloaded or updated under this tab; nothing useful to do here.
  }
}
