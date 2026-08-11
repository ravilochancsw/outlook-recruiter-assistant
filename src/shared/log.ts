/**
 * Namespaced logging, off unless debug logging is enabled in options.
 * Never pass candidate PII here — call sites redact first.
 */

let enabled = false;

export function setDebugLogging(value: boolean): void {
  enabled = value;
}

export function log(...args: unknown[]): void {
  if (enabled) console.log('[RecruiterAssistant]', ...args);
}

export function warn(...args: unknown[]): void {
  console.warn('[RecruiterAssistant]', ...args);
}

export function error(...args: unknown[]): void {
  console.error('[RecruiterAssistant]', ...args);
}
