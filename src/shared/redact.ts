/**
 * Email extraction.
 *
 * This was a PII-redaction module when the extension exported diagnostic reports.
 * Those are gone, and nothing leaves the browser any more, so only the address
 * scanner remains — used to read recipient chips, senders, and the `mailtouri`
 * parameter.
 */

const EMAIL_RE = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;

export function extractEmails(input: string | null | undefined): string[] {
  if (!input) return [];
  return input.match(EMAIL_RE) ?? [];
}
