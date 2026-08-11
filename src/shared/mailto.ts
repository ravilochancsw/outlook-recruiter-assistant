import { extractEmails } from './redact';

/**
 * Parses the `mailtouri` parameter of an Outlook compose deep link.
 *
 * Shared between the content script (corroborating the recipient it reads from
 * the DOM) and the service worker (deciding whether a compose tab is genuinely
 * the one that was opened for a given candidate).
 */
export function parseMailtoRecipient(href: string | undefined | null): string | null {
  if (!href) return null;
  try {
    const uri = new URL(href).searchParams.get('mailtouri');
    if (!uri) return null;
    const withoutScheme = uri.replace(/^mailto:/i, '').split('?')[0] ?? '';
    const [email] = extractEmails(decodeURIComponent(withoutScheme));
    return email ? email.toLowerCase() : null;
  } catch {
    return null;
  }
}
