import type { EmailTemplate, Settings } from './config';

/**
 * Template rendering.
 *
 * Two rules drive the design:
 *
 * 1. Substituted values are escaped for the HTML body. A candidate address is
 *    untrusted text as far as markup is concerned.
 * 2. If anything is unresolved, rendering FAILS rather than producing a partly
 *    filled email. A candidate receiving a literal `{{BOOKING_URL}}` is worse
 *    than the extension refusing to act.
 */

export interface CandidateContext {
  email: string;
  name?: string;
}

export interface RenderResult {
  ok: boolean;
  subject: string;
  bodyHtml: string;
  /** Plain-text rendering, used to verify what actually landed in the editor. */
  bodyText: string;
  /** Non-fatal observations. */
  warnings: string[];
  /** Reasons rendering failed. Non-empty means do not touch the compose window. */
  errors: string[];
}

const VARIABLE_RE = /\{\{\s*([A-Z0-9_]+)\s*\}\}/g;

export const KNOWN_VARIABLES = [
  'JOB_TITLE',
  'COMPANY',
  'CANDIDATE_NAME',
  'CANDIDATE_EMAIL',
  'BOOKING_URL',
  'INTERVIEWER',
  'INTERVIEWER_TITLE',
  'SENDER_NAME',
  'SENDER_TITLE',
  'GREETING',
] as const;

export function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** https only. Returns null when unusable. */
export function validateBookingUrl(raw: string): { url: string | null; error?: string; warning?: string } {
  const trimmed = raw.trim();
  if (!trimmed) return { url: null, error: 'No Microsoft Bookings URL is configured.' };
  let parsed: URL;
  try {
    parsed = new URL(trimmed);
  } catch {
    return { url: null, error: `Bookings URL is not a valid URL: ${trimmed}` };
  }
  if (parsed.protocol !== 'https:') {
    return { url: null, error: `Bookings URL must use https, got ${parsed.protocol}` };
  }
  const looksMicrosoft = /(^|\.)(microsoft|office|office365|outlook|cloud\.microsoft)\b/i.test(
    parsed.hostname,
  );
  return {
    url: parsed.toString(),
    warning: looksMicrosoft
      ? undefined
      : `Bookings URL host "${parsed.hostname}" does not look like a Microsoft Bookings host — double-check it.`,
  };
}

/** Rough HTML → text, good enough to verify the editor contents afterwards. */
export function htmlToText(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Removes paragraphs that became empty through substitution, so a suppressed
 * greeting does not leave a blank line at the top of the email.
 *
 * `<p>&nbsp;</p>` is deliberately left alone: that is how the templates create an
 * intentional blank line between paragraphs, and stripping it would silently undo
 * the spacing. Only genuinely empty paragraphs go.
 */
function stripEmptyParagraphs(html: string): string {
  return html
    .replace(/<p>(?:\s|<br\s*\/?>)*<\/p>\s*/gi, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function greetingFor(settings: Settings, candidate: CandidateContext): string {
  if (settings.greetingMode === 'none') return '';
  const name = candidate.name?.trim();
  return name ? `Hi ${name},` : `Hi ${settings.greetingFallback},`;
}

function substitute(
  input: string,
  values: Record<string, string>,
  context: 'html' | 'text',
  unresolved: Set<string>,
  /**
   * Variables that are legitimately allowed to render as nothing. Without this,
   * an intentionally empty value (a suppressed greeting) is indistinguishable
   * from a missing one, and the literal `{{GREETING}}` survives into the body.
   */
  allowEmpty: ReadonlySet<string> = new Set(),
): string {
  return input.replace(VARIABLE_RE, (_match, rawName: string) => {
    const name = rawName.toUpperCase();
    const value = values[name];
    if (value == null || value === '') {
      if (allowEmpty.has(name)) return '';
      unresolved.add(name);
      return `{{${name}}}`;
    }
    // BOOKING_URL is pre-rendered as markup for the HTML context, so it must
    // not be escaped again.
    if (context === 'html' && name !== 'BOOKING_URL') return escapeHtml(value);
    return value;
  });
}

export function renderTemplate(
  template: EmailTemplate,
  candidate: CandidateContext,
  settings: Settings,
): RenderResult {
  const warnings: string[] = [];
  const errors: string[] = [];

  if (template.isPlaceholder) {
    warnings.push(
      `The "${template.label}" template is still marked as a placeholder — review the wording in Options before sending.`,
    );
  }

  let bookingHtml = '';
  let bookingText = '';
  const usesBooking = /\{\{\s*BOOKING_URL\s*\}\}/.test(template.bodyHtml) || template.requiresBookingUrl;
  if (usesBooking) {
    const { url, error, warning } = validateBookingUrl(template.bookingUrl ?? '');
    if (error) errors.push(`${template.label}: ${error}`);
    if (warning) warnings.push(`${template.label}: ${warning}`);
    if (url) {
      // href must be entity-escaped so `&` between query params is valid HTML;
      // the link text keeps the same escaping and decodes back on display.
      const safe = escapeHtml(url);
      bookingHtml = `<a href="${safe}">${safe}</a>`;
      bookingText = url;
    }
  }

  const greeting = greetingFor(settings, candidate);
  const htmlValues: Record<string, string> = {
    JOB_TITLE: settings.jobTitle,
    COMPANY: settings.companyName,
    CANDIDATE_NAME: candidate.name ?? '',
    CANDIDATE_EMAIL: candidate.email,
    BOOKING_URL: bookingHtml,
    INTERVIEWER: template.interviewerName ?? '',
    INTERVIEWER_TITLE: template.interviewerTitle ?? '',
    SENDER_NAME: settings.senderName,
    SENDER_TITLE: settings.senderTitle,
    GREETING: greeting,
  };
  const textValues: Record<string, string> = { ...htmlValues, BOOKING_URL: bookingText };

  const unresolvedBody = new Set<string>();
  const unresolvedSubject = new Set<string>();

  // A suppressed greeting renders as nothing, by design.
  const allowEmpty = new Set(settings.greetingMode === 'none' ? ['GREETING'] : []);

  const bodyHtml = stripEmptyParagraphs(
    substitute(template.bodyHtml, htmlValues, 'html', unresolvedBody, allowEmpty),
  );
  const subject = substitute(template.subject, textValues, 'text', unresolvedSubject, allowEmpty)
    .replace(/\s+/g, ' ')
    .trim();

  const unresolved = new Set([...unresolvedBody, ...unresolvedSubject]);
  // BOOKING_URL failure is already reported precisely above.
  unresolved.delete('BOOKING_URL');

  for (const name of unresolved) {
    const known = (KNOWN_VARIABLES as readonly string[]).includes(name);
    errors.push(
      known
        ? `{{${name}}} has no value configured — set it in Options.`
        : `{{${name}}} is not a known template variable.`,
    );
  }

  if (!subject) errors.push('Subject is empty.');
  if (!htmlToText(bodyHtml)) errors.push('Body is empty.');

  return {
    ok: errors.length === 0,
    subject,
    bodyHtml,
    bodyText: htmlToText(bodyHtml),
    warnings,
    errors,
  };
}
