import { extractEmails } from '../shared/redact';
import { parseMailtoRecipient } from '../shared/mailto';

/**
 * Locating the parts of an Outlook compose surface.
 *
 * Every lookup is a list of strategies tried in order, all built from semantic
 * attributes (role, accessible name, placeholder). No CSS class names, no
 * generated ids. Which strategy won is reported so a change in Outlook shows up
 * as "fell back to strategy 3" rather than as silence.
 */

export interface Located<T extends Element = Element> {
  el: T;
  strategy: string;
}

export interface ComposeSurface {
  subject: Located<HTMLElement>;
  body: Located<HTMLElement>;
  /** Absent when Outlook has not rendered a recipient well we recognise. */
  recipientWell?: Located<HTMLElement>;
}

export interface Recipient {
  email: string;
  name?: string;
  /** Where the address came from, for the audit line in the panel. */
  sources: string[];
}

type Strategy<T extends Element> = [name: string, find: () => T | null | undefined];

function attempt<T extends Element>(strategies: Array<Strategy<T>>): Located<T> | null {
  for (const [strategy, find] of strategies) {
    try {
      const el = find();
      if (el && isUsable(el)) return { el, strategy };
    } catch {
      /* try the next strategy */
    }
  }
  return null;
}

function isUsable(el: Element): boolean {
  if (el.getAttribute('aria-hidden') === 'true') return false;
  const check = (el as Element & { checkVisibility?: (o?: object) => boolean }).checkVisibility;
  if (typeof check === 'function') {
    return check.call(el, { checkOpacity: true, checkVisibilityCSS: true });
  }
  const style = getComputedStyle(el);
  return style.display !== 'none' && style.visibility !== 'hidden';
}

function accName(el: Element): string {
  const labelled = el.getAttribute('aria-labelledby');
  if (labelled) {
    const parts = labelled
      .trim()
      .split(/\s+/)
      .map((id) => document.getElementById(id)?.textContent ?? '')
      .join(' ')
      .trim();
    if (parts) return parts;
  }
  return (
    el.getAttribute('aria-label') ??
    el.getAttribute('placeholder') ??
    el.getAttribute('title') ??
    ''
  ).trim();
}

function all(selector: string): HTMLElement[] {
  try {
    return Array.from(document.querySelectorAll<HTMLElement>(selector));
  } catch {
    return [];
  }
}

const SUBJECT_RE = /subject/i;
const BODY_RE = /(message body|body|message)/i;
const TO_RE = /^\s*(to\b|to:|recipients?\b)/i;

/* ------------------------------------------------------------------ */
/* subject                                                             */
/* ------------------------------------------------------------------ */

export function findSubject(): Located<HTMLElement> | null {
  return attempt<HTMLElement>([
    [
      'input with accessible name matching /subject/i',
      () => all('input, [role="textbox"]').find((el) => SUBJECT_RE.test(accName(el))),
    ],
    [
      'placeholder "Add a subject"',
      () => all('input, [contenteditable], [role="textbox"]').find((el) =>
        /add a subject/i.test(accName(el)),
      ),
    ],
    [
      'single-line input that is not the recipient well',
      () =>
        all('input[type="text"], input:not([type])').find(
          (el) => !TO_RE.test(accName(el)) && !/cc|bcc|search/i.test(accName(el)),
        ),
    ],
  ]);
}

/* ------------------------------------------------------------------ */
/* body                                                               */
/* ------------------------------------------------------------------ */

export function findBody(): Located<HTMLElement> | null {
  const editors = () =>
    all('[contenteditable="true"], [contenteditable=""]').filter((el) => isUsable(el));

  return attempt<HTMLElement>([
    [
      'contenteditable role=textbox named message/body',
      () =>
        all('[contenteditable="true"][role="textbox"], [contenteditable=""][role="textbox"]').find(
          (el) => BODY_RE.test(accName(el)),
        ),
    ],
    [
      'contenteditable named message/body',
      () => editors().find((el) => BODY_RE.test(accName(el))),
    ],
    [
      'contenteditable with aria-multiline',
      () => editors().find((el) => el.getAttribute('aria-multiline') === 'true'),
    ],
    [
      'largest contenteditable on the page',
      () =>
        editors().sort((a, b) => {
          const ra = a.getBoundingClientRect();
          const rb = b.getBoundingClientRect();
          return rb.width * rb.height - ra.width * ra.height;
        })[0],
    ],
  ]);
}

/* ------------------------------------------------------------------ */
/* recipient                                                           */
/* ------------------------------------------------------------------ */

export function findRecipientWell(): Located<HTMLElement> | null {
  return attempt<HTMLElement>([
    [
      'element named "To"',
      () =>
        all('[role="combobox"], [role="textbox"], [role="group"], input').find((el) =>
          TO_RE.test(accName(el)),
        ),
    ],
    [
      // Deliberately the labelled element's nearest *chip-bearing* descendant
      // container, never a plain ancestor: walking up reaches the compose dialog,
      // which contains `From: <your own address>`, and that reads as a second
      // recipient and blocks every action.
      'labelled "To" element that itself contains recipient chips',
      () => {
        const named = all('*').filter((el) => TO_RE.test(accName(el)));
        return (
          named.find((el) =>
            el.querySelector('[role="listitem"], [role="option"], [aria-label*="@"], [title*="@"]'),
          ) ?? null
        );
      },
    ],
  ]);
}

/**
 * Reads the candidate address from the compose recipient area only.
 *
 * Deliberately not a page-wide sweep: the reading pane behind compose is full
 * of addresses, and picking the wrong one would mean emailing the wrong person.
 */
export function readRecipientsFromDom(): Array<{ email: string; name?: string }> {
  const well = findRecipientWell();
  if (!well) return [];

  // Scoped to the To well itself, never an ancestor. Widening even one level to
  // the compose dialog sweeps up `From: <your own address>` and the panel then
  // sees two recipients and refuses to act. The sender's address is inside the
  // same dialog as the recipient chips.
  const scope = well.el;

  const found = new Map<string, { email: string; name?: string }>();
  // Descendants first: a chip's own `Name <addr>` label is more specific than
  // the well's concatenated text, which would otherwise register the address
  // with no name and win by arriving first.
  const candidates: HTMLElement[] = [
    ...Array.from(
      scope.querySelectorAll<HTMLElement>(
        '[aria-label], [title], [role="listitem"], [role="option"], span, div',
      ),
    ),
    scope,
  ];

  for (const el of candidates) {
    if (!isUsable(el)) continue;
    // A chip's remove button repeats the address; skip controls so a chip and
    // its own "Remove <addr>" button do not read as two different things.
    if (el !== scope && el.closest('button') && /remove|delete/i.test(accName(el))) continue;
    const haystack = `${el.getAttribute('aria-label') ?? ''} ${el.getAttribute('title') ?? ''} ${
      el.textContent ?? ''
    }`;
    for (const email of extractEmails(haystack)) {
      const lower = email.toLowerCase();
      const name = displayNameFrom(haystack, email);
      const existing = found.get(lower);
      if (!existing) {
        found.set(lower, { email, ...(name ? { name } : {}) });
      } else if (!existing.name && name) {
        // Upgrade: a nameless hit was recorded first, and this element knows the name.
        found.set(lower, { ...existing, name });
      }
    }
  }
  return Array.from(found.values());
}

/**
 * Pulls a display name out of a chip label of the form `Name <addr>`.
 * Returns undefined when the "name" is just the address repeated, which is what
 * Outlook renders for a mailto: deep link.
 */
function displayNameFrom(label: string, email: string): string | undefined {
  const match = label.match(/^\s*([^<]+?)\s*<\s*([^>]+)\s*>/);
  if (!match) return undefined;
  const name = match[1]?.trim();
  if (!name) return undefined;
  if (name.toLowerCase() === email.toLowerCase()) return undefined;
  if (extractEmails(name).length > 0) return undefined;
  return name;
}

/**
 * The `mailtouri` query parameter on a compose deep link. Shares one parser with
 * the service worker, which applies the same check before it will let a source
 * email be deleted.
 */
export function readRecipientFromUrl(href = location.href): string | null {
  return parseMailtoRecipient(href);
}

/**
 * Combines both sources. The DOM wins, because the user can edit the recipient
 * after the deep link opened; the URL is corroboration and is reported when the
 * two disagree so a stale URL never silently decides who gets the email.
 */
export function resolveRecipient(): { recipient?: Recipient; problem?: string } {
  const domRecipients = readRecipientsFromDom();
  const urlEmail = readRecipientFromUrl();

  if (domRecipients.length === 0) {
    return {
      problem: urlEmail
        ? 'Could not read the candidate address from the Outlook recipient field. The compose URL suggests one, but the To field is what actually gets emailed — add the recipient and try again.'
        : 'Could not identify the candidate email from Outlook.',
    };
  }

  if (domRecipients.length > 1) {
    return {
      problem: `Compose has ${domRecipients.length} recipients. Leave exactly one so there is no doubt who this email is for.`,
    };
  }

  const primary = domRecipients[0] as { email: string; name?: string };
  const sources = ['Outlook To field'];
  if (urlEmail) {
    sources.push(
      urlEmail.toLowerCase() === primary.email.toLowerCase()
        ? 'compose URL (matches)'
        : 'compose URL (differs — To field used)',
    );
  }

  return { recipient: { email: primary.email, name: primary.name, sources } };
}

/* ------------------------------------------------------------------ */
/* surface                                                             */
/* ------------------------------------------------------------------ */

export function findComposeSurface(): { surface?: ComposeSurface; problem?: string } {
  const subject = findSubject();
  const body = findBody();

  if (!subject && !body) return {};
  if (!subject) return { problem: 'Found the Outlook message editor but not the subject field.' };
  if (!body) return { problem: 'Could not find the Outlook message editor.' };

  const well = findRecipientWell();
  return { surface: { subject, body, ...(well ? { recipientWell: well } : {}) } };
}

