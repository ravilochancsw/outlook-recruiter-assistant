import type { ComposeSurface } from './compose';
import { htmlToText } from '../shared/template';

/**
 * Writing into Outlook's compose fields.
 *
 * Two things make this harder than assignment:
 *
 * 1. The subject is a React-controlled `<input>`. Setting `.value` updates the
 *    DOM but not React's internal state, so Outlook overwrites it on the next
 *    render and the send goes out with an empty subject. The native value setter
 *    plus a bubbling `input` event is what React actually listens for.
 * 2. The body is a rich-text `contenteditable`. `innerHTML = …` bypasses the
 *    editor's own model, so it may not persist into the draft. `execCommand`
 *    goes through the editor's input handling, which is what it expects.
 *
 * Every write is read back and compared. A write that cannot be verified is
 * reported as a failure — it never reports success it has not confirmed.
 */

export interface FillOutcome {
  ok: boolean;
  subjectWritten: boolean;
  bodyWritten: boolean;
  /** Which technique succeeded, for the audit line. */
  method: { subject?: string; body?: string };
  errors: string[];
}

export interface ExistingContent {
  subject: string;
  bodyText: string;
  hasUserContent: boolean;
}

/** Normalises whitespace and entities so read-back comparison is not brittle. */
function normalise(value: string): string {
  return value
    .replace(/ /g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function readSubject(el: HTMLElement): string {
  if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) return el.value;
  return el.textContent ?? '';
}

function readBodyText(el: HTMLElement): string {
  return el.innerText ?? el.textContent ?? '';
}

/**
 * Outlook's own signature or a draft the user already started. Filling over it
 * would destroy work, so the caller asks first.
 */
export function inspectExisting(surface: ComposeSurface): ExistingContent {
  const subject = readSubject(surface.subject.el).trim();
  const bodyText = readBodyText(surface.body.el).trim();
  return {
    subject,
    bodyText,
    // A short body is usually just Outlook's signature stub; treat anything
    // substantial as the user's own writing.
    hasUserContent: subject.length > 0 || bodyText.length > 40,
  };
}

/* ------------------------------------------------------------------ */
/* subject                                                             */
/* ------------------------------------------------------------------ */

function setViaNativeSetter(el: HTMLElement, value: string): boolean {
  const proto =
    el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
  const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
  if (!setter) return false;
  setter.call(el, value);
  el.dispatchEvent(new Event('input', { bubbles: true }));
  el.dispatchEvent(new Event('change', { bubbles: true }));
  return true;
}

function writeSubject(el: HTMLElement, subject: string): { ok: boolean; method?: string } {
  el.focus();

  if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) {
    if (setViaNativeSetter(el, subject) && normalise(readSubject(el)) === normalise(subject)) {
      return { ok: true, method: 'native value setter + input event' };
    }
    // Fallback: select-all then type through the editing pipeline.
    el.select();
    if (
      document.execCommand('insertText', false, subject) &&
      normalise(readSubject(el)) === normalise(subject)
    ) {
      return { ok: true, method: 'execCommand insertText' };
    }
    return { ok: false };
  }

  // Some Outlook builds render the subject as a contenteditable instead.
  selectAllIn(el);
  if (
    document.execCommand('insertText', false, subject) &&
    normalise(readSubject(el)) === normalise(subject)
  ) {
    return { ok: true, method: 'execCommand insertText (contenteditable subject)' };
  }
  el.textContent = subject;
  el.dispatchEvent(new InputEvent('input', { bubbles: true }));
  return normalise(readSubject(el)) === normalise(subject)
    ? { ok: true, method: 'textContent + input event' }
    : { ok: false };
}

/* ------------------------------------------------------------------ */
/* body                                                               */
/* ------------------------------------------------------------------ */

function selectAllIn(el: HTMLElement): void {
  const range = document.createRange();
  range.selectNodeContents(el);
  const selection = window.getSelection();
  selection?.removeAllRanges();
  selection?.addRange(range);
}

function placeCaretAtStart(el: HTMLElement): void {
  const range = document.createRange();
  range.selectNodeContents(el);
  range.collapse(true);
  const selection = window.getSelection();
  selection?.removeAllRanges();
  selection?.addRange(range);
}

/**
 * Inserts at the caret rather than replacing everything, so an existing
 * signature below is preserved.
 */
/**
 * Picks several distinctive lines to verify against, not just the first.
 *
 * The first line is identical across both screening templates, so checking only
 * that would let a write that changed nothing — or that the editor silently
 * sanitised — pass as verified. The last line and the longest line are far more
 * specific to the template actually being applied.
 */
function verificationLines(text: string): string[] {
  const lines = text
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line.length > 12);
  if (lines.length === 0) return [];
  const longest = [...lines].sort((a, b) => b.length - a.length)[0] as string;
  const first = lines[0] as string;
  const last = lines[lines.length - 1] as string;
  return Array.from(new Set([first, longest, last]));
}

function writeBody(
  el: HTMLElement,
  bodyHtml: string,
  expected: string[],
  replaceAll: boolean,
): { ok: boolean; method?: string } {
  el.focus();
  if (replaceAll) selectAllIn(el);
  else placeCaretAtStart(el);

  // Every one of the sampled lines must be present, so a partial insert fails.
  const contains = () => {
    if (expected.length === 0) return false;
    const actual = normalise(readBodyText(el));
    return expected.every((line) => actual.includes(normalise(line)));
  };

  if (document.execCommand('insertHTML', false, bodyHtml) && contains()) {
    return { ok: true, method: 'execCommand insertHTML' };
  }

  // Fallback: a synthetic paste, which rich-text editors also handle.
  try {
    const data = new DataTransfer();
    data.setData('text/html', bodyHtml);
    data.setData('text/plain', htmlToText(bodyHtml));
    el.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: data }));
    if (contains()) return { ok: true, method: 'synthetic paste event' };
  } catch {
    /* fall through */
  }

  // Last resort: write the DOM directly and tell the editor about it.
  if (replaceAll) el.innerHTML = bodyHtml;
  else el.insertAdjacentHTML('afterbegin', bodyHtml);
  el.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertFromPaste' }));
  if (contains()) return { ok: true, method: 'innerHTML + input event (unverified by Outlook)' };

  return { ok: false };
}

/* ------------------------------------------------------------------ */
/* public entry point                                                  */
/* ------------------------------------------------------------------ */

export function fillCompose(
  surface: ComposeSurface,
  subject: string,
  bodyHtml: string,
  options: { replaceBody: boolean },
): FillOutcome {
  const errors: string[] = [];
  const method: FillOutcome['method'] = {};

  const subjectResult = writeSubject(surface.subject.el, subject);
  if (subjectResult.ok) method.subject = subjectResult.method;
  else
    errors.push(
      'Subject could not be written, or did not read back as expected. Outlook may have changed its compose fields.',
    );

  const expected = verificationLines(htmlToText(bodyHtml));
  const bodyResult = writeBody(surface.body.el, bodyHtml, expected, options.replaceBody);
  if (bodyResult.ok) method.body = bodyResult.method;
  else
    errors.push(
      'Message body could not be written, or did not read back as expected. Nothing was sent — review the compose window before doing anything else.',
    );

  // A subject written without a body leaves a half-prepared email addressed to a
  // real candidate, so say so explicitly rather than reporting a generic failure.
  if (subjectResult.ok && !bodyResult.ok) {
    errors.push(
      'The subject was written but the body was not. Clear the compose window before doing anything else.',
    );
  }

  return {
    ok: subjectResult.ok && bodyResult.ok,
    subjectWritten: subjectResult.ok,
    bodyWritten: bodyResult.ok,
    method,
    errors,
  };
}

export { normalise as normaliseForCompare };
