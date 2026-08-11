import type { OfferDeletionMsg, SourceSnapshot } from '../shared/messages';
import { extractEmails } from '../shared/redact';
import { sendToBackground } from '../shared/send-message';

/**
 * The source (inbox) side of the workflow: identifying the LinkedIn application
 * email that is open, and deleting it — but only ever the exact message that was
 * open when the compose window was opened from it.
 *
 * Deletion happens through Outlook's own Delete control, found by accessible
 * name. No coordinates, no keyboard shortcuts, no APIs.
 */

const APPLICATION_SUBJECT_RE = /(new application|new applicant|has applied|application:)/i;
const LINKEDIN_RE = /linkedin/i;
const DELETE_RE = /^\s*delete\s*$/i;

function isRendered(el: Element): boolean {
  if (el.getAttribute('aria-hidden') === 'true') return false;
  const check = (el as Element & { checkVisibility?: (o?: object) => boolean }).checkVisibility;
  if (typeof check === 'function') return check.call(el, { checkOpacity: true, checkVisibilityCSS: true });
  const style = getComputedStyle(el);
  return style.display !== 'none' && style.visibility !== 'hidden';
}

function text(el: Element | null | undefined): string {
  return (el?.textContent ?? '').replace(/\s+/g, ' ').trim();
}

function accName(el: Element): string {
  return (el.getAttribute('aria-label') ?? el.getAttribute('title') ?? text(el))
    .replace(/\s+/g, ' ')
    .trim();
}

/** Outlook puts the open message's id in the path: /mail/inbox/id/<id>. */
export function readMessageIdFromUrl(href = location.href): string | null {
  try {
    const match = new URL(href).pathname.match(/\/id\/([^/?#]+)/);
    return match?.[1] ? decodeURIComponent(match[1]) : null;
  } catch {
    return null;
  }
}

/** Containers that hold *other* messages — never read subject or sender from these. */
const LIST_ROLES = '[role="listbox"], [role="grid"], [role="tree"], [role="table"], [role="list"]';

/**
 * Locates the reading pane, reporting which strategy matched.
 *
 * There is a `[role="main"]` fallback because requiring `[role="document"]` or a
 * conveniently-named region is a guess about Outlook's markup, and when the guess
 * is wrong nothing is found, no snapshot is taken, and cleanup silently never
 * happens. The audit finding that motivated dropping it — reading the subject or
 * sender out of the *message list* — is instead handled by excluding list
 * containers wherever those two things are read.
 */
/** True for the attachment-preview overlay, which is not the message. */
function isAttachmentPreview(el: Element): boolean {
  const sample = text(el).slice(0, 500);
  return /download/i.test(sample) && /(print|save to onedrive)/i.test(sample);
}

/**
 * Every plausible reading-pane container, most specific first.
 *
 * They are all returned rather than just the first, because the resume PDF preview
 * is open at the same time as the message — it is what the candidate's address was
 * clicked in — and it presents its own document-ish surface. Committing to the
 * first match meant the sender and subject were read out of the *resume*, which
 * never looks like a LinkedIn application, so cleanup was silently refused.
 */
function readingPaneCandidates(): Array<{ el: Element; via: string }> {
  const out: Array<{ el: Element; via: string }> = [];
  const seen = new Set<Element>();
  const add = (el: Element | null | undefined, via: string) => {
    if (!el || seen.has(el) || !isRendered(el)) return;
    if (isAttachmentPreview(el)) return;
    seen.add(el);
    out.push({ el, via });
  };

  for (const el of Array.from(document.querySelectorAll('[role="document"], article'))) {
    add(el, 'role=document/article');
  }
  for (const el of Array.from(
    document.querySelectorAll('[role="region"], [role="main"], [role="group"], [role="complementary"]'),
  )) {
    if (/reading|message|conversation/i.test(accName(el))) {
      add(el, 'region named reading/message/conversation');
    }
  }
  for (const el of Array.from(document.querySelectorAll('[role="main"]'))) {
    add(el, 'role=main (message lists excluded when reading)');
  }
  return out;
}

/**
 * Picks the candidate that actually reads as the open message.
 *
 * A pane that yields a LinkedIn sender and an application subject wins outright.
 * Failing that, the first pane that yields any sender and subject is returned so
 * the popup can show what was read and why it did not qualify.
 */
function resolveMessage(): { sender: string; subject: string; via: string } {
  const candidates = readingPaneCandidates();
  let fallback: { sender: string; subject: string; via: string } | null = null;

  for (const candidate of candidates) {
    const subject = readSubject(candidate.el);
    const sender = readSender(candidate.el);
    if (judgeApplication(sender, subject)) {
      return { sender, subject, via: candidate.via };
    }
    if (!fallback && (sender || subject)) {
      fallback = { sender, subject, via: `${candidate.via} — did not qualify` };
    }
  }

  return fallback ?? { sender: '', subject: '', via: 'no reading pane found' };
}

/** Excludes nodes sitting inside a message list, so only the open message is read. */
function outsideLists(el: Element): boolean {
  return !el.closest(LIST_ROLES);
}

/**
 * The rendered message body, so sender detection can exclude it.
 *
 * Otherwise the word "LinkedIn" appearing anywhere in a message — a signature, a
 * shared profile link, a colleague's email — makes that message read as being
 * *from* LinkedIn, and combined with a matching subject that is enough to qualify
 * a normal email for deletion.
 */
function messageBody(pane: Element): Element | null {
  return (
    pane.querySelector('[id*="body" i], [class*="body" i], [aria-label*="message body" i]') ?? null
  );
}

function readSubject(pane: Element): string {
  const headings = Array.from(pane.querySelectorAll('h1, h2, h3, [role="heading"]'))
    .filter((el) => isRendered(el) && outsideLists(el))
    .map(text)
    .filter((t) => t.length > 3);
  // The message subject is the longest heading in the pane; navigation headings
  // ("Focused", "Today") are short.
  return headings.sort((a, b) => b.length - a.length)[0] ?? '';
}

function readSender(pane: Element): string {
  const body = messageBody(pane);
  // The element must neither sit inside the message body nor wrap it: an ancestor's
  // textContent includes the body, so a message that merely mentions LinkedIn would
  // otherwise make its own container read as the sender.
  const nodes = Array.from(pane.querySelectorAll('[title], [aria-label], a, button, span')).filter(
    (el) =>
      isRendered(el) &&
      outsideLists(el) &&
      !(body && (body.contains(el) || el.contains(body))),
  );
  for (const el of nodes) {
    const label = `${el.getAttribute('aria-label') ?? ''} ${el.getAttribute('title') ?? ''} ${text(el)}`;
    if (LINKEDIN_RE.test(label)) return 'LinkedIn';
  }
  for (const el of nodes) {
    const label = `${el.getAttribute('aria-label') ?? ''} ${el.getAttribute('title') ?? ''}`;
    const [email] = extractEmails(label);
    if (email) return email;
  }
  return '';
}

export function captureSnapshot(): SourceSnapshot | null {
  const messageId = readMessageIdFromUrl();
  if (!messageId) return null;

  const { sender, subject, via } = resolveMessage();

  return {
    messageId,
    sender,
    subject,
    looksLikeApplication: judgeApplication(sender, subject),
    capturedAt: Date.now(),
    paneVia: via,
  };
}

/**
 * Two independent signals are required: the sender must look like LinkedIn *and*
 * the subject must look like an application notification. One alone is not
 * enough — a normal email from a colleague mentioning LinkedIn must never
 * qualify.
 */
export function judgeApplication(sender: string, subject: string): boolean {
  const senderIsLinkedIn = LINKEDIN_RE.test(sender);
  const subjectIsApplication = APPLICATION_SUBJECT_RE.test(subject);
  return senderIsLinkedIn && subjectIsApplication;
}

/* ------------------------------------------------------------------ */
/* deletion                                                           */
/* ------------------------------------------------------------------ */

export interface DeleteVerdict {
  ok: boolean;
  detail: string;
}

/**
 * The resume PDF preview is still open in this tab — it is what the candidate's
 * address was clicked in. It renders as a full overlay above the reading pane,
 * which is where the message Delete control lives. Closing it first puts the tab
 * back into the ordinary reading-pane state before anything is clicked.
 *
 * Guarded so it only ever closes an *attachment* preview: a Close control is
 * required to sit alongside something that looks like an attachment filename.
 */
function dismissAttachmentPreview(): boolean {
  const closers = Array.from(
    document.querySelectorAll<HTMLElement>('button, [role="button"]'),
  ).filter((el) => isRendered(el) && /^\s*(close|dismiss)\b/i.test(accName(el)));

  for (const closer of closers) {
    const surface =
      closer.closest('[role="dialog"], [role="region"], [role="group"]') ??
      closer.parentElement?.parentElement ??
      null;
    if (!surface) continue;
    const looksLikeAttachment = /\.(pdf|docx?|png|jpe?g)\b/i.test(text(surface).slice(0, 400));
    const hasPreviewControls = /download|print|save to onedrive/i.test(text(surface).slice(0, 400));
    if (looksLikeAttachment && hasPreviewControls) {
      closer.click();
      return true;
    }
  }
  return false;
}

/**
 * Outlook hides Delete behind a "More actions" menu whenever the ribbon is
 * collapsed or the window is narrow, so the visible-toolbar lookup alone is not
 * enough. This opens that menu and looks for an exact Delete inside it.
 */
async function findDeleteViaMoreActions(): Promise<HTMLElement | null> {
  const more = Array.from(document.querySelectorAll<HTMLElement>('button, [role="button"]')).find(
    (el) => isRendered(el) && /^\s*(more actions|more options|more commands)\b/i.test(accName(el)),
  );
  if (!more) return null;

  more.click();
  const appeared = await waitFor(
    () =>
      Array.from(document.querySelectorAll('[role="menuitem"], [role="menuitemcheckbox"]')).some(
        (el) => DELETE_RE.test(accName(el)),
      ),
    1500,
  );
  if (!appeared) return null;

  return (
    Array.from(document.querySelectorAll<HTMLElement>('[role="menuitem"]')).find(
      (el) => isRendered(el) && DELETE_RE.test(accName(el)),
    ) ?? null
  );
}

function findDeleteControl(): HTMLElement | null {
  const controls = Array.from(
    document.querySelectorAll<HTMLElement>('button, [role="button"], [role="menuitem"]'),
  ).filter(isRendered);

  // Exact "Delete" only. "Delete draft", "Delete folder", "Deleted Items" must
  // not match.
  const exact = controls.filter((el) => DELETE_RE.test(accName(el)));
  const inToolbar = exact.find((el) => el.closest('[role="toolbar"], [role="menubar"]'));
  return inToolbar ?? exact[0] ?? null;
}

/** Reports which route actually deleted the message, for the result detail. */
async function clickDelete(): Promise<{ clicked: boolean; via: string }> {
  const direct = findDeleteControl();
  if (direct) {
    direct.click();
    return { clicked: true, via: 'the Delete button' };
  }
  const viaMenu = await findDeleteViaMoreActions();
  if (viaMenu) {
    viaMenu.click();
    return { clicked: true, via: 'Delete in the More actions menu' };
  }
  return { clicked: false, via: '' };
}

/**
 * Re-verifies everything before touching Outlook. The snapshot was taken when
 * the compose tab opened; if the user has since clicked a different email, the
 * message id will differ and this refuses.
 */
export function verifySourceStillMatches(snapshot: SourceSnapshot): DeleteVerdict {
  const currentId = readMessageIdFromUrl();
  if (!currentId) {
    return { ok: false, detail: 'no message is open in this tab any more' };
  }
  if (currentId !== snapshot.messageId) {
    return { ok: false, detail: 'a different email is now open in this tab' };
  }

  const { sender, subject } = resolveMessage();

  if (!judgeApplication(sender, subject)) {
    return {
      ok: false,
      detail: 'the open email no longer reads as a LinkedIn application notification',
    };
  }
  if (subject && snapshot.subject && subject !== snapshot.subject) {
    return { ok: false, detail: 'the subject of the open email changed since it was recorded' };
  }
  return { ok: true, detail: 'message id, sender and subject all still match' };
}

export async function deleteOpenMessage(snapshot: SourceSnapshot): Promise<DeleteVerdict> {
  // Get the attachment preview out of the way first, then re-verify: closing it
  // changes what is on screen, so the checks have to run against the final state.
  const closedPreview = dismissAttachmentPreview();
  if (closedPreview) await settle(500);

  const verdict = verifySourceStillMatches(snapshot);
  if (!verdict.ok) return verdict;

  const { clicked, via } = await clickDelete();
  if (!clicked) {
    return {
      ok: false,
      detail: 'could not find Outlook’s Delete control, in the toolbar or the More actions menu',
    };
  }

  // Confirm it actually went, rather than assuming a click worked.
  const gone = await waitFor(() => {
    const id = readMessageIdFromUrl();
    if (id !== snapshot.messageId) return true;
    const { subject } = resolveMessage();
    return Boolean(snapshot.subject) && subject !== snapshot.subject;
  }, 4000);

  return gone
    ? { ok: true, detail: `deleted via ${via}` }
    : {
        ok: false,
        detail: `used ${via} but could not confirm the message went — check the inbox yourself`,
      };
}

function settle(ms: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function waitFor(predicate: () => boolean, timeoutMs: number): Promise<boolean> {
  return new Promise((resolve) => {
    const started = Date.now();
    const tick = () => {
      if (predicate()) return resolve(true);
      if (Date.now() - started > timeoutMs) return resolve(false);
      window.setTimeout(tick, 200);
    };
    tick();
  });
}

/* ------------------------------------------------------------------ */
/* confirmation UI (shown in the source tab)                           */
/* ------------------------------------------------------------------ */

const HOST_ID = 'recruiter-assistant-cleanup';

const STYLE = `
:host { all: initial; }
* { box-sizing: border-box; font-family: "Segoe UI", system-ui, sans-serif; }
.card { position: fixed; bottom: 18px; right: 18px; width: 340px; background: #fff; color: #201f1e;
  border: 1px solid #d1d1d1; border-radius: 8px; box-shadow: 0 8px 28px rgba(0,0,0,.20);
  z-index: 2147483644; font-size: 13px; overflow: hidden; }
header { padding: 9px 12px; background: #107c10; color: #fff; font-weight: 600; font-size: 12px; }
header.warn { background: #8a6100; }
header.err { background: #a4262c; }
.inner { padding: 12px; }
.subject { padding: 7px 9px; background: #f3f2f1; border-radius: 4px; margin: 8px 0;
  font-size: 12px; word-break: break-word; }
.meta { color: #605e5c; font-size: 11px; margin-top: 6px; }
.row { display: flex; gap: 8px; margin-top: 12px; }
button { flex: 1; padding: 7px; border-radius: 4px; border: 1px solid #c8c6c4; background: #fff;
  font: inherit; font-size: 12.5px; cursor: pointer; }
button.go { background: #c4314b; border-color: #c4314b; color: #fff; }
button:disabled { opacity: .6; cursor: default; }
button:focus, button:focus-visible { outline: 2px solid #0f6cbd; outline-offset: 1px; }
.keys { margin-top: 9px; color: #605e5c; font-size: 11px; }
.keys kbd { font-family: ui-monospace, Consolas, monospace; border: 1px solid #d1d1d1;
  border-radius: 3px; padding: 0 4px; background: #faf9f8; }
`;

let host: HTMLElement | null = null;
/** Bound only while a card with a confirmable delete is on screen. */
let cardKeyHandler: ((event: KeyboardEvent) => void) | null = null;

function releaseCardKeys(): void {
  if (cardKeyHandler) {
    document.removeEventListener('keydown', cardKeyHandler, true);
    cardKeyHandler = null;
  }
}

/**
 * Keyboard confirmation for the cleanup prompt.
 *
 * The Delete button is focused when the card appears, so plain Enter activates it
 * natively. Cmd/Ctrl+Enter also works from anywhere in the tab — the same gesture
 * that sent the email, so it is one piece of muscle memory rather than two — and
 * Escape keeps the message.
 *
 * Bound only while a card offering a delete is on screen, so these keys never do
 * anything in Outlook at any other moment.
 */
function bindCardKeys(onConfirm: () => void, onDismiss: () => void): void {
  releaseCardKeys();
  cardKeyHandler = (event: KeyboardEvent) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      onDismiss();
      return;
    }
    if ((event.metaKey || event.ctrlKey) && (event.key === 'Enter' || event.code === 'Enter')) {
      event.preventDefault();
      event.stopPropagation();
      onConfirm();
    }
  };
  document.addEventListener('keydown', cardKeyHandler, true);
}

function mountCard(): ShadowRoot {
  host?.remove();
  host = document.createElement('div');
  host.id = HOST_ID;
  document.documentElement.append(host);
  const root = host.attachShadow({ mode: 'open' });
  const style = document.createElement('style');
  style.textContent = STYLE;
  root.append(style);
  return root;
}

function dismissCard(): void {
  releaseCardKeys();
  host?.remove();
  host = null;
}

function card(root: ShadowRoot, kind: 'ok' | 'warn' | 'err', title: string): HTMLElement {
  const wrap = document.createElement('div');
  wrap.className = 'card';
  const head = document.createElement('header');
  head.className = kind === 'ok' ? '' : kind;
  head.textContent = title;
  const inner = document.createElement('div');
  inner.className = 'inner';
  wrap.append(head, inner);
  root.append(wrap);
  return inner;
}

async function performDelete(offer: OfferDeletionMsg): Promise<void> {
  const result = await deleteOpenMessage(offer.snapshot);
  sendToBackground({
    type: 'RA_DELETE_RESULT',
    workflowId: offer.workflowId,
    ok: result.ok,
    detail: result.detail,
  });

  const root = mountCard();
  if (result.ok) {
    const inner = card(root, 'ok', 'Original application deleted');
    inner.append(document.createTextNode(`${offer.snapshot.subject || 'The LinkedIn email'} was deleted.`));
    window.setTimeout(dismissCard, 6000);
  } else {
    const inner = card(root, 'err', 'Original application NOT deleted');
    inner.append(document.createTextNode(result.detail));
    const meta = document.createElement('div');
    meta.className = 'meta';
    meta.textContent = 'Delete it yourself if you still want it gone. Nothing was changed.';
    inner.append(meta);
    const row = document.createElement('div');
    row.className = 'row';
    const ok = document.createElement('button');
    ok.textContent = 'Dismiss';
    ok.addEventListener('click', dismissCard);
    row.append(ok);
    inner.append(row);
  }
}

export function handleDeletionOffer(offer: OfferDeletionMsg): void {
  // Re-verify before even offering, so the prompt is never shown for a message
  // that has already changed underneath us.
  const verdict = verifySourceStillMatches(offer.snapshot);
  if (!verdict.ok) {
    sendToBackground({
      type: 'RA_DELETE_RESULT',
      workflowId: offer.workflowId,
      ok: false,
      detail: verdict.detail,
    });
    const root = mountCard();
    const inner = card(root, 'warn', 'Original application left alone');
    inner.append(document.createTextNode(`${offer.actionLabel} was sent, but ${verdict.detail}.`));
    const row = document.createElement('div');
    row.className = 'row';
    const dismiss = document.createElement('button');
    dismiss.textContent = 'Dismiss';
    dismiss.addEventListener('click', dismissCard);
    row.append(dismiss);
    inner.append(row);
    return;
  }

  if (!offer.requireConfirmation) {
    void performDelete(offer);
    return;
  }

  const root = mountCard();
  const inner = card(root, 'ok', `${offer.actionLabel} sent`);
  inner.append(document.createTextNode('Delete the original LinkedIn application email?'));

  const subject = document.createElement('div');
  subject.className = 'subject';
  subject.textContent = offer.snapshot.subject || '(subject unavailable)';
  inner.append(subject);

  const meta = document.createElement('div');
  meta.className = 'meta';
  meta.textContent = `From ${offer.snapshot.sender || 'unknown sender'} · confirmed by: ${offer.signals.join(', ')}`;
  inner.append(meta);

  const row = document.createElement('div');
  row.className = 'row';
  const del = document.createElement('button');
  del.className = 'go';
  del.textContent = 'Delete original';
  const keep = document.createElement('button');
  keep.textContent = 'Keep';

  let settled = false;
  const confirm = () => {
    if (settled) return;
    settled = true;
    del.disabled = true;
    keep.disabled = true;
    releaseCardKeys();
    void performDelete(offer);
  };
  const decline = () => {
    if (settled) return;
    settled = true;
    sendToBackground({ type: 'RA_DELETE_DECLINED', workflowId: offer.workflowId });
    dismissCard();
  };

  del.addEventListener('click', confirm);
  keep.addEventListener('click', decline);
  row.append(del, keep);
  inner.append(row);

  const keys = document.createElement('div');
  keys.className = 'keys';
  const kbd = (label: string) => {
    const el = document.createElement('kbd');
    el.textContent = label;
    return el;
  };
  keys.append(kbd('Enter'), document.createTextNode(' or '), kbd('⌘/Ctrl+Enter'));
  keys.append(document.createTextNode(' to delete · '), kbd('Esc'), document.createTextNode(' to keep'));
  inner.append(keys);

  bindCardKeys(confirm, decline);

  /*
   * Focus the delete button so Enter alone works.
   *
   * The card is created while this tab is in the background — the user is still in
   * the compose tab — so focusing once is not enough. Focus has to be reclaimed
   * when they switch back, which is the moment they will actually press a key.
   */
  const claimFocus = () => {
    if (settled || !host?.isConnected) return;
    del.focus({ preventScroll: true });
  };
  claimFocus();
  window.setTimeout(claimFocus, 150);
  window.addEventListener('focus', claimFocus);
  document.addEventListener('visibilitychange', function onVisible() {
    if (document.visibilityState !== 'visible') return;
    document.removeEventListener('visibilitychange', onVisible);
    window.setTimeout(claimFocus, 50);
  });
}

export function hasSourceMessageOpen(): boolean {
  return Boolean(readMessageIdFromUrl());
}
