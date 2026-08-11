import { findComposeSurface } from './compose';
import { sendToBackground } from '../shared/send-message';

/**
 * Deciding whether a message was actually sent.
 *
 * The distinction that matters is **"Send was clicked" vs "the message went
 * out"**, and the one that matters even more is **"sent" vs "discarded"**. A
 * compose window disappearing proves neither on its own: it happens on send, on
 * discard, and on closing a draft.
 *
 * So sends and discards are both observed directly, in the capture phase — a Send
 * click *or* Outlook's Cmd/Ctrl+Enter shortcut, which is how this workflow is
 * actually driven — and a disappearing compose window is only read as a send when
 * one of those preceded it and no discard did. Anything else is reported as
 * abandoned, and abandonment never leads to a deletion.
 *
 * Clicks are only observed — never synthesised. This module does not click
 * anything.
 */

const SEND_RE = /^\s*send\b/i;
const DISCARD_RE = /(discard|delete draft|throw away)/i;
/** A Send click older than this is stale; treat a later disappearance as unrelated. */
const SEND_WINDOW_MS = 90_000;
const DISAPPEAR_DEBOUNCE_MS = 600;

export interface SendOutcome {
  confirmed: boolean;
  signals: string[];
  reason?: string;
}

type Listener = (outcome: SendOutcome) => void;

let armed = false;
let sendClickedAt = 0;
/** How the send was triggered, for the confirmation prompt's audit line. */
let sendVia = '';
let discardClickedAt = 0;
let sawComposeSurface = false;
let settled = false;
let listener: Listener | null = null;
let observer: MutationObserver | null = null;
let debounce: number | undefined;

function accName(el: Element): string {
  return (
    el.getAttribute('aria-label') ??
    el.getAttribute('title') ??
    el.textContent ??
    ''
  )
    .replace(/\s+/g, ' ')
    .trim();
}

/** Walks the composed path so a click on an icon inside the button still counts. */
function controlFrom(event: Event): { name: string; el: Element } | null {
  for (const node of event.composedPath()) {
    if (!(node instanceof Element)) continue;
    const role = node.getAttribute('role');
    if (node.tagName === 'BUTTON' || role === 'button' || role === 'menuitem') {
      // An inert control cannot have sent anything. Outlook disables Send while
      // a draft is invalid, and counting that click as a send is how a deletion
      // gets offered for a message that never left.
      const disabled =
        (node as HTMLButtonElement).disabled === true ||
        node.getAttribute('aria-disabled') === 'true';
      if (disabled) return null;
      return { name: accName(node), el: node };
    }
  }
  return null;
}

function onClickCapture(event: Event): void {
  if (!armed || settled) return;
  const control = controlFrom(event);
  if (!control) return;

  if (SEND_RE.test(control.name)) {
    noteSend('the Send button was clicked');
    return;
  }
  if (DISCARD_RE.test(control.name)) {
    noteDiscard();
    return;
  }
  // The discard control in the ribbon is a bare trash icon with no text. Treat an
  // unnamed control sitting next to Send as a discard candidate rather than
  // risking it being read as a send.
  if (!control.name && control.el.closest('[role="toolbar"], [role="menubar"]')) {
    const svgTitle = control.el.querySelector('title')?.textContent ?? '';
    if (DISCARD_RE.test(svgTitle)) noteDiscard();
  }
}

function noteSend(via: string): void {
  sendClickedAt = Date.now();
  sendVia = via;
  sendToBackground({ type: 'RA_SEND_CLICKED' });
}

/**
 * Outlook sends on Cmd+Enter (Ctrl+Enter on Windows), and that is how this
 * workflow is actually driven — the Send button is never clicked. Watching only
 * for clicks meant the compose window vanished with no recorded send, the outcome
 * was reported as abandoned, and the cleanup step never ran.
 */
function onKeyCapture(event: KeyboardEvent): void {
  if (!armed || settled) return;
  if (!(event.metaKey || event.ctrlKey)) return;
  if (event.key !== 'Enter' && event.code !== 'Enter' && event.code !== 'NumpadEnter') return;
  // Only meaningful while there is a compose surface to send.
  if (!findComposeSurface().surface) return;
  noteSend(`${event.metaKey ? 'Cmd' : 'Ctrl'}+Enter was pressed`);
}

/**
 * A discard is a hard veto, and it has to reach the service worker immediately:
 * if the tab is torn down before this module's debounce settles, the worker would
 * otherwise still see the workflow as SEND_CLICKED.
 */
function noteDiscard(): void {
  discardClickedAt = Date.now();
  sendToBackground({ type: 'RA_DISCARD_CLICKED' });
}

function evaluate(): void {
  if (!armed || settled) return;

  const { surface } = findComposeSurface();
  if (surface) {
    sawComposeSurface = true;
    return;
  }
  // Nothing to conclude until a compose surface has actually been seen.
  if (!sawComposeSurface) return;

  const now = Date.now();
  const sendWasClicked = sendClickedAt > 0 && now - sendClickedAt < SEND_WINDOW_MS;
  const discardWasClicked = discardClickedAt > 0;

  settled = true;

  if (discardWasClicked && discardClickedAt >= sendClickedAt) {
    finish({ confirmed: false, signals: [], reason: 'the draft was discarded' });
    return;
  }
  if (!sendWasClicked) {
    finish({
      confirmed: false,
      signals: [],
      reason: 'the compose window closed without Send being clicked',
    });
    return;
  }

  finish({
    confirmed: true,
    signals: [
      sendVia || 'a send was triggered in this tab',
      'the compose window then disappeared',
      'no discard was pressed',
    ],
  });
}

function finish(outcome: SendOutcome): void {
  if (outcome.confirmed) {
    sendToBackground({ type: 'RA_SEND_CONFIRMED', signals: outcome.signals });
  } else {
    sendToBackground({ type: 'RA_SEND_ABANDONED', reason: outcome.reason ?? 'unknown' });
  }
  listener?.(outcome);
  teardown();
}

function schedule(): void {
  if (debounce) window.clearTimeout(debounce);
  debounce = window.setTimeout(evaluate, DISAPPEAR_DEBOUNCE_MS);
}

/** Called once the template has been written, so a send can be attributed to it. */
export function armSendDetection(onOutcome: Listener): void {
  if (armed) {
    listener = onOutcome;
    return;
  }
  armed = true;
  settled = false;
  sendClickedAt = 0;
  sendVia = '';
  discardClickedAt = 0;
  sawComposeSurface = Boolean(findComposeSurface().surface);
  listener = onOutcome;

  document.addEventListener('click', onClickCapture, true);
  document.addEventListener('keydown', onKeyCapture, true);
  observer = new MutationObserver(schedule);
  observer.observe(document.documentElement, { childList: true, subtree: true });
}

export function teardown(): void {
  armed = false;
  listener = null;
  document.removeEventListener('click', onClickCapture, true);
  document.removeEventListener('keydown', onKeyCapture, true);
  observer?.disconnect();
  observer = null;
  if (debounce) window.clearTimeout(debounce);
}

/** Exposed for tests: the decision logic, free of DOM and messaging. */
export function decideOutcome(input: {
  sawComposeSurface: boolean;
  composeStillPresent: boolean;
  sendClickedAt: number;
  discardClickedAt: number;
  now: number;
  sendVia?: string;
}): SendOutcome | null {
  if (input.composeStillPresent) return null;
  if (!input.sawComposeSurface) return null;

  const sendWasClicked =
    input.sendClickedAt > 0 && input.now - input.sendClickedAt < SEND_WINDOW_MS;

  if (input.discardClickedAt > 0 && input.discardClickedAt >= input.sendClickedAt) {
    return { confirmed: false, signals: [], reason: 'the draft was discarded' };
  }
  if (!sendWasClicked) {
    return {
      confirmed: false,
      signals: [],
      reason: 'the compose window closed without Send being clicked',
    };
  }
  return {
    confirmed: true,
    signals: [
      input.sendVia || 'a send was triggered in this tab',
      'the compose window then disappeared',
      'no discard was pressed',
    ],
  };
}
