import { ACTION_ORDER, type ActionId, type Settings } from '../shared/config';
import { renderTemplate } from '../shared/template';
import { loadSettings } from '../shared/settings';
import { findComposeSurface, resolveRecipient, type ComposeSurface, type Recipient } from './compose';
import { fillCompose, inspectExisting } from './fill';
import { armSendDetection } from './send-detector';
import { setDebugLogging } from '../shared/log';
import { sendToBackground } from '../shared/send-message';

/**
 * The recruiter panel. Shows the five actions once a compose window with exactly
 * one recipient is detected, fills the chosen template, and then gets out of the
 * way so the user reviews and presses Outlook's own Send button.
 *
 * It never clicks Send. It never deletes anything.
 */

const HOST_ID = 'recruiter-assistant-panel';

const STYLE = `
:host { all: initial; }
* { box-sizing: border-box; font-family: "Segoe UI", system-ui, sans-serif; }
.wrap {
  position: fixed; bottom: 18px; right: 18px; width: 320px;
  background: #fff; color: #201f1e; border: 1px solid #d1d1d1; border-radius: 8px;
  box-shadow: 0 8px 28px rgba(0,0,0,.20); z-index: 2147483644; font-size: 13px; overflow: hidden;
}
header { display: flex; align-items: center; gap: 8px; padding: 9px 12px;
  background: #0f6cbd; color: #fff; cursor: move; user-select: none; }
header .t { flex: 1; font-weight: 600; font-size: 12px; }
header button { background: rgba(255,255,255,.18); color: #fff; border: 0; border-radius: 3px;
  width: 20px; height: 20px; cursor: pointer; font-size: 13px; line-height: 1; }
header button:hover { background: rgba(255,255,255,.32); }
.body { padding: 12px; }
.who { padding: 8px 10px; background: #f3f2f1; border-radius: 4px; margin-bottom: 10px; }
.who .name { font-weight: 600; }
.who .mail { font-family: ui-monospace, Consolas, monospace; font-size: 12px; word-break: break-all; }
.who .src { color: #605e5c; font-size: 11px; margin-top: 3px; }
.actions { display: flex; flex-direction: column; gap: 6px; }
button.act { text-align: left; padding: 9px 11px; border-radius: 5px; cursor: pointer;
  border: 1px solid #c8c6c4; background: #fff; font: inherit; }
button.act:hover { background: #f3f2f1; }
button.act { position: relative; padding-right: 38px; }
button.act .lab { font-weight: 600; display: block; }
button.act .key { position: absolute; top: 8px; right: 8px; font-size: 10.5px;
  font-family: ui-monospace, Consolas, monospace; color: #605e5c;
  border: 1px solid #d1d1d1; border-radius: 3px; padding: 0 4px; line-height: 15px; }
.hint-row { margin-top: 8px; color: #605e5c; font-size: 11px; }
button.act .sub { color: #605e5c; font-size: 11.5px; }
button.act.neg { border-left: 4px solid #c4314b; }
button.act.pos { border-left: 4px solid #107c10; }
button.act:disabled { opacity: .5; cursor: default; }
.msg { margin-top: 10px; padding: 9px 11px; border-radius: 4px; font-size: 12.5px; }
.msg.ok { background: #dff6dd; color: #0b5a0b; }
.msg.warn { background: #fff4ce; color: #6b5300; }
.msg.err { background: #fde7e9; color: #a4262c; }
.msg .hint { display: block; margin-top: 5px; color: inherit; opacity: .85; }
.confirm { margin-top: 10px; padding: 9px 11px; background: #fff4ce; border-radius: 4px; font-size: 12.5px; }
.confirm .row { display: flex; gap: 6px; margin-top: 8px; }
.confirm button { flex: 1; padding: 5px; border-radius: 4px; border: 1px solid #c8c6c4;
  background: #fff; font: inherit; font-size: 12px; cursor: pointer; }
.confirm button.go { background: #c4314b; border-color: #c4314b; color: #fff; }
.audit { margin-top: 8px; color: #605e5c; font-size: 11px; }
`;

interface State {
  settings: Settings;
  recipient?: Recipient;
  surface?: ComposeSurface;
  problem?: string;
  busy: boolean;
  filledWith?: ActionId;
  /**
   * True from the moment a template is written until the send outcome is known.
   * The compose surface disappearing is precisely the signal we are waiting for,
   * so the panel must survive it — otherwise the outcome is rendered into a
   * panel that has already been removed and the user sees nothing.
   */
  awaitingOutcome: boolean;
  /**
   * Once a send verdict is in, the panel shows only that. Falling back to the
   * "waiting for a recipient" state would be technically true — compose is gone —
   * but it buries the one thing the user needs to read.
   */
  verdictOnly: boolean;
}

let host: HTMLElement | null = null;
let root: ShadowRoot | null = null;
let state: State | null = null;
let pendingConfirm: ActionId | null = null;
/** Set by the × button so the observer does not immediately re-mount the panel. */
let dismissedFor: string | null = null;

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  props: Partial<HTMLElementTagNameMap[K]> = {},
  children: Array<Node | string> = [],
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  Object.assign(node, props);
  for (const child of children) node.append(child);
  return node;
}

function message(kind: 'ok' | 'warn' | 'err', text: string, hint?: string): HTMLElement {
  const box = el('div', { className: `msg ${kind}` }, [text]);
  if (hint) box.append(el('span', { className: 'hint', textContent: hint }));
  return box;
}

/* ------------------------------------------------------------------ */
/* actions                                                             */
/* ------------------------------------------------------------------ */

function apply(action: ActionId, replaceBody: boolean): void {
  if (!state?.surface || !state.recipient) return;
  state.busy = true;
  render();

  const template = state.settings.templates[action];
  const rendered = renderTemplate(
    template,
    { email: state.recipient.email, name: state.recipient.name },
    state.settings,
  );

  if (!rendered.ok) {
    state.busy = false;
    render(
      message(
        'err',
        `${template.label} was not applied.`,
        rendered.errors.join(' · '),
      ),
    );
    return;
  }

  // Re-read the recipient immediately before writing: the user may have edited
  // the To field while the panel was open, and the email must match what the
  // template was rendered for.
  const recheck = resolveRecipient();
  if (
    !recheck.recipient ||
    recheck.recipient.email.toLowerCase() !== state.recipient.email.toLowerCase()
  ) {
    state.busy = false;
    detect();
    render(
      message(
        'err',
        'The recipient changed while the panel was open.',
        'Nothing was written. Check who the email is addressed to and pick an action again.',
      ),
    );
    return;
  }

  const outcome = fillCompose(state.surface, rendered.subject, rendered.bodyHtml, { replaceBody });
  state.busy = false;
  state.filledWith = outcome.ok ? action : undefined;

  if (!outcome.ok) {
    render(message('err', `${template.label}: could not fill the compose window.`, outcome.errors.join(' ')));
    return;
  }

  // Register the fill so the service worker can tie this compose tab to the
  // inbox tab it was opened from, then start watching for a real send.
  sendToBackground({
    type: 'RA_REGISTER_FILL',
    candidateEmail: state.recipient.email,
    actionLabel: template.label,
  });

  state.awaitingOutcome = true;
  armSendDetection((sendOutcome) => {
    if (!state) return;
    state.awaitingOutcome = false;
    state.verdictOnly = true;
    mount();
    if (sendOutcome.confirmed) {
      render(
        message(
          'ok',
          'Sent.',
          'If the original LinkedIn email is still open in its tab, a cleanup prompt appears there.',
        ),
      );
    } else {
      render(
        message(
          'warn',
          'Not sent.',
          `${sendOutcome.reason ?? 'The compose window closed'} — the original application email was left alone.`,
        ),
      );
    }
    // The compose window is gone by now; leave the verdict up long enough to
    // read, then get out of the way.
    window.setTimeout(() => {
      if (!state) return;
      state.verdictOnly = false;
      if (!state.surface) unmount();
    }, 12_000);
  });

  const warnings = rendered.warnings.length ? ` ${rendered.warnings.join(' ')}` : '';
  const audit = el('div', {
    className: 'audit',
    textContent: `subject: ${outcome.method.subject ?? '—'} · body: ${outcome.method.body ?? '—'}`,
  });
  const box = message(
    warnings ? 'warn' : 'ok',
    `${template.label} prepared.`,
    `Review the email, then press Outlook's Send button yourself.${warnings}`,
  );
  render(box, audit);
}

function requestApply(action: ActionId): void {
  if (!state?.surface) return;
  const existing = inspectExisting(state.surface);
  if (existing.hasUserContent && state.filledWith !== action) {
    pendingConfirm = action;
    render();
    return;
  }
  apply(action, false);
}

/* ------------------------------------------------------------------ */
/* render                                                             */
/* ------------------------------------------------------------------ */

function render(...extras: HTMLElement[]): void {
  if (!root || !state) return;
  const body = root.querySelector('.body');
  if (!body) return;
  body.replaceChildren();

  if (state.verdictOnly) {
    for (const extra of extras) body.append(extra);
    return;
  }

  if (state.problem || !state.recipient) {
    body.append(
      message(
        'warn',
        state.problem ?? 'Waiting for a candidate recipient.',
        'The panel needs a compose window with exactly one recipient.',
      ),
    );
    for (const extra of extras) body.append(extra);
    return;
  }

  const who = el('div', { className: 'who' });
  if (state.recipient.name) who.append(el('div', { className: 'name', textContent: state.recipient.name }));
  who.append(el('div', { className: 'mail', textContent: state.recipient.email }));
  who.append(el('div', { className: 'src', textContent: state.recipient.sources.join(' · ') }));
  body.append(who);

  if (pendingConfirm) {
    const template = state.settings.templates[pendingConfirm];
    const confirm = el('div', { className: 'confirm' }, [
      `This draft already has content. Applying "${template.label}" will insert the template at the top.`,
    ]);
    const row = el('div', { className: 'row' });
    const replace = el('button', { className: 'go', textContent: 'Replace everything' });
    const insert = el('button', { textContent: 'Insert at top' });
    const cancel = el('button', { textContent: 'Cancel' });
    replace.addEventListener('click', () => {
      const action = pendingConfirm as ActionId;
      pendingConfirm = null;
      apply(action, true);
    });
    insert.addEventListener('click', () => {
      const action = pendingConfirm as ActionId;
      pendingConfirm = null;
      apply(action, false);
    });
    cancel.addEventListener('click', () => {
      pendingConfirm = null;
      render();
    });
    row.append(replace, insert, cancel);
    confirm.append(row);
    body.append(confirm);
    return;
  }

  const actions = el('div', { className: 'actions' });
  for (const id of ACTION_ORDER) {
    const template = state.settings.templates[id];
    const button = el('button', {
      className: `act ${template.tone === 'positive' ? 'pos' : 'neg'}`,
      title: template.description,
      disabled: state.busy,
    });
    // Show the subject as it will actually be sent, not the raw {{JOB_TITLE}}.
    const preview = renderTemplate(
      template,
      { email: state.recipient.email, name: state.recipient.name },
      state.settings,
    );
    const index = ACTION_ORDER.indexOf(id) + 1;
    const key = el('span', { className: 'key', textContent: `⌃${index}` });
    button.append(key);
    button.append(el('span', { className: 'lab', textContent: template.label }));
    button.append(el('span', { className: 'sub', textContent: preview.subject || template.subject }));
    button.addEventListener('click', () => requestApply(id));
    actions.append(button);
  }
  body.append(actions);
  body.append(
    el('div', {
      className: 'hint-row',
      textContent: 'Ctrl+1–5 to apply · ⌘+Enter in Outlook to send',
    }),
  );

  for (const extra of extras) body.append(extra);
}

/**
 * Ctrl+1..5 applies the corresponding action.
 *
 * The combination is heavily constrained, and every alternative is ruled out by
 * something concrete:
 *
 * - a **bare digit** would fire a template at a real candidate the moment it was
 *   typed into the compose body, which is where the caret is while the panel shows;
 * - **Cmd+digit** never reaches the page — Chrome reserves it for tab switching;
 * - **Cmd+Shift+3/4/5** are macOS screenshot shortcuts;
 * - **Option+digit** collides with third-party key remappers, and on macOS it also
 *   produces a symbol rather than a digit.
 *
 * Ctrl+digit it is. Note that macOS may bind Ctrl+digit to Spaces switching in
 * System Settings → Keyboard Shortcuts → Mission Control; if a digit does nothing,
 * that is where it went.
 *
 * The default is suppressed so the keystroke never reaches Outlook's editor.
 */
function onShortcut(event: KeyboardEvent): void {
  if (!state || !root) return;
  // Ctrl is the modifier, as requested. Cmd+digit is unusable — Chrome reserves it
  // for tab switching and the page never receives the event — and Option is
  // commonly claimed by key-remapping utilities. Ctrl+Ctrl-with-Cmd and
  // Ctrl+Shift both pass, so the same digits work however they are reached.
  if (!event.ctrlKey || event.altKey) return;
  // Only while the action list is actually on screen.
  if (state.busy || state.verdictOnly || pendingConfirm || !state.recipient || !state.surface) return;

  // Match the physical key: a modified event.key can be a symbol rather than a digit.
  const match = /^(?:Digit|Numpad)([1-9])$/.exec(event.code);
  const index = match?.[1] ? Number(match[1]) - 1 : -1;
  const action = ACTION_ORDER[index];
  if (!action) return;

  event.preventDefault();
  event.stopPropagation();
  requestApply(action);
}

function makeDraggable(wrap: HTMLElement, handle: HTMLElement): void {
  handle.addEventListener('mousedown', (e) => {
    if ((e.target as HTMLElement).tagName === 'BUTTON') return;
    const rect = wrap.getBoundingClientRect();
    const dx = e.clientX - rect.left;
    const dy = e.clientY - rect.top;
    const move = (ev: MouseEvent) => {
      wrap.style.left = `${ev.clientX - dx}px`;
      wrap.style.top = `${ev.clientY - dy}px`;
      wrap.style.right = 'auto';
      wrap.style.bottom = 'auto';
    };
    const up = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
    e.preventDefault();
  });
}

function mount(): void {
  if (host?.isConnected) return;
  host = document.createElement('div');
  host.id = HOST_ID;
  document.documentElement.append(host);
  root = host.attachShadow({ mode: 'open' });

  const style = document.createElement('style');
  style.textContent = STYLE;
  root.append(style);

  const wrap = el('div', { className: 'wrap' });
  const header = el('header');
  header.append(el('span', { className: 't', textContent: 'Recruiter Assistant' }));
  const close = el('button', { textContent: '×', title: 'Hide for this candidate' });
  close.addEventListener('click', () => {
    // Without recording the dismissal the MutationObserver re-mounts within one
    // debounce, and the panel appears to ignore the ×.
    dismissedFor = state?.recipient?.email ?? 'unknown';
    unmount();
  });
  header.append(close);
  wrap.append(header, el('div', { className: 'body' }));
  root.append(wrap);
  makeDraggable(wrap, header);
  document.addEventListener('keydown', onShortcut, true);
}

function unmount(): void {
  document.removeEventListener('keydown', onShortcut, true);
  host?.remove();
  host = null;
  root = null;
  // Force the next mount to re-render rather than trusting a stale key.
  lastKey = '';
}

/* ------------------------------------------------------------------ */
/* detection loop                                                      */
/* ------------------------------------------------------------------ */

function detect(): void {
  if (!state) return;
  const previousEmail = state.recipient?.email;

  const { surface, problem: surfaceProblem } = findComposeSurface();
  if (!surface) {
    state.surface = undefined;
    state.recipient = undefined;
    state.problem = surfaceProblem;
    return;
  }
  const { recipient, problem } = resolveRecipient();
  state.surface = surface;
  state.recipient = recipient;
  state.problem = problem;

  /*
   * A different candidate means a different email. Every flag scoped to the
   * previous one has to go, or the panel keeps suppressing re-renders
   * (filledWith), shows a confirm box for the wrong action (pendingConfirm), or
   * stays hidden after a × that was meant for someone else.
   */
  if (recipient?.email !== previousEmail) {
    state.filledWith = undefined;
    pendingConfirm = null;
    if (dismissedFor !== recipient?.email) dismissedFor = null;
  }
}

let debounce: number | undefined;
let observer: MutationObserver | null = null;
let lastKey = '';
let settingsRevision = 0;

function sync(): void {
  if (!state) return;
  detect();

  if (dismissedFor && dismissedFor === state.recipient?.email) return;

  const hasSurface = Boolean(state.surface);
  if (!hasSurface) {
    // Keep the panel while a send verdict is pending — its arrival is the whole
    // point, and it lands after the compose surface is gone.
    if (!state.awaitingOutcome && !state.verdictOnly) {
      unmount();
      lastKey = '';
    }
    return;
  }

  // Idempotent: only re-render when something meaningful changed, so Outlook's
  // constant DOM churn does not rebuild the panel (or duplicate it).
  if (state.verdictOnly) return;
  // settingsRevision is part of the key so a template edited in Options is
  // reflected in an already-open panel.
  const key = `${state.recipient?.email ?? ''}|${state.problem ?? ''}|${state.filledWith ?? ''}|${settingsRevision}`;
  mount();
  if (key !== lastKey) {
    const previous = lastKey;
    lastKey = key;
    // Suppress re-render only while nothing meaningful moved. A filled panel whose
    // recipient or problem changed must repaint — otherwise it keeps affirming the
    // address the template was written for after the user edited the To field.
    const onlyFilledFlagChanged =
      previous.split('|').slice(0, 2).join('|') === key.split('|').slice(0, 2).join('|');
    if (state.filledWith && onlyFilledFlagChanged) return;
    render();
  }
}

function schedule(): void {
  if (debounce) window.clearTimeout(debounce);
  debounce = window.setTimeout(sync, 400);
}

export async function startRecruiterAssistant(): Promise<void> {
  const settings = await loadSettings();
  state = { settings, busy: false, awaitingOutcome: false, verdictOnly: false };

  // Settings changes (a template edit) should be picked up without a reload.
  chrome.storage.onChanged.addListener(() => {
    void loadSettings().then((next) => {
      if (!state) return;
      state.settings = next;
      setDebugLogging(next.debugLogging);
      settingsRevision += 1;
      // Repaint so edited labels and subjects show without a page reload.
      if (!state.verdictOnly && state.surface) render();
    });
  });

  observer?.disconnect();
  observer = new MutationObserver(schedule);
  observer.observe(document.documentElement, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ['aria-label', 'role', 'contenteditable', 'aria-hidden', 'title'],
  });

  // Outlook renders compose asynchronously after the deep link resolves.
  sync();
  for (const delay of [400, 1200, 2500, 5000]) window.setTimeout(sync, delay);
}

export function isAssistantMounted(): boolean {
  return host != null && host.isConnected;
}

/**
 * Shows why the original email was or was not cleaned up. Without this the
 * decision lived only in the service-worker console and "no prompt appeared" was
 * indistinguishable from every possible cause of it.
 */
export function showCleanupStatus(outcome: 'offered' | 'blocked', reason: string): void {
  if (!state) return;
  state.verdictOnly = true;
  mount();
  render(
    outcome === 'offered'
      ? message('ok', 'Sent — cleanup prompt is waiting in your inbox tab.', reason)
      : message('warn', 'Sent, but the original email was left alone.', reason),
  );
  window.setTimeout(() => {
    if (!state) return;
    state.verdictOnly = false;
    if (!state.surface) unmount();
  }, 20_000);
}

/** For the toolbar popup's status report. */
export function currentCandidateEmail(): string | undefined {
  return state?.recipient?.email;
}
