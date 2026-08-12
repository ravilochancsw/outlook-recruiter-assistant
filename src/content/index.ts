import { loadSettings } from '../shared/settings';
import { setDebugLogging, log } from '../shared/log';
import { messageType, type OfferDeletionMsg } from '../shared/messages';
import {
  startRecruiterAssistant,
  isAssistantMounted,
  currentCandidateEmail,
  showCleanupStatus,
} from './recruiter-ui';
import { captureSnapshot, handleDeletionOffer, hasSourceMessageOpen } from './source-email';
import { findComposeSurface } from './compose';

/**
 * Content script entry point. Runs in every Outlook frame and plays one of two
 * roles depending on what the tab is showing:
 *
 * - a **compose** tab gets the recruiter panel;
 * - an **inbox** tab answers snapshot requests and, after a confirmed send, may
 *   be asked to delete the application email it has open.
 *
 * Both roles are always available, because a tab can become either one without
 * a page load.
 */

const MIN_FRAME_WIDTH = 360;
const MIN_FRAME_HEIGHT = 280;

function isTopFrame(): boolean {
  try {
    return window.top === window.self;
  } catch {
    return false;
  }
}

function frameIsWorthAttaching(): boolean {
  if (isTopFrame()) return true;
  return window.innerWidth >= MIN_FRAME_WIDTH && window.innerHeight >= MIN_FRAME_HEIGHT;
}

function version(): string {
  try {
    return chrome.runtime.getManifest().version;
  } catch {
    return 'unknown';
  }
}

/**
 * One unconditional line, path only, no PII. Without it there is no way to tell
 * "not injected" from "injected but broken".
 */
function announce(): void {
  console.log(
    `[RecruiterAssistant] v${version()} active — ${location.host}${location.pathname}` +
      ` (${isTopFrame() ? 'top' : 'sub'} frame)`,
  );
}

function installListeners(): void {
  chrome.runtime.onMessage.addListener((message, _sender, respond) => {
    switch (messageType(message)) {
      case 'RA_CAPTURE_SNAPSHOT': {
        // Only the top frame speaks for the tab. Sub-frames answering a broadcast
        // would race the real answer and could return a null snapshot.
        if (!isTopFrame()) return false;
        const snapshot = captureSnapshot();
        log('snapshot requested', {
          haveMessage: Boolean(snapshot),
          looksLikeApplication: snapshot?.looksLikeApplication ?? false,
        });
        respond({ type: 'RA_SNAPSHOT_REPLY', snapshot });
        return true;
      }

      case 'RA_CLEANUP_STATUS': {
        const { outcome, reason } = message as { outcome: 'offered' | 'blocked'; reason: string };
        showCleanupStatus(outcome, reason);
        respond({ ok: true });
        return true;
      }

      case 'RA_OFFER_DELETION': {
        // Deliberately NOT restricted to the top frame. That guard was added on a
        // theory and is a candidate cause of the prompt not appearing; letting every
        // frame handle the offer is what the last known-good build did.
        handleDeletionOffer(message as OfferDeletionMsg);
        respond({ ok: true });
        return true;
      }

      case 'RA_PING': {
        respond({
          type: 'RA_PONG',
          version: version(),
          host: location.host,
          path: location.pathname,
          topFrame: isTopFrame(),
          assistantMounted: isAssistantMounted(),
          composeDetected: Boolean(findComposeSurface().surface),
          candidateEmail: currentCandidateEmail(),
          sourceMessageOpen: hasSourceMessageOpen(),
          snapshot: captureSnapshot(),
        });
        return true;
      }

      default:
        return false;
    }
  });
}

/**
 * True while this tab can still talk to the extension.
 *
 * Reloading the extension in chrome://extensions severs `chrome.runtime` for every
 * already-open tab. The old content script keeps running, but the service worker's
 * `chrome.tabs.sendMessage` to it fails — so a cleanup offer silently never
 * arrives and no prompt appears. Nothing in the page looks wrong, which makes this
 * indistinguishable from a bug in the prompt itself.
 */
function contextAlive(): boolean {
  try {
    return Boolean(chrome.runtime?.id);
  } catch {
    return false;
  }
}

const ORPHAN_HOST_ID = 'recruiter-assistant-orphan';

function showReloadBanner(): void {
  if (document.getElementById(ORPHAN_HOST_ID)) return;
  const host = document.createElement('div');
  host.id = ORPHAN_HOST_ID;
  document.documentElement.append(host);
  const root = host.attachShadow({ mode: 'open' });

  const style = document.createElement('style');
  style.textContent = `
    :host { all: initial; }
    .bar { position: fixed; bottom: 18px; right: 18px; width: 320px; z-index: 2147483645;
      background: #fff4ce; color: #6b5300; border: 1px solid #e8c96a; border-radius: 8px;
      padding: 12px 14px; font: 13px/1.5 "Segoe UI", system-ui, sans-serif;
      box-shadow: 0 8px 28px rgba(0,0,0,.18); }
    strong { display: block; margin-bottom: 4px; }
    button { margin-top: 10px; padding: 6px 12px; border-radius: 5px; border: 1px solid #b58b17;
      background: #fff; color: #6b5300; font: inherit; font-size: 12.5px; cursor: pointer; }
  `;
  const bar = document.createElement('div');
  bar.className = 'bar';
  const title = document.createElement('strong');
  title.textContent = 'Recruiter Assistant needs this tab reloaded';
  const body = document.createElement('span');
  body.textContent =
    'The extension was reloaded or updated, so this tab can no longer reach it. Templates and the cleanup prompt will not work until you refresh.';
  const button = document.createElement('button');
  button.textContent = 'Reload this tab';
  button.addEventListener('click', () => location.reload());
  bar.append(title, body, button);
  root.append(style, bar);
}

/**
 * Watches for the connection being severed. Checked on a slow interval and
 * whenever the tab is brought forward, which is when the user is about to rely on
 * it — for the inbox tab, that is exactly the moment the cleanup prompt is due.
 */
function watchForOrphaning(): void {
  if (!isTopFrame()) return;
  const check = () => {
    if (contextAlive()) return;
    console.warn('[RecruiterAssistant] extension context invalidated — this tab needs reloading');
    showReloadBanner();
    window.clearInterval(timer);
  };
  const timer = window.setInterval(check, 20_000);
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') check();
  });
  window.addEventListener('focus', check);
}

async function init(): Promise<void> {
  announce();
  installListeners();

  const settings = await loadSettings();
  setDebugLogging(settings.debugLogging);
  watchForOrphaning();

  if (!frameIsWorthAttaching()) return;

  try {
    await startRecruiterAssistant();
  } catch (err) {
    console.error('[RecruiterAssistant] failed to start', err);
  }
}

void init().catch((err) => {
  console.error('[RecruiterAssistant] init failed', err);
});
