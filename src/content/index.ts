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
        // Only the top frame speaks for the tab. A sub-frame would build the card
        // inside an iframe where it cannot be seen, and would still answer the
        // message — so the offer could be consumed by an invisible copy.
        if (!isTopFrame()) return false;
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

async function init(): Promise<void> {
  announce();
  installListeners();

  const settings = await loadSettings();
  setDebugLogging(settings.debugLogging);

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
