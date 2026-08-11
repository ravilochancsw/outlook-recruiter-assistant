import { DEFAULT_SETTINGS } from '../shared/config';
import { loadSettings } from '../shared/settings';
import { parseMailtoRecipient } from '../shared/mailto';
import {
  messageType,
  type OfferDeletionMsg,
  type RegisterFillMsg,
  type SendConfirmedMsg,
  type SendAbandonedMsg,
  type SnapshotReplyMsg,
  type SourceSnapshot,
  type Workflow,
  type WorkflowState,
} from '../shared/messages';

/**
 * Cross-tab coordinator.
 *
 * The hard problem is knowing *which* inbox tab holds the application email that
 * a given compose window came from. Rather than guess by matching subjects or
 * candidate names, this uses `openerTabId`: Chrome tells us which tab opened the
 * compose tab, and clicking the candidate's address inside the PDF preview is
 * exactly such an open. That link is authoritative.
 *
 * The moment a compose tab appears, its opener is asked to snapshot the message
 * it currently has open — Outlook's own message id, plus sender and subject. A
 * deletion later on is only ever offered for that exact message id.
 *
 * State lives in chrome.storage.session because an MV3 service worker is killed
 * between events.
 */

const KEY = 'workflows';
const LAST_DECISION = 'lastCleanupDecision';
const OPENERS = 'openers';
const MAX_AGE_MS = 60 * 60 * 1000;
/** An opener record older than this can no longer authorise a deletion. */
const OPENER_MAX_AGE_MS = 30 * 60 * 1000;
/** Must match send-detector's window: a Send click older than this is stale. */
const SEND_WINDOW_MS = 90_000;

interface OpenerRecord {
  sourceTabId: number;
  snapshot: SourceSnapshot | null;
  createdAt: number;
  /**
   * The address the compose window was opened for, captured from the deep link's
   * `mailtouri` as early as possible.
   *
   * It must be recorded here rather than read at fill time: Outlook rewrites the
   * compose URL once the SPA takes over, dropping the query string, so by the
   * time a template is applied the `mailtouri` is usually gone. Reading it late
   * meant the recipient check could never pass and cleanup was never offered.
   */
  intendedRecipient?: string;
}

async function readWorkflows(): Promise<Record<string, Workflow>> {
  const stored = await chrome.storage.session.get(KEY);
  return (stored[KEY] as Record<string, Workflow> | undefined) ?? {};
}

async function writeWorkflows(workflows: Record<string, Workflow>): Promise<void> {
  const now = Date.now();
  for (const [id, wf] of Object.entries(workflows)) {
    if (now - wf.createdAt > MAX_AGE_MS) delete workflows[id];
  }
  await chrome.storage.session.set({ [KEY]: workflows });
}

async function readOpeners(): Promise<Record<string, OpenerRecord>> {
  const stored = await chrome.storage.session.get(OPENERS);
  return (stored[OPENERS] as Record<string, OpenerRecord> | undefined) ?? {};
}

async function writeOpeners(openers: Record<string, OpenerRecord>): Promise<void> {
  // Age-prune, like workflows. Without this a record written for any tab ever
  // opened from an inbox tab lives until the tab closes, and can later be
  // attached to an unrelated compose action.
  const now = Date.now();
  for (const [id, record] of Object.entries(openers)) {
    if (now - record.createdAt > OPENER_MAX_AGE_MS) delete openers[id];
  }
  await chrome.storage.session.set({ [OPENERS]: openers });
}

function log(...args: unknown[]): void {
  console.log('[RecruiterAssistant/bg]', ...args);
}

/**
 * Records why cleanup was or was not offered, and tells the compose tab.
 *
 * Every reason used to live only in the service-worker console, which nobody
 * looks at — so "no prompt appeared" was indistinguishable from every possible
 * cause. The popup reads this, and the compose panel shows it while it is open.
 */
async function recordDecision(
  outcome: 'offered' | 'blocked',
  reason: string,
  composeTabId?: number,
): Promise<void> {
  await chrome.storage.session.set({
    [LAST_DECISION]: { outcome, reason, at: Date.now() },
  });
  log(`cleanup ${outcome}: ${reason}`);
  if (composeTabId != null) {
    try {
      await chrome.tabs.sendMessage(composeTabId, {
        type: 'RA_CLEANUP_STATUS',
        outcome,
        reason,
      });
    } catch {
      // The compose tab has already closed; the popup still has the reason.
    }
  }
}

/* ------------------------------------------------------------------ */
/* install                                                            */
/* ------------------------------------------------------------------ */

chrome.runtime.onInstalled.addListener(async (details) => {
  const stored = await chrome.storage.local.get('settings');
  if (!stored.settings) await chrome.storage.local.set({ settings: DEFAULT_SETTINGS });
  if (details.reason === 'install') void chrome.runtime.openOptionsPage();
});

/* ------------------------------------------------------------------ */
/* opener tracking                                                    */
/* ------------------------------------------------------------------ */

/**
 * Fires before the compose tab has finished loading, which is the only moment
 * the opener relationship is available. The snapshot request is retried because
 * the opener answers immediately but the new tab may not be an Outlook page yet.
 */
chrome.tabs.onCreated.addListener((tab) => {
  if (tab.id == null || tab.openerTabId == null) return;
  const composeTabId = tab.id;
  const sourceTabId = tab.openerTabId;
  // pendingUrl is the URL being navigated to; at onCreated time `url` is often
  // still blank, and this is the one moment the mailtouri is guaranteed present.
  const intended = parseMailtoRecipient(tab.pendingUrl ?? tab.url) ?? undefined;

  void (async () => {
    const snapshot = await requestSnapshot(sourceTabId);
    const openers = await readOpeners();
    const existing = openers[String(composeTabId)];
    openers[String(composeTabId)] = {
      sourceTabId,
      snapshot,
      createdAt: Date.now(),
      ...(intended ?? existing?.intendedRecipient
        ? { intendedRecipient: intended ?? existing?.intendedRecipient }
        : {}),
    };
    await writeOpeners(openers);
    log('linked compose tab', composeTabId, '→ source tab', sourceTabId, {
      messageId: snapshot?.messageId ? `${snapshot.messageId.slice(0, 12)}…` : null,
      looksLikeApplication: snapshot?.looksLikeApplication ?? false,
      haveIntendedRecipient: Boolean(intended),
    });
  })();
});

/**
 * A second chance to capture the mailtouri: the deep-link URL may only become
 * visible on the first navigation, before Outlook's SPA replaces it.
 */
chrome.tabs.onUpdated.addListener((tabId, changeInfo) => {
  const intended = parseMailtoRecipient(changeInfo.url);
  if (!intended) return;
  void (async () => {
    const openers = await readOpeners();
    const record = openers[String(tabId)];
    if (!record || record.intendedRecipient) return;
    record.intendedRecipient = intended;
    await writeOpeners(openers);
    log('captured the intended recipient for compose tab', tabId);
  })();
});

async function requestSnapshot(tabId: number): Promise<SourceSnapshot | null> {
  try {
    const reply = (await chrome.tabs.sendMessage(tabId, { type: 'RA_CAPTURE_SNAPSHOT' })) as
      | SnapshotReplyMsg
      | undefined;
    return reply?.snapshot ?? null;
  } catch {
    // Not an Outlook tab, or no content script there.
    return null;
  }
}

chrome.tabs.onRemoved.addListener((tabId) => {
  void (async () => {
    const openers = await readOpeners();
    if (openers[String(tabId)]) {
      delete openers[String(tabId)];
      await writeOpeners(openers);
    }

    // A compose tab closing after Send is itself a signal: Outlook closes the
    // deep-link tab once the message goes out.
    const workflows = await readWorkflows();
    const workflow = Object.values(workflows).find((wf) => wf.composeTabId === tabId);
    if (!workflow) return;

    if (workflow.state === 'SEND_CLICKED') {
      /*
       * Outlook closes a deep-link compose tab after a send, so a close is
       * suggestive — but it is NOT proof. Outlook can also reject the send
       * (offline, a policy block, an unresolved recipient) and leave the draft
       * open, and the user then closes the tab in frustration. The in-page
       * detector never gets to settle in that case.
       *
       * So this path applies the same freshness bound the detector uses, and
       * always asks for confirmation regardless of the setting, with wording
       * that says the send itself could not be observed.
       */
      const clickedAt = workflow.sendClickedAt ?? 0;
      const fresh = clickedAt > 0 && Date.now() - clickedAt < SEND_WINDOW_MS;
      if (!fresh) {
        await setState(workflow.id, 'SEND_ABANDONED');
        log('compose tab closed but the Send click was stale — source email left alone');
        return;
      }
      await setState(workflow.id, 'SEND_CONFIRMED');
      await offerDeletion(
        workflow.id,
        ['Send was clicked', 'the compose tab then closed — the send itself was not observed'],
        { forceConfirmation: true },
      );
    } else if (workflow.state === 'EMAIL_FILLED' || workflow.state === 'COMPOSE_DETECTED') {
      await setState(workflow.id, 'SEND_ABANDONED');
      log('compose tab closed without a Send click — source email left alone');
    }
  })();
});

/* ------------------------------------------------------------------ */
/* workflow state                                                     */
/* ------------------------------------------------------------------ */

async function setState(id: string, state: WorkflowState): Promise<Workflow | undefined> {
  const workflows = await readWorkflows();
  const workflow = workflows[id];
  if (!workflow) return undefined;
  workflow.state = state;
  await writeWorkflows(workflows);
  return workflow;
}

async function workflowForTab(tabId: number): Promise<Workflow | undefined> {
  const workflows = await readWorkflows();
  return Object.values(workflows).find((wf) => wf.composeTabId === tabId);
}

async function registerFill(
  tabId: number,
  composeUrl: string | undefined,
  windowId: number | undefined,
  message: RegisterFillMsg,
): Promise<void> {
  const openers = await readOpeners();
  const stored = openers[String(tabId)];
  const workflows = await readWorkflows();

  /*
   * An opener record alone must never authorise a deletion. It is written for
   * EVERY tab opened from an Outlook tab — ctrl-click, "open in new tab", a link
   * in a message body, a popped-out window — so on its own it only says "this
   * tab came from that tab at some point".
   *
   * The link is accepted only when the compose deep link's own `mailtouri` names
   * the very address this email is going to. That ties the recorded source
   * message to *this* candidate, and it rules out two ways of deleting the wrong
   * email: a tab that was opened from the inbox for an unrelated reason, and a
   * recipient edited by hand after the deep link opened.
   */
  // Prefer the recipient recorded when the tab was created; fall back to the
  // live URL for the case where the query string is somehow still present.
  const intendedRecipient = stored?.intendedRecipient ?? parseMailtoRecipient(composeUrl);
  const candidate = message.candidateEmail.toLowerCase();
  let opener: OpenerRecord | undefined;
  let linkRefusedBecause: string | undefined;

  if (!stored) {
    linkRefusedBecause = 'this tab was not opened from another Outlook tab';
  } else if (!intendedRecipient) {
    linkRefusedBecause =
      'the address this compose window was opened for could not be determined, so the source email cannot be matched to it';
  } else if (intendedRecipient !== candidate) {
    linkRefusedBecause = 'the recipient differs from the address the compose window was opened for';
  } else if (Date.now() - stored.createdAt > OPENER_MAX_AGE_MS) {
    linkRefusedBecause = 'the recorded source message is too old to trust';
  } else {
    opener = stored;
  }

  // Keyed by compose tab + candidate, so re-filling the same compose window
  // updates one workflow instead of accumulating duplicates.
  const id = `${tabId}:${message.candidateEmail.toLowerCase()}`;
  const existing = workflows[id];

  workflows[id] = {
    id,
    composeTabId: tabId,
    ...(windowId != null ? { composeWindowId: windowId } : {}),
    ...(opener?.sourceTabId != null ? { sourceTabId: opener.sourceTabId, sourceVia: 'opener' } : {}),
    ...(opener?.snapshot ? { sourceSnapshot: opener.snapshot } : {}),
    candidateEmail: message.candidateEmail,
    actionLabel: message.actionLabel,
    state: 'EMAIL_FILLED',
    createdAt: existing?.createdAt ?? Date.now(),
  };

  // Drop any other workflow for this tab: the user changed their mind and picked
  // a different action, and two live workflows for one compose window could
  // offer two deletions.
  for (const [key, wf] of Object.entries(workflows)) {
    if (wf.composeTabId === tabId && key !== id) delete workflows[key];
  }

  await writeWorkflows(workflows);
  log('fill registered', {
    action: message.actionLabel,
    sourceTabId: opener?.sourceTabId ?? null,
    haveSnapshot: Boolean(opener?.snapshot),
    ...(linkRefusedBecause ? { noCleanupBecause: linkRefusedBecause } : {}),
  });
}

/* ------------------------------------------------------------------ */
/* deletion offer                                                     */
/* ------------------------------------------------------------------ */

/**
 * Used when Chrome reported no opener for the compose tab — which happens if
 * Outlook opens the deep link with `noopener`. Every other Outlook tab in the same
 * window is asked what it has open, and this only proceeds when **exactly one**
 * reports a LinkedIn application. Two candidates, or none, means abort.
 *
 * This is weaker evidence than an opener link, so it always asks for confirmation.
 */
async function findSoleApplicationTab(
  workflow: Workflow,
): Promise<{ tabId: number; snapshot: SourceSnapshot } | null> {
  let tabs: chrome.tabs.Tab[];
  try {
    tabs = await chrome.tabs.query(
      workflow.composeWindowId != null ? { windowId: workflow.composeWindowId } : {},
    );
  } catch {
    return null;
  }

  const matches: Array<{ tabId: number; snapshot: SourceSnapshot }> = [];
  for (const tab of tabs) {
    if (tab.id == null || tab.id === workflow.composeTabId) continue;
    const snapshot = await requestSnapshot(tab.id);
    if (snapshot?.looksLikeApplication) matches.push({ tabId: tab.id, snapshot });
  }

  if (matches.length !== 1) {
    log(
      `fallback source search found ${matches.length} candidate tabs — need exactly one, so nothing will be deleted`,
    );
    return null;
  }
  return matches[0] ?? null;
}

async function offerDeletion(
  workflowId: string,
  signals: string[],
  options: { forceConfirmation?: boolean } = {},
): Promise<void> {
  const settings = await loadSettings();
  const workflows = await readWorkflows();
  const workflow = workflows[workflowId];
  if (!workflow) return;

  if (workflow.state === 'DELETE_OFFERED' || workflow.state === 'DELETED') {
    log('a deletion was already offered for this workflow — not offering again');
    return;
  }

  if (!settings.attemptSourceDeletion) {
    await recordDecision('blocked', 'cleanup is switched off in Settings', workflow.composeTabId);
    return;
  }
  // No opener link, or the opener had nothing useful open: fall back to finding a
  // single unambiguous application tab in the same window.
  let extraSignals = signals;
  if (workflow.sourceTabId == null || !workflow.sourceSnapshot?.looksLikeApplication) {
    const sole = await findSoleApplicationTab(workflow);
    if (sole) {
      const all = await readWorkflows();
      const current = all[workflowId];
      if (current) {
        current.sourceTabId = sole.tabId;
        current.sourceSnapshot = sole.snapshot;
        current.sourceVia = 'sole-application-tab';
        await writeWorkflows(all);
        workflow.sourceTabId = sole.tabId;
        workflow.sourceSnapshot = sole.snapshot;
        workflow.sourceVia = 'sole-application-tab';
      }
      extraSignals = [
        ...signals,
        'the source email was matched by being the only LinkedIn application open in this window',
      ];
      options = { ...options, forceConfirmation: true };
    }
  }

  if (workflow.sourceTabId == null) {
    await setState(workflowId, 'DELETE_UNSAFE');
    await recordDecision(
      'blocked',
      'this compose tab could not be tied to an inbox tab: Chrome reported no opener, and no single LinkedIn application tab was open in this window',
      workflow.composeTabId,
    );
    return;
  }
  if (!workflow.sourceSnapshot) {
    await setState(workflowId, 'DELETE_UNSAFE');
    await recordDecision(
      'blocked',
      'the inbox tab did not report an open message',
      workflow.composeTabId,
    );
    return;
  }
  if (!workflow.sourceSnapshot.looksLikeApplication) {
    await setState(workflowId, 'DELETE_UNSAFE');
    await recordDecision(
      'blocked',
      `the source message did not read as a LinkedIn application (sender "${
        workflow.sourceSnapshot.sender || 'unreadable'
      }", subject "${workflow.sourceSnapshot.subject || 'unreadable'}")`,
      workflow.composeTabId,
    );
    return;
  }

  const offer: OfferDeletionMsg = {
    type: 'RA_OFFER_DELETION',
    workflowId,
    candidateEmail: workflow.candidateEmail,
    actionLabel: workflow.actionLabel,
    snapshot: workflow.sourceSnapshot,
    requireConfirmation: options.forceConfirmation || settings.requireDeletionConfirmation,
    signals: extraSignals,
  };

  try {
    await chrome.tabs.sendMessage(workflow.sourceTabId, offer);
    await setState(workflowId, 'DELETE_OFFERED');
    await recordDecision(
      'offered',
      `prompt shown in the inbox tab for "${workflow.sourceSnapshot.subject}"`,
      workflow.composeTabId,
    );
  } catch {
    await setState(workflowId, 'DELETE_UNSAFE');
    await recordDecision(
      'blocked',
      'the inbox tab is no longer open, so there was nowhere to show the prompt',
      workflow.composeTabId,
    );
  }
}

/* ------------------------------------------------------------------ */
/* message routing                                                    */
/* ------------------------------------------------------------------ */

chrome.runtime.onMessage.addListener((message, sender, respond) => {
  const type = messageType(message);
  const tabId = sender.tab?.id;

  switch (type) {
    case 'RA_REGISTER_FILL': {
      if (tabId == null) return false;
      void registerFill(
        tabId,
        sender.tab?.url,
        sender.tab?.windowId,
        message as RegisterFillMsg,
      ).then(() => respond({ ok: true }));
      return true;
    }

    case 'RA_SEND_CLICKED': {
      if (tabId == null) return false;
      void (async () => {
        const workflows = await readWorkflows();
        const workflow = Object.values(workflows).find((wf) => wf.composeTabId === tabId);
        if (workflow) {
          workflow.state = 'SEND_CLICKED';
          workflow.sendClickedAt = Date.now();
          await writeWorkflows(workflows);
        }
        respond({ ok: Boolean(workflow) });
      })();
      return true;
    }

    case 'RA_SEND_CONFIRMED': {
      if (tabId == null) return false;
      void (async () => {
        const workflow = await workflowForTab(tabId);
        if (workflow) {
          await setState(workflow.id, 'SEND_CONFIRMED');
          await offerDeletion(workflow.id, (message as SendConfirmedMsg).signals);
        }
        respond({ ok: Boolean(workflow) });
      })();
      return true;
    }

    case 'RA_DISCARD_CLICKED': {
      if (tabId == null) return false;
      void (async () => {
        // A discard is a hard veto: it must be recorded even if the tab is torn
        // down before the in-page detector's debounce can settle.
        const workflow = await workflowForTab(tabId);
        if (workflow) await setState(workflow.id, 'SEND_ABANDONED');
        log('discard clicked — source email left alone');
        respond({ ok: true });
      })();
      return true;
    }

    case 'RA_SEND_ABANDONED': {
      if (tabId == null) return false;
      void (async () => {
        const workflow = await workflowForTab(tabId);
        if (workflow) await setState(workflow.id, 'SEND_ABANDONED');
        await recordDecision(
          'blocked',
          `no send was confirmed (${(message as SendAbandonedMsg).reason})`,
          tabId,
        );
        respond({ ok: true });
      })();
      return true;
    }

    case 'RA_LAST_DECISION': {
      void (async () => {
        const stored = await chrome.storage.session.get(LAST_DECISION);
        respond({ decision: stored[LAST_DECISION] ?? null });
      })();
      return true;
    }

    case 'RA_DELETE_RESULT': {
      void (async () => {
        const { workflowId, ok, detail } = message as {
          workflowId: string;
          ok: boolean;
          detail: string;
        };
        await setState(workflowId, ok ? 'DELETED' : 'DELETE_UNSAFE');
        log('delete result', ok ? 'ok' : 'refused', detail);
        respond({ ok: true });
      })();
      return true;
    }

    case 'RA_DELETE_DECLINED': {
      void (async () => {
        await setState((message as { workflowId: string }).workflowId, 'DELETE_DECLINED');
        respond({ ok: true });
      })();
      return true;
    }

    default:
      return false;
  }
});
