
/**
 * Status popup. Its job is to distinguish the reasons the panel can be absent,
 * which are otherwise indistinguishable from "nothing happened":
 *
 *   1. the tab is not an Outlook host the manifest matches
 *   2. Chrome has site access set to something other than "on all sites"
 *   3. the extension or the tab needs reloading
 *   4. the tab simply has no compose window with one recipient
 */

const SUPPORTED_HOSTS = [
  'outlook.cloud.microsoft',
  'outlook.office.com',
  'outlook.office365.com',
  'outlook.microsoft.com',
];

interface Snapshot {
  messageId: string;
  sender: string;
  subject: string;
  looksLikeApplication: boolean;
  paneVia?: string;
}

interface Pong {
  type: 'RA_PONG';
  version: string;
  host: string;
  path: string;
  topFrame: boolean;
  assistantMounted: boolean;
  composeDetected: boolean;
  candidateEmail?: string;
  sourceMessageOpen: boolean;
  snapshot?: Snapshot | null;
}

function line(key: string, value: string, cls: '' | 'ok' | 'warn' | 'err' = ''): HTMLElement {
  const row = document.createElement('div');
  row.className = 'line';
  const k = document.createElement('span');
  k.className = 'k';
  k.textContent = key;
  const v = document.createElement('span');
  v.className = `v ${cls}`.trim();
  v.textContent = value;
  row.append(k, v);
  return row;
}

function fixBlock(title: string, steps: string[]): HTMLElement {
  const box = document.createElement('div');
  box.className = 'fix';
  const strong = document.createElement('strong');
  strong.textContent = title;
  box.append(strong);
  const ol = document.createElement('ol');
  for (const step of steps) {
    const li = document.createElement('li');
    li.textContent = step;
    ol.append(li);
  }
  box.append(ol);
  return box;
}

/** The service worker's record of why cleanup last happened, or did not. */
async function appendLastDecision(report: HTMLElement): Promise<void> {
  try {
    const reply = (await chrome.runtime.sendMessage({ type: 'RA_LAST_DECISION' })) as
      | { decision: { outcome: 'offered' | 'blocked'; reason: string; at: number } | null }
      | undefined;
    const decision = reply?.decision;
    if (!decision) return;
    const mins = Math.round((Date.now() - decision.at) / 60000);
    report.append(
      line(
        'Last cleanup',
        `${decision.outcome === 'offered' ? 'prompt shown' : 'blocked'} · ${
          mins < 1 ? 'just now' : `${mins}m ago`
        }`,
        decision.outcome === 'offered' ? 'ok' : 'warn',
      ),
    );
    report.append(line('· because', decision.reason, decision.outcome === 'offered' ? '' : 'warn'));
  } catch {
    /* worker asleep or no decision recorded yet */
  }
}

async function activeTab(): Promise<chrome.tabs.Tab | undefined> {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab;
}

async function ping(tabId: number): Promise<Pong | null> {
  try {
    const response = (await chrome.tabs.sendMessage(tabId, { type: 'RA_PING' })) as Pong | undefined;
    return response ?? null;
  } catch {
    // No receiver = the content script is not running in that tab.
    return null;
  }
}

async function render(): Promise<void> {
  const report = document.getElementById('report');
  if (!report) return;
  report.replaceChildren();

  const manifest = chrome.runtime.getManifest();
  report.append(line('Version', manifest.version));

  const tab = await activeTab();
  if (!tab?.id) {
    report.append(line('Tab', 'could not read the active tab', 'err'));
    return;
  }

  // tab.url is only visible on hosts the extension has permission for, which is
  // itself a useful signal.
  let host = '';
  try {
    host = tab.url ? new URL(tab.url).host : '';
  } catch {
    host = '';
  }

  const supported = SUPPORTED_HOSTS.includes(host);
  if (host) {
    report.append(line('Tab host', host, supported ? 'ok' : 'warn'));
  } else {
    report.append(line('Tab host', 'hidden — no permission for this tab', 'warn'));
  }

  const pong = await ping(tab.id);

  if (pong) {
    report.append(line('Content script', `running (${pong.topFrame ? 'top frame' : 'sub-frame'})`, 'ok'));
    report.append(line('Path', pong.path));
    report.append(
      line(
        'Recruiter panel',
        pong.assistantMounted ? 'showing' : 'no compose detected on this tab',
        pong.assistantMounted ? 'ok' : 'warn',
      ),
    );
    if (pong.candidateEmail) report.append(line('Candidate', pong.candidateEmail, 'ok'));

    // What this tab would report if asked to clean up — the single most useful
    // thing to see when no prompt appeared.
    if (pong.snapshot) {
      const snap = pong.snapshot;
      report.append(
        line(
          'As a source tab',
          snap.looksLikeApplication
            ? 'recognised as a LinkedIn application'
            : 'NOT recognised as a LinkedIn application',
          snap.looksLikeApplication ? 'ok' : 'err',
        ),
      );
      report.append(line('· sender read', snap.sender || '(nothing found)', snap.sender ? '' : 'err'));
      report.append(line('· subject read', snap.subject || '(nothing found)', snap.subject ? '' : 'err'));
      if (snap.paneVia) report.append(line('· reading pane', snap.paneVia));
    } else if (pong.path.includes('/id/')) {
      report.append(line('As a source tab', 'no snapshot could be taken', 'err'));
    }

    await appendLastDecision(report);
    if (!pong.assistantMounted) {
      report.append(
        fixBlock('The recruiter panel only shows on a compose window', [
          'Open a compose window with the candidate in the To field — clicking their address in the resume preview does this.',
          'It needs exactly one recipient, so there is no doubt who the email is for.',
          pong.composeDetected
            ? 'A compose surface was found here, so the recipient is what is missing.'
            : 'No compose surface was found on this tab.',
        ]),
      );
    }
    return;
  }

  report.append(line('Content script', 'NOT running in this tab', 'err'));

  if (!supported) {
    report.append(
      fixBlock('This tab is not a supported Outlook host', [
        `The extension only runs on: ${SUPPORTED_HOSTS.join(', ')}.`,
        host
          ? `This tab is ${host}. Add it to manifest.json (host_permissions and content_scripts matches) and rebuild if it should be supported.`
          : 'Open your Outlook tab and check again.',
      ]),
    );
    return;
  }

  report.append(
    fixBlock('Injected nowhere on a supported host — usually one of these', [
      'Right-click the extension icon → check that site access is "On all sites", not "On click".',
      'chrome://extensions → Reload on Recruiter Assistant (a manifest change needs a reload).',
      'chrome://extensions → check for an "Errors" button on the extension card.',
      'Reload this Outlook tab — a content script only injects on page load.',
    ]),
  );
}

document.getElementById('options')?.addEventListener('click', () => {
  void chrome.runtime.openOptionsPage();
});

document.getElementById('recheck')?.addEventListener('click', () => {
  void render();
});

void render();
