import test from 'node:test';
import assert from 'node:assert/strict';
import { loadFixture, lib } from './dom-setup.mjs';

const libSync = await lib();
const {
  decideOutcome,
  judgeApplication,
  readMessageIdFromUrl,
  verifySourceStillMatches,
  captureSnapshot,
  DEFAULT_SETTINGS,
  renderTemplate,
  htmlToText,
} = libSync;

const NOW = 1_000_000;

function outcome(patch = {}) {
  return decideOutcome({
    sawComposeSurface: true,
    composeStillPresent: false,
    sendClickedAt: 0,
    discardClickedAt: 0,
    now: NOW,
    ...patch,
  });
}

/* ------------------------------------------------------------------ */
/* send detection — the gate on every deletion                         */
/* ------------------------------------------------------------------ */

test('a Send click followed by the compose window disappearing counts as sent', () => {
  const result = outcome({ sendClickedAt: NOW - 1000 });
  assert.equal(result.confirmed, true);
  assert.ok(result.signals.length >= 2);
});

test('the compose window disappearing with no Send click is NOT a send', () => {
  const result = outcome();
  assert.equal(result.confirmed, false);
  assert.match(result.reason, /without Send being clicked/);
});

test('a discard after Send is NOT a send', () => {
  const result = outcome({ sendClickedAt: NOW - 5000, discardClickedAt: NOW - 1000 });
  assert.equal(result.confirmed, false);
  assert.match(result.reason, /discarded/);
});

test('a discard before Send does not veto the send', () => {
  // Discarding one draft then sending a different one must still count.
  const result = outcome({ discardClickedAt: NOW - 9000, sendClickedAt: NOW - 1000 });
  assert.equal(result.confirmed, true);
});

test('a stale Send click does not confirm a much later disappearance', () => {
  const result = outcome({ sendClickedAt: NOW - 10 * 60 * 1000 });
  assert.equal(result.confirmed, false);
  assert.match(result.reason, /without Send being clicked/);
});

test('nothing is concluded while the compose window is still present', () => {
  assert.equal(outcome({ composeStillPresent: true, sendClickedAt: NOW - 100 }), null);
});

test('nothing is concluded if a compose surface was never seen', () => {
  assert.equal(outcome({ sawComposeSurface: false, sendClickedAt: NOW - 100 }), null);
});

/* ------------------------------------------------------------------ */
/* source-email identification                                         */
/* ------------------------------------------------------------------ */

test('an application email needs BOTH a LinkedIn sender and an application subject', () => {
  assert.equal(judgeApplication('LinkedIn', 'New application: AI Full Stack Software Engineer'), true);
  assert.equal(judgeApplication('LinkedIn', 'Your weekly job digest'), false);
  assert.equal(judgeApplication('recruiter@acme.com', 'New application: Engineer'), false);
  assert.equal(judgeApplication('', ''), false);
});

test('a colleague merely mentioning LinkedIn is not an application email', () => {
  assert.equal(judgeApplication('venkat@cloudsecurityweb.com', 'Check this LinkedIn profile'), false);
});

test('the message id is read from the Outlook reading-pane URL', () => {
  assert.equal(
    readMessageIdFromUrl('https://outlook.cloud.microsoft/mail/inbox/id/AAQkADNmNjUwMWQ4LTA0NGQ'),
    'AAQkADNmNjUwMWQ4LTA0NGQ',
  );
  assert.equal(readMessageIdFromUrl('https://outlook.cloud.microsoft/mail/'), null);
  assert.equal(readMessageIdFromUrl('nonsense'), null);
});

test('a snapshot of a LinkedIn application email is recognised as one', () => {
  loadFixture('mock-reading-pane.html', 'https://outlook.cloud.microsoft/mail/inbox/id/MSG-1');
  const snapshot = captureSnapshot();
  assert.equal(snapshot.messageId, 'MSG-1');
  assert.equal(snapshot.sender, 'LinkedIn');
  assert.match(snapshot.subject, /New application/);
  assert.equal(snapshot.looksLikeApplication, true);
});

test('no snapshot is produced when no message is open', () => {
  loadFixture('mock-reading-pane.html', 'https://outlook.cloud.microsoft/mail/');
  assert.equal(captureSnapshot(), null);
});

test('deletion is refused when a different message is now open', () => {
  loadFixture('mock-reading-pane.html', 'https://outlook.cloud.microsoft/mail/inbox/id/MSG-2');
  const verdict = verifySourceStillMatches({
    messageId: 'MSG-1',
    sender: 'LinkedIn',
    subject: 'New application: AI Full Stack Software Engineer',
    looksLikeApplication: true,
    capturedAt: NOW,
  });
  assert.equal(verdict.ok, false);
  assert.match(verdict.detail, /different email/);
});

test('deletion is refused when no message is open any more', () => {
  loadFixture('mock-reading-pane.html', 'https://outlook.cloud.microsoft/mail/');
  const verdict = verifySourceStillMatches({
    messageId: 'MSG-1',
    sender: 'LinkedIn',
    subject: 'New application: AI Full Stack Software Engineer',
    looksLikeApplication: true,
    capturedAt: NOW,
  });
  assert.equal(verdict.ok, false);
});

test('deletion is refused when the open email stopped looking like an application', () => {
  loadFixture('mock-reading-pane.html', 'https://outlook.cloud.microsoft/mail/inbox/id/MSG-1');
  document.getElementById('sender').setAttribute('aria-label', 'Venkat <venkat@cloudsecurityweb.com>');
  document.getElementById('sender').textContent = 'Venkat';
  document.getElementById('message-subject').textContent = 'Lunch tomorrow?';
  const verdict = verifySourceStillMatches({
    messageId: 'MSG-1',
    sender: 'LinkedIn',
    subject: 'New application: AI Full Stack Software Engineer',
    looksLikeApplication: true,
    capturedAt: NOW,
  });
  assert.equal(verdict.ok, false);
  assert.match(verdict.detail, /no longer reads as a LinkedIn application/);
});

test('deletion is allowed when message id, sender and subject all still match', () => {
  loadFixture('mock-reading-pane.html', 'https://outlook.cloud.microsoft/mail/inbox/id/MSG-1');
  const snapshot = captureSnapshot();
  const verdict = verifySourceStillMatches(snapshot);
  assert.equal(verdict.ok, true, verdict.detail);
});

/* ------------------------------------------------------------------ */
/* template spacing                                                    */
/* ------------------------------------------------------------------ */

test('both rejection templates put a blank line between paragraphs', () => {
  for (const id of ['reject-direct', 'reject-in-process']) {
    const template = DEFAULT_SETTINGS.templates[id];
    assert.ok(
      template.bodyHtml.includes('<p>&nbsp;</p>'),
      `${id} has no blank-line spacer between paragraphs`,
    );
  }
});

test('blank-line spacers survive rendering and are not stripped', () => {
  const out = renderTemplate(
    DEFAULT_SETTINGS.templates['reject-direct'],
    { email: 'a@b.com' },
    DEFAULT_SETTINGS,
  );
  assert.equal(out.ok, true, out.errors.join('; '));
  const spacers = out.bodyHtml.match(/<p>&nbsp;<\/p>/g) ?? [];
  assert.equal(spacers.length, 3, 'expected one spacer between each pair of paragraphs');
});

test('an empty paragraph left by a suppressed variable is still removed', () => {
  const template = {
    ...DEFAULT_SETTINGS.templates['reject-direct'],
    bodyHtml: '<p>{{GREETING}}</p><p>&nbsp;</p><p>Body.</p>',
  };
  const out = renderTemplate(template, { email: 'a@b.com' }, { ...DEFAULT_SETTINGS, greetingMode: 'none' });
  assert.ok(!out.bodyHtml.includes('<p></p>'), 'the emptied greeting paragraph should be gone');
  assert.ok(out.bodyHtml.includes('<p>&nbsp;</p>'), 'the intentional spacer should remain');
});

test('the plain-text rendering keeps paragraphs readable', () => {
  const out = renderTemplate(
    DEFAULT_SETTINGS.templates['reject-in-process'],
    { email: 'a@b.com' },
    DEFAULT_SETTINGS,
  );
  assert.ok(!/\n{3,}/.test(out.bodyText), 'no runs of three or more newlines');
  assert.match(htmlToText(out.bodyHtml), /Best regards,/);
});

/* ------------------------------------------------------------------ */
/* regressions from the adversarial audit                             */
/* ------------------------------------------------------------------ */

test('the opener link is only accepted when the compose URL names this candidate', () => {
  const { parseMailtoRecipient } = libSync;
  const url =
    'https://outlook.cloud.microsoft/mail/deeplink/compose?mailtouri=mailto%3Aesrisanjay2005%40gmail.com';

  // The exact check the service worker applies before it will attach a source
  // snapshot to a workflow.
  assert.equal(parseMailtoRecipient(url), 'esrisanjay2005@gmail.com');
  // A tab opened from the inbox for any other reason carries no mailtouri, so the
  // recorded source message can never authorise a deletion.
  assert.equal(parseMailtoRecipient('https://outlook.cloud.microsoft/mail/inbox/id/MSG-1'), null);
  // A hand-edited recipient no longer matches what the compose was opened for.
  assert.notEqual(parseMailtoRecipient(url), 'someone.else@example.com');
});

test('the sender is not read from the message body', () => {
  loadFixture('mock-reading-pane.html', 'https://outlook.cloud.microsoft/mail/inbox/id/MSG-1');
  // A colleague's email whose *body* mentions LinkedIn must not read as being
  // from LinkedIn — that plus a matching subject would qualify it for deletion.
  document.getElementById('sender').setAttribute('aria-label', 'Venkat <venkat@cloudsecurityweb.com>');
  document.getElementById('sender').removeAttribute('title');
  document.getElementById('sender').textContent = 'Venkat';
  document.getElementById('message-body').textContent =
    'Have a look at this LinkedIn profile for the new application';

  const snapshot = captureSnapshot();
  assert.ok(!/linkedin/i.test(snapshot.sender), `sender leaked from the body: ${snapshot.sender}`);
  assert.equal(snapshot.looksLikeApplication, false);
});

test('no reading pane means no snapshot, rather than reading the message list', () => {
  loadFixture('mock-reading-pane.html', 'https://outlook.cloud.microsoft/mail/inbox/id/MSG-1');
  document.getElementById('message-container').remove();
  const snapshot = captureSnapshot();
  assert.equal(snapshot.sender, '');
  assert.equal(snapshot.looksLikeApplication, false);
});

/* ------------------------------------------------------------------ */
/* keyboard send — the way this workflow is actually driven            */
/* ------------------------------------------------------------------ */

test('a Cmd+Enter send is confirmed the same as a Send click', () => {
  const result = outcome({ sendClickedAt: NOW - 800, sendVia: 'Cmd+Enter was pressed' });
  assert.equal(result.confirmed, true);
  assert.ok(
    result.signals.some((s) => /Cmd\+Enter/.test(s)),
    'the prompt should say how the send was triggered',
  );
});

test('a discard still vetoes a keyboard send', () => {
  const result = outcome({
    sendClickedAt: NOW - 4000,
    sendVia: 'Cmd+Enter was pressed',
    discardClickedAt: NOW - 500,
  });
  assert.equal(result.confirmed, false);
  assert.match(result.reason, /discarded/);
});

test('signals name the trigger, falling back to a neutral phrase', () => {
  const result = outcome({ sendClickedAt: NOW - 500 });
  assert.match(result.signals[0], /a send was triggered/);
});

/* ------------------------------------------------------------------ */
/* reading-pane fallback — why no prompt ever appeared                 */
/* ------------------------------------------------------------------ */

test('a snapshot is still taken when the tenant has no role=document', () => {
  loadFixture('mock-reading-pane.html', 'https://outlook.cloud.microsoft/mail/inbox/id/MSG-1');
  // Flatten role=document away, leaving only role=main — requiring role=document
  // meant no snapshot, which silently meant cleanup never ran.
  const container = document.getElementById('message-container');
  const parent = container.parentElement;
  while (container.firstChild) parent.appendChild(container.firstChild);
  container.remove();
  parent.setAttribute('role', 'main');
  parent.removeAttribute('aria-label');

  const snapshot = captureSnapshot();
  assert.equal(snapshot.looksLikeApplication, true, `pane resolved via: ${snapshot.paneVia}`);
  assert.equal(snapshot.sender, 'LinkedIn');
  assert.match(snapshot.paneVia, /role=main/);
});

test('the subject is never read out of the message list', () => {
  loadFixture('mock-reading-pane.html', 'https://outlook.cloud.microsoft/mail/inbox/id/MSG-1');
  const container = document.getElementById('message-container');
  const parent = container.parentElement;
  while (container.firstChild) parent.appendChild(container.firstChild);
  container.remove();
  parent.setAttribute('role', 'main');
  parent.removeAttribute('aria-label');

  // A longer heading belonging to a DIFFERENT message, inside the list.
  const list = document.createElement('div');
  list.setAttribute('role', 'listbox');
  const other = document.createElement('div');
  other.setAttribute('role', 'heading');
  other.textContent = 'New application: A Much Longer Subject Belonging To Somebody Else Entirely';
  list.append(other);
  parent.prepend(list);

  const snapshot = captureSnapshot();
  // The fixture's own subject, not the longer one planted in the list.
  assert.equal(snapshot.subject, 'New application: Software Engineer');
  assert.ok(!/Somebody Else/.test(snapshot.subject), 'read the wrong message from the list');
});

test('the intended recipient is parsed from the deep link, not the rewritten URL', () => {
  const { parseMailtoRecipient } = libSync;
  // Captured at tab creation from the deep link...
  assert.equal(
    parseMailtoRecipient(
      'https://outlook.cloud.microsoft/mail/deeplink/compose?mailtouri=mailto%3Aesrisanjay2005%40gmail.com',
    ),
    'esrisanjay2005@gmail.com',
  );
  // ...because by fill time Outlook has replaced it with this, and reading it
  // then blocked every cleanup.
  assert.equal(parseMailtoRecipient('https://outlook.cloud.microsoft/mail/0/deeplink/compose'), null);
});

/* ------------------------------------------------------------------ */
/* stored-settings migration — why the blank lines never appeared      */
/* ------------------------------------------------------------------ */

test('a stored template whose wording matches is upgraded to the shipped spacing', () => {
  const { mergeStoredSettings } = libSync;
  const shipped = DEFAULT_SETTINGS.templates['reject-direct'];
  // Exactly what an earlier build saved: same words, no blank-line spacers.
  const stale = shipped.bodyHtml.replace(/<p>&nbsp;<\/p>\n?/g, '');
  assert.ok(!stale.includes('&nbsp;'));

  const merged = mergeStoredSettings({
    templates: { 'reject-direct': { ...shipped, bodyHtml: stale } },
  });
  assert.ok(
    merged.templates['reject-direct'].bodyHtml.includes('<p>&nbsp;</p>'),
    'stored copy should have been upgraded to the spaced default',
  );
});

test('a genuinely edited template is never overwritten', () => {
  const { mergeStoredSettings } = libSync;
  const shipped = DEFAULT_SETTINGS.templates['reject-direct'];
  const mine = '<p>My own wording entirely.</p>';
  const merged = mergeStoredSettings({
    templates: { 'reject-direct': { ...shipped, bodyHtml: mine } },
  });
  assert.equal(merged.templates['reject-direct'].bodyHtml, mine);
});

test('an unreviewed placeholder is replaced by confirmed copy', () => {
  const { mergeStoredSettings } = libSync;
  const merged = mergeStoredSettings({
    templates: {
      'shortlist-abhi': {
        ...DEFAULT_SETTINGS.templates['shortlist-abhi'],
        bodyHtml: '<p>Old placeholder wording.</p>',
        isPlaceholder: true,
      },
    },
  });
  assert.match(merged.templates['shortlist-abhi'].bodyHtml, /technical screening over Microsoft Teams with/);
  assert.equal(merged.templates['shortlist-abhi'].isPlaceholder, false);
});

test('a Bookings link the user set is kept through an upgrade', () => {
  const { mergeStoredSettings } = libSync;
  const shipped = DEFAULT_SETTINGS.templates['shortlist-abhi'];
  const mineUrl = 'https://outlook.office.com/bookwithme/user/mine/meet';
  const merged = mergeStoredSettings({
    templates: {
      'shortlist-abhi': { ...shipped, bodyHtml: '<p>Old placeholder.</p>', isPlaceholder: true, bookingUrl: mineUrl },
    },
  });
  assert.equal(merged.templates['shortlist-abhi'].bookingUrl, mineUrl);
});

/* ------------------------------------------------------------------ */
/* the resume preview open beside the message                          */
/* ------------------------------------------------------------------ */

test('the resume preview is not mistaken for the message', () => {
  loadFixture('mock-reading-pane.html', 'https://outlook.cloud.microsoft/mail/inbox/id/MSG-1');

  // Exactly the real situation: the PDF preview is still open in this tab, and it
  // presents its own document surface *before* the message in document order.
  const preview = document.createElement('div');
  preview.setAttribute('role', 'document');
  preview.innerHTML =
    '<div>Sri_Sanjay_E_Resume.pdf <button>Download</button><button>Print</button>' +
    '<button>Save to OneDrive</button></div>' +
    '<div role="heading">SRI SANJAY E — Full-Stack Developer, a very long resume heading indeed</div>' +
    '<p>esrisanjay2005@gmail.com</p>';
  document.body.prepend(preview);

  const snapshot = captureSnapshot();
  assert.equal(
    snapshot.looksLikeApplication,
    true,
    `read the wrong surface — sender "${snapshot.sender}", subject "${snapshot.subject}", via ${snapshot.paneVia}`,
  );
  assert.equal(snapshot.sender, 'LinkedIn');
  assert.match(snapshot.subject, /New application/);
  assert.ok(!/SRI SANJAY/.test(snapshot.subject), 'subject came from the resume');
});

test('verification also ignores the resume preview', () => {
  loadFixture('mock-reading-pane.html', 'https://outlook.cloud.microsoft/mail/inbox/id/MSG-1');
  const preview = document.createElement('div');
  preview.setAttribute('role', 'document');
  preview.innerHTML =
    '<div>CV.pdf <button>Download</button><button>Print</button></div>' +
    '<div role="heading">Some resume heading that is longer than the real subject line</div>';
  document.body.prepend(preview);

  const snapshot = captureSnapshot();
  const verdict = verifySourceStillMatches(snapshot);
  assert.equal(verdict.ok, true, verdict.detail);
});

/* ------------------------------------------------------------------ */
/* keyboard shortcuts                                                  */
/* ------------------------------------------------------------------ */

test('every action has a distinct Option+digit slot, in panel order', () => {
  // The panel labels buttons ⌥1..⌥4 from ACTION_ORDER, so the order is the
  // contract: reordering ACTION_ORDER silently remaps the user's shortcuts.
  const { ACTION_ORDER } = libSync;
  assert.deepEqual(ACTION_ORDER, [
    'reject-direct',
    'reject-in-process',
    'shortlist-ravilochan',
    'shortlist-abhi',
  ]);
  assert.equal(new Set(ACTION_ORDER).size, ACTION_ORDER.length);
  assert.ok(ACTION_ORDER.length <= 9, 'more than nine actions would exceed the digit keys');
});
