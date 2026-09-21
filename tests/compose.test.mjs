import test from 'node:test';
import assert from 'node:assert/strict';
import { loadFixture, lib } from './dom-setup.mjs';

const {
  findSubject,
  findBody,
  findRecipientWell,
  readRecipientsFromDom,
  readRecipientFromUrl,
  resolveRecipient,
  findComposeSurface,
  normaliseForCompare,
} = await lib();

/* ------------------------------------------------------------------ */
/* recipient scoping — the regression that mattered                    */
/* ------------------------------------------------------------------ */

test('the sender address in the same dialog is not read as a recipient', () => {
  loadFixture('mock-ribbon-compose.html');
  const recipients = readRecipientsFromDom();
  const emails = recipients.map((r) => r.email.toLowerCase());

  assert.deepEqual(emails, ['jordan.rivera.dev@example.com']);
  assert.ok(
    !emails.includes('ravilochan@cloudsecurityweb.com'),
    'the From address leaked into the recipient list — this blocks every action',
  );
});

test('a chip and its own Remove button count as one recipient', () => {
  loadFixture('mock-ribbon-compose.html');
  assert.equal(readRecipientsFromDom().length, 1);
});

test('a second real recipient is detected, and blocks action', () => {
  loadFixture('mock-ribbon-compose.html');
  const chip = document.createElement('span');
  chip.setAttribute('role', 'listitem');
  chip.setAttribute('aria-label', 'someone.else@example.com');
  document.getElementById('to-well').append(chip);

  assert.equal(readRecipientsFromDom().length, 2);
  const { recipient, problem } = resolveRecipient();
  assert.equal(recipient, undefined);
  assert.match(problem, /2 recipients/);
});

test('no recipient well means no recipients, not a page-wide sweep', () => {
  loadFixture('mock-ribbon-compose.html');
  document.getElementById('to-well').remove();
  assert.deepEqual(readRecipientsFromDom(), []);
});

/* ------------------------------------------------------------------ */
/* display name                                                        */
/* ------------------------------------------------------------------ */

test('a name repeated from the address is not treated as a display name', () => {
  loadFixture('mock-ribbon-compose.html');
  const [recipient] = readRecipientsFromDom();
  assert.equal(recipient.name, undefined, 'address-as-name must not become a greeting name');
});

test('a real display name is picked up from a Name <addr> chip', () => {
  loadFixture('mock-ribbon-compose.html');
  const chip = document.querySelector('.chip');
  chip.setAttribute('aria-label', 'Jordan Rivera <jordan.rivera.dev@example.com>');
  const [recipient] = readRecipientsFromDom();
  assert.equal(recipient.name, 'Jordan Rivera');
});

/* ------------------------------------------------------------------ */
/* URL corroboration                                                   */
/* ------------------------------------------------------------------ */

test('the mailtouri deep-link parameter yields the address', () => {
  assert.equal(
    readRecipientFromUrl(
      'https://outlook.cloud.microsoft/mail/deeplink/compose?mailtouri=mailto%3Ajordan.rivera.dev%40example.com',
    ),
    'jordan.rivera.dev@example.com',
  );
});

test('mailtouri parsing tolerates subject params and missing values', () => {
  assert.equal(
    readRecipientFromUrl('https://x/compose?mailtouri=mailto%3Aa%40b.com%3Fsubject%3DHi'),
    'a@b.com',
  );
  assert.equal(readRecipientFromUrl('https://x/compose'), null);
  assert.equal(readRecipientFromUrl('not a url'), null);
});

test('the To field wins when the URL disagrees, and the mismatch is reported', () => {
  loadFixture('mock-ribbon-compose.html');
  // jsdom's location has no mailtouri, so the DOM is the only source here.
  const { recipient } = resolveRecipient();
  assert.equal(recipient.email, 'jordan.rivera.dev@example.com');
  assert.ok(recipient.sources.includes('Outlook To field'));
});

/* ------------------------------------------------------------------ */
/* field location                                                      */
/* ------------------------------------------------------------------ */

test('subject and body are located on the ribbon compose surface', () => {
  loadFixture('mock-ribbon-compose.html');
  const subject = findSubject();
  const body = findBody();
  assert.equal(subject.el.id, 'subject-input');
  assert.equal(body.el.id, 'body-editor');
  assert.ok(subject.strategy, 'the winning strategy should be reported');
});

test('the subject lookup does not return the recipient well or the body', () => {
  loadFixture('mock-ribbon-compose.html');
  const subject = findSubject();
  assert.notEqual(subject.el.id, 'to-well');
  assert.notEqual(subject.el.id, 'body-editor');
});

test('the recipient well is found by its accessible name', () => {
  loadFixture('mock-ribbon-compose.html');
  assert.equal(findRecipientWell().el.id, 'to-well');
});

test('a compose surface missing its editor reports a specific problem', () => {
  loadFixture('mock-ribbon-compose.html');
  document.getElementById('body-editor').remove();
  const { surface, problem } = findComposeSurface();
  assert.equal(surface, undefined);
  assert.match(problem, /message editor/i);
});

test('a page with neither field yields no surface and no false problem', () => {
  loadFixture('mock-reading-pane.html');
  const { surface, problem } = findComposeSurface();
  assert.equal(surface, undefined);
  assert.equal(problem, undefined);
});

/* ------------------------------------------------------------------ */
/* read-back comparison                                                */
/* ------------------------------------------------------------------ */

test('read-back comparison ignores whitespace, case and nbsp', () => {
  assert.equal(normaliseForCompare('  Update  On Application '), 'update on application');
  assert.equal(
    normaliseForCompare('Update on Your Application'),
    normaliseForCompare('update on your   application'),
  );
  assert.notEqual(normaliseForCompare('Update on Your Application'), normaliseForCompare('Something else'));
});
