/** Test-only barrel: bundled by tests/build-tests.mjs so Node can import the TS sources. */
export { extractEmails } from '../src/shared/redact';
export { DEFAULT_SETTINGS, ACTION_ORDER } from '../src/shared/config';
export {
  renderTemplate,
  validateBookingUrl,
  escapeHtml,
  htmlToText,
  KNOWN_VARIABLES,
} from '../src/shared/template';
export {
  findSubject,
  findBody,
  findRecipientWell,
  readRecipientsFromDom,
  readRecipientFromUrl,
  resolveRecipient,
  findComposeSurface,
} from '../src/content/compose';
export { normaliseForCompare } from '../src/content/fill';
export { decideOutcome } from '../src/content/send-detector';
export {
  captureSnapshot,
  judgeApplication,
  readMessageIdFromUrl,
  verifySourceStillMatches,
} from '../src/content/source-email';
export { parseMailtoRecipient } from '../src/shared/mailto';
export { mergeStoredSettings } from '../src/shared/settings';
