# Recruiter Assistant — Outlook Web

A local Chrome extension for reviewing LinkedIn job applications in Outlook Web. You
make every hiring decision and press Send yourself; the extension writes the email and
cleans up afterwards.

No backend, no APIs, no analytics. There is not a single network call in the built code.

---

## Contents

- [What it does](#what-it-does)
- [Install](#install)
- [The four actions](#the-four-actions)
- [Cleanup of the original email](#cleanup-of-the-original-email)
- [Settings](#settings)
- [Safety rules](#safety-rules)
- [When the cleanup prompt does not appear](#when-the-cleanup-prompt-does-not-appear)
- [If nothing appears](#if-nothing-appears)
- [Architecture](#architecture)
- [Development](#development)
- [Scope](#scope)
- [Known limitations](#known-limitations)

---

## What it does

Your workflow is unchanged. The extension joins at step 5.

1. Open a LinkedIn application email in Outlook.
2. Preview the resume and review the candidate — GitHub, LeetCode, whatever you normally do.
3. Click the candidate's address inside the PDF.
4. Outlook opens a compose window with them already in the **To** field.
5. **The Recruiter Assistant panel appears, bottom-right, with four actions.**
6. Click one. Subject and body are filled from your template.
7. Read it, then press Outlook's **Send**.
8. Once the send is confirmed, a prompt in your inbox tab offers to delete the original
   LinkedIn email.

It never sends mail. It never decides anything about a candidate.

## Install

```bash
npm install && npm run build
```

Then in Chrome: `chrome://extensions/` → **Developer mode** → **Load unpacked** →
select `dist/`. Reload any open Outlook tab afterwards.

Supported hosts:

```
https://outlook.cloud.microsoft/*     ← this tenant
https://outlook.office.com/*
https://outlook.office365.com/*
https://outlook.microsoft.com/*
```

`outlook.cloud.microsoft` is Microsoft's newer unified M365 domain. If the extension's
toolbar menu says *"Can't read or change site's data"*, the host in your address bar is
not in the manifest — add it to `src/manifest.json` (both `host_permissions` and
`content_scripts[0].matches`), rebuild, then reload the extension **and** the tab.

## The four actions

| Action | Subject | Routes to |
| --- | --- | --- |
| **Reject** | `Update on Your Application` | — |
| **Reject (in process)** | `Update in Interview Process` | — |
| **Shortlist → Ravilochan** | `Technical Screening – {{JOB_TITLE}}` | Ravilochan's Bookings link |
| **Shortlist → Abhi** | `Technical Screening – {{JOB_TITLE}}` | Abhi's Bookings link |

All four carry confirmed copy. The panel appears only when a compose surface is found
**and** the To field holds exactly one recipient — with two, there is no way to be
certain who the email is for, so it refuses and says so.

The two screening templates are the same round with a different interviewer. Both set
the expectation that the 20 minutes goes on technical substance rather than a resume
walkthrough, and both say plainly there is no coding exercise but that the candidate
should arrive able to talk about one or two projects in real depth. The self-run version
is first person; the delegated one names the interviewer, their role, and makes them the
point of contact for the stage. Each carries its own Bookings link — that is the entire
difference between the two actions, and clearing one blocks only that action.

### How the writing works

Field lookups are lists of semantic strategies (accessible name → placeholder →
structure), never CSS class names. The winning strategy shows in the panel's audit line,
so an Outlook change surfaces as "fell back to strategy 3" rather than silence.

- **Subject** is a React-controlled `<input>`. Assigning `.value` updates the DOM but not
  React's state, so Outlook reverts it and the mail goes out with an empty subject. The
  fill uses the native value setter plus a bubbling `input` event.
- **Body** is a rich-text `contenteditable`. `execCommand('insertHTML')` goes through the
  editor's own input handling; fallbacks are a synthetic paste, then direct DOM
  insertion. Bold and `<ul>` lists survive.

**Every write is read back and compared.** A write that cannot be verified is reported as
a failure — the panel never claims success it has not confirmed.

If the draft already has content you get *replace everything* / *insert at top* /
*cancel* rather than silent clobbering.

## Cleanup of the original email

The hard part is knowing *which* inbox tab holds the application email a given compose
window came from. This does not guess by matching subjects or names:

1. Clicking the address in the PDF opens a new tab, so Chrome reports its
   **`openerTabId`** — an authoritative link between compose tab and inbox tab.
2. At that moment two things are recorded: the address the deep link was opened for
   (from `mailtouri`, which Outlook drops from the URL once its SPA takes over — reading
   it later blocked every cleanup), and a snapshot of what the inbox tab has open:
   Outlook's own **message id** from the reading-pane URL, the sender, and the subject.
3. After a confirmed send, a prompt appears **in the inbox tab** — the tab that stays
   open — offering to delete that exact message.
4. Before deleting, everything is re-verified. If a different email is now open, or the
   sender and subject no longer read as a LinkedIn application, it refuses and says why.
5. The **resume PDF preview is closed first**. It is still open in that tab — it is what
   the address was clicked in — and it renders as a full overlay above the reading pane,
   where the Delete control lives. The checks then re-run against the settled state.
6. Deletion uses Outlook's own **Delete** control, found by accessible name: the reading-
   pane toolbar first, then **Delete inside the "More actions" menu**, which is where
   Outlook puts it on a narrow window or a collapsed ribbon. No coordinates, no keyboard
   shortcuts, no APIs.
7. Afterwards it confirms the message actually went, and the result says which route was
   used. If it cannot confirm, it says so instead of claiming success.

### Distinguishing sent from discarded

A compose window disappearing proves nothing — it happens on send, on discard, and on
closing a draft. So sends and discards are observed directly, in the capture phase, and a
disappearing compose window is read as a send **only** when one of those preceded it and
no discard did.

A send counts whether it came from the **Send button or `⌘/Ctrl+Enter`**. The keyboard
shortcut is how this workflow is actually driven, and watching only for clicks meant the
compose window vanished with no recorded send, the outcome read as abandoned, and the
cleanup step never ran at all. The confirmation prompt names which one it saw.

Anything else is reported as abandoned, and abandonment never leads to a deletion. The
extension only ever *observes* those clicks — it never synthesises one.

## Settings

Toolbar icon → **Settings**, or right-click the icon → Options.

- **Role & sender** — job title, company, your name and title.
- **Templates** — subject, body HTML, per-template Bookings link, interviewer and their
  role, with a live preview against a fictional candidate. `Insert blank line` adds a
  `<p>&nbsp;</p>` spacer at the cursor.
- **Greeting** — off by default, matching your templates. A name is never guessed from an
  email address.
- **Cleanup** — whether to offer deletion at all, and whether to ask first (both on).
- **Advanced** — debug logging, and a full reset.

Template variables:

```
{{JOB_TITLE}}  {{COMPANY}}  {{SENDER_NAME}}  {{SENDER_TITLE}}  {{CANDIDATE_EMAIL}}
{{CANDIDATE_NAME}}  {{INTERVIEWER}}  {{INTERVIEWER_TITLE}}  {{BOOKING_URL}}  {{GREETING}}
```

Rendering **fails closed**: an unresolved variable, an empty subject or body, or a
missing Bookings link blocks the action rather than sending a literal `{{…}}` to a
candidate. Substituted values are HTML-escaped, so an address containing markup cannot
break the body.

## Safety rules

Design constraints, enforced by tests:

- The extension **never sends** an email.
- Deletion is only considered after a send is **positively confirmed** — not after a Send
  click, which is not the same thing.
- Deletion requires the **exact message id** recorded when the compose tab opened, plus a
  LinkedIn sender **and** an application subject. Two independent signals; one is not
  enough, so a colleague's email mentioning LinkedIn can never qualify.
- Deletion goes through Outlook's own UI.
- Each workflow is keyed by real Chrome tab id, so two candidates open at once cannot be
  crossed.
- Every failure fails closed: nothing sent, nothing deleted, and an explicit message
  saying which step could not be completed.

## When the cleanup prompt does not appear

Every reason cleanup can be skipped used to live only in the service-worker console,
which made "no prompt appeared" indistinguishable from all of its causes. Now:

- The **compose panel** shows the outcome — `Sent — cleanup prompt is waiting in your
  inbox tab`, or `Sent, but the original email was left alone` with the reason.
- The **toolbar popup** shows `Last cleanup` and `· because`, plus — when you are on an
  inbox tab with a message open — exactly what a snapshot of that tab reads: the sender,
  the subject, which reading-pane strategy matched, and whether it counts as a LinkedIn
  application.

That last block is the fastest way to see why a particular email was not offered.

### Stored settings win over shipped defaults

Anything saved in Settings takes precedence over the shipped templates, which is right
for wording you wrote — but it also meant an improvement to a template never reached a
config saved earlier. That is why blank lines were missing from real emails while the
Options preview looked correct.

A stored template is now replaced by the shipped one in exactly two cases: the wording is
identical and only the spacing differs, or the stored copy was still an unreviewed
placeholder. A genuine edit is never overwritten, and a Bookings link or interviewer you
set is always kept.

## If nothing appears

Click the extension's toolbar icon. The popup reports which cause it is:

| Popup says | Meaning |
| --- | --- |
| `Content script: NOT running` + unsupported host | the tab's host is not in the manifest |
| `Content script: NOT running` on a supported host | site access is "On click", or the extension/tab needs reloading |
| `Recruiter panel: no compose detected` | no compose surface, or not exactly one recipient |

The content script also writes **one unconditional line** to the page console on every
load:

```
[RecruiterAssistant] v0.1.0 active — outlook.cloud.microsoft/mail/… (top frame)
```

If that line is absent, the extension is not injected and nothing else matters.

## Architecture

```
src/
├── manifest.json            MV3 — storage + tabs, four Outlook hosts
├── background/
│   └── service-worker.ts    openerTabId → source tab, snapshots, workflow state
├── content/
│   ├── index.ts             entry; routes a tab to its compose or inbox role
│   ├── compose.ts           multi-strategy subject / body / recipient lookups
│   ├── fill.ts              React-safe + editor-safe writes with read-back checks
│   ├── send-detector.ts     sent vs discarded vs abandoned
│   ├── source-email.ts      snapshot, verify, delete, cleanup prompt
│   └── recruiter-ui.ts      the four-action panel
├── options/                 settings page
├── popup/                   toolbar status popup
└── shared/
    ├── config.ts            settings shape + the four templates
    ├── template.ts          substitution, escaping, fail-closed validation
    ├── messages.ts          typed cross-tab messages
    ├── settings.ts          storage wrapper + migrations
    ├── redact.ts            email extraction / masking
    └── log.ts               namespaced logging, off by default
```

Workflow state lives in `chrome.storage.session` because an MV3 service worker is killed
between events.

## Development

```bash
npm run build      # one-shot build into dist/
npm run dev        # rebuild on change (still needs a reload in chrome://extensions)
npm run typecheck  # tsc --noEmit
npm test           # node:test + jsdom
```

### Browser harness

jsdom has no layout engine and no `execCommand`, so it cannot exercise the panel, the
fill path, or the deletion flow. Two harness pages load the real built `dist/content.js`
against mock DOMs with a stubbed `chrome.*`:

```bash
npm run build && python3 -m http.server 8731
```

- `tests/harness/compose-harness.html` — the ribbon compose surface, including a
  **simulated React-controlled subject input that reverts naive writes**, and a Send
  button that records if anything clicks it.
- `tests/harness/source-harness.html` — the inbox tab with a LinkedIn application open,
  and a Delete button that records clicks.

These caught bugs the unit tests structurally could not: a visibility test that only
failed in a real browser, a panel that unmounted at exactly the moment it should have
reported the send verdict, and a recipient scan that read the `From:` address as a second
recipient.

### Adversarial audit

The send-detection and deletion logic was additionally reviewed by independent agents
whose findings were each put through a refutation pass. Two critical defects survived,
both able to delete the wrong email, and both are fixed with regression tests:

1. **A stale opener record could authorise a deletion.** Chrome records an opener for
   *every* tab opened from another tab — ctrl-click, "open in new tab", a link in a
   message body — so a tab could carry an unrelated inbox snapshot indefinitely. The same
   hole let a hand-edited recipient delete the original candidate's application. The
   opener link is now only accepted when the compose deep link's own `mailtouri` names the
   exact address being emailed, and opener records are age-pruned.
2. **A closing compose tab was treated as proof of a send.** If Outlook rejected the send
   and the user closed the tab, the application email was deleted for a message that never
   left. That path now applies the same 90-second freshness bound as the in-page detector,
   honours a discard veto reported straight to the service worker, and always asks for
   confirmation regardless of the setting.

Also fixed from the same audit: a click on a *disabled* Send counting as a send; body
read-back verifying only the first line, which is identical across both screening
templates; the sender being read from the message body, so any email mentioning LinkedIn
could qualify; a `[role="main"]` fallback that could read the subject from the message
list; and several flags (`filledWith`, `pendingConfirm`, the × dismissal) that were never
cleared and so stuck for the life of the tab.

## Scope

Deliberately **not** built, per the brief: AI resume screening, automatic accept/reject
decisions, resume PDF downloading, OCR or resume parsing, LinkedIn / GitHub / LeetCode
API integration, Power Automate, Microsoft Graph, any backend, candidate database, ATS
integration, or automated interview evaluation.

Documented as possible later, and not started: resume summaries, candidate scoring,
interview notes, candidate tracking, application statistics, automated follow-ups. The
template and action model would extend to these; nothing in the code anticipates them.

## Known limitations

- Field lookups assume English accessible names (`Send`, `Delete`, `To`, `Subject`). A
  localised tenant needs the regexes in `compose.ts`, `send-detector.ts` and
  `source-email.ts` extended.
- The panel does not appear in frames smaller than 360×280.
- If Outlook rejects a send *after* the click while still closing the compose window,
  that reads as a confirmed send. The confirmation prompt is the backstop — which is why
  **Ask me first** defaults to on. A compose tab that merely *closes* after a Send click
  is treated as weaker evidence: it always asks, whatever the setting says, and the prompt
  states that the send itself was not observed.
- The **shortlist and cleanup paths have not been run against real Outlook yet** — only
  against the browser harnesses. The two rejection templates have been used for real.
- The "More actions" fallback for Delete, and the subject-as-`contenteditable` fallback,
  are tested against mocks rather than an Outlook that actually needs them.
- Both Bookings links are defaults in `src/shared/config.ts`, so they would land in
  version control if this repo is ever pushed somewhere shared. Ravilochan's is a private
  meeting type carrying a `bookingcode` access parameter.
