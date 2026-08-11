# Outlook Web — what was actually observed

Facts established against the live tenant during development, kept because the
implementation depends on them and because several are non-obvious.

This file no longer describes a diagnostic panel: that tooling was removed once the
workflow was built. Selectors now live in `src/content/compose.ts`,
`src/content/send-detector.ts` and `src/content/source-email.ts`, each as an ordered
list of strategies with the winning one reported at runtime.

## Environment

| Field | Value |
| --- | --- |
| Host | `outlook.cloud.microsoft` — Microsoft's unified M365 domain, **not** `outlook.office.com` |
| UI generation | Ribbon-based "new Outlook" web client (File / Message / Insert / Format text / Draw / Options) |
| Compose route | `/mail/deeplink/compose?mailtouri=mailto%3A<address>` |
| Reading-pane route | `/mail/inbox/id/<messageId>` |
| Compose opens in | a **new tab in the same window**; the inbox tab stays open behind it |
| Send gesture used | **⌘+Enter**, not the Send button |

## Load-bearing observations

**The compose URL carries the candidate's address**, as `mailtouri`. This is how a
compose tab is tied to the candidate it was opened for. It is captured at
`tabs.onCreated` from `pendingUrl`, because Outlook's SPA rewrites the URL after load
and drops the query string — reading it later blocked every cleanup.

**The reading-pane URL carries Outlook's own message id.** That id is the only
unambiguous handle on "the email this compose window came from", and a deletion is only
ever offered for the exact id recorded when the compose tab opened.

**A recipient chip renders the address twice** — `addr <addr>` — because a `mailto:`
deep link has no display name. So for this workflow there is normally *no* candidate
name to greet, which is why the templates open without a greeting and why a name is
never derived from an address.

**`From: <your own address>` sits inside the same compose dialog as the recipient
chips.** Reading recipients from any ancestor of the To well therefore picks up your own
address as a second recipient, and the panel then refuses every action. Recipient
scanning is scoped to the To well itself.

**Consecutive `<p>` tags render with no visible gap** in the compose editor. Paragraph
spacing has to be explicit `<p>&nbsp;</p>` spacers; relying on margins produced emails
with every paragraph on its own line and no blank lines.

**The resume PDF preview is a full overlay above the reading pane**, and it is still
open in the inbox tab after the compose window opens — it is what the address was
clicked in. The message Delete control sits behind it, so it is closed before deleting.

**Popped-out compose runs on `about:blank`.** Seen in the Templates editor: the full
ribbon compose UI renders in a separate window whose URL is `about:blank`, inheriting
the opener's origin. `"match_about_blank": true` covers it.

## Not yet observed

- Whether Delete for the open message is ever *only* available via "More actions" in
  this tenant. A fallback exists and is tested against a mock, not the real UI.
- Whether Outlook ever renders the subject as a `contenteditable` rather than an
  `<input>` here. A fallback exists for it.
- What the DOM looks like when a send is *rejected* (offline, policy block). This is the
  one case that can read as a confirmed send, and it is why the confirmation prompt
  defaults to on.
