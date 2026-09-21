# Contributing

This is a small, personally-maintained Chrome extension. PRs are welcome, but
everything lands through review — direct pushes to `main` are blocked by branch
protection.

## Before opening a PR

```bash
npm run typecheck   # tsc --noEmit
npm test            # node:test + jsdom
npm run build       # dist/ must build cleanly
```

All three must pass. If behavior changed, update the relevant section of
[README.md](README.md) in the same PR — stale docs (wrong action count, wrong
keyboard shortcuts, an outdated description) are worse than no docs, and this repo
has shipped that mistake before.

## Never commit

- A **real** Microsoft Bookings URL, mailbox address, or booking code. Bookings
  links are per-installation and belong in Options (`chrome.storage.local`), never
  in `src/shared/config.ts` — see the comment above `BOOKING_URL_UNSET`. Shipped
  defaults leave `bookingUrl` unset; the app fails closed until it's configured.
- A **real** candidate's name, email, or resume filename in a test fixture or
  harness. Use synthetic data (`example.com`, invented names) — real applicants
  never consented to appearing in this repository.
- Any API key, token, or credential. This extension has no backend and makes no
  network calls; it should never need one.

## Review

Every PR requires an approving review before it can merge. Because this repo has
one maintainer, GitHub won't let that person approve their own PR — as the
repository admin, use **"Merge without waiting for requirements to be met"** once
you've reviewed your own diff, or have someone else approve it first.

## Commit style

Short, imperative, scoped messages (`fix(shortcuts): ...`, `feat(actions): ...`,
`docs: ...`), matching the existing log. Prefer several small, reviewable commits
over one large one.
