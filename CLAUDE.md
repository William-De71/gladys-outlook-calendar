# CLAUDE.md

Project rules for coding assistants (Claude Code reads this file) and for human
contributors.

## What this is

A Gladys Assistant **external integration** of type **calendar** that keeps the
**Outlook calendars** of each Gladys user in sync: a Node.js service (Node 22+,
ESM, no build step) in its own sandboxed Docker container, talking to Gladys
through
[`@gladysassistant/integration-sdk`](https://github.com/GladysAssistant/integration-sdk-js)
and to Microsoft (identity platform + Graph v1.0), **read-only**, with the
Microsoft account each user links and one app registration of the
administrator (public client, device code flow). Runtime dependency: the SDK
only (plain `fetch`). `type: "calendar"`, Gladys 5.3.0+,
`transports: ["cloud"]`, `categories: ["services"]`. Not affiliated with
Microsoft.

The contract (SDK README, "Calendar providers"): the integration syncs, the
core stores; `getCalendarAccounts()` gives the users who enabled "My
calendars", `publishCalendars` / `publishCalendarEvents` (with a window that
prunes) / `deleteCalendar`, user-scoped external ids, `onCalendarAccountUpdated`.
**Never tested with a real Microsoft account**: request and response shapes
follow learn.microsoft.com (device authorization grant, `calendarView`,
`Prefer: outlook.timezone`), checked on 2026-11-10.

## Commands

```bash
npm install
npm test                                  # node --test; a fake Microsoft on 127.0.0.1
node --test test/app.test.js              # one file
UPDATE_BASELINE=1 npm test                # rewrite test/fixtures/manifest.baseline.json (on purpose only)
npm run lint                              # ESLint
npm run format:check                      # Prettier, as the CI runs it
npx github:GladysAssistant/integration-store . --skip-image-check   # store checks
scripts/build-cover.sh                    # cover.jpg from docs/assets/cover.svg
docs/illustrations/render.sh              # docs/assets/widgets.<lang>.png
node .github/scripts/release.mjs next     # the version the next release would get
```

## Architecture

```
index.js               SDK wiring only (new GladysIntegration + createApp)
src/app.js             every handler: calendar account updates, the two buttons, the
                       scene action, the widget and its button, configuration, loop
src/users.js           per user: linking (device code shown in a temporary calendar,
                       polls, renewals) and sync (calendars, events by window, spaced
                       writes, change detection, 403 skip, revoked → relink)
src/microsoft-auth.js  device code, token polling, refresh (rotation), error kinds
src/graph.js           GET /me, /me/calendars, calendarView; pages, Retry-After, 401 once
src/convert.js         Graph → Gladys: hashed ids, colors, UTC instants, full-day dates,
                       private / declined / cancelled events, links
src/store.js           /data/accounts.json (mode 600): refresh token and state per user
src/events.js          triggers: link_required, linked, sync_error (once per kind and
                       user), sync_recovered
src/widgets.js         the accounts widget (first names and states only)
src/config.js          app id, tenant, sync period, horizon; per-user choices
src/i18n.js            runtime texts, en / fr / de / es (sign-in calendar in the user's)
src/time.js            local times (TZ of Gladys, Intl only)
test/helpers/fake-microsoft.js  the fake identity platform and Graph
.github/scripts/release.mjs     version, changelog, bump (tested in test/)
```

## Rules

- **Every key users store is forever**: the keys of `config_schema` and
  `account_schema` and the values of their selects, `actions`, the widget and
  its button, the triggers (fields, options, variables), the scene action
  (outputs). `test/structure.test.js` compares them to
  `test/fixtures/manifest.baseline.json`: never update it to make a test pass
  without a documented migration. External ids too: `<user>:cal-<hash>`,
  `<user>:ev-<hash>`, `<user>:link` — changing the hash duplicates every event.
- **Read-only**: scopes `offline_access User.Read Calendars.Read`, Graph GET
  only. Gladys never writes back to the provider (contract).
- **The sign-in code goes to the user only**: in their own temporary calendar
  (private until shared) and its event, never in the status (admins only), a
  log line or the widget. The widget shows first names and states, never an
  e-mail, a calendar name or an event.
- **Tokens**: refresh tokens in `/data/accounts.json` (mode 600), access
  tokens in memory; never logged, never in an error message. Microsoft rotates
  the refresh token: always keep the new one.
- **Calendar writes are spaced** (`WRITE_SPACING_MS`, ~27 a minute, under the
  30 of the core); an unchanged calendar is not pushed again; more than 500
  events → disjoint sub-windows, a straddling event in each it overlaps.
- **Times**: requested in UTC (`Prefer: outlook.timezone="UTC"`); full-day
  events by calendar date, a midnight shifted back by a conversion rounded to
  the next day.
- **Scene events** once per transition and user. A revoked sign-in is not a
  `sync_error`: it is `link_required` with reason `revoked`.
- **The container rootfs is read-only.** Write only under `/data`.
- **Four languages.** Every user-facing text (manifest, `src/i18n.js`,
  `docs/*.md`, README) exists in en, fr, de, es; `test/i18n.test.js` checks it,
  and the sign-in calendar name stays under 100 characters.
- **Widget images.** After any change to a widget content, run
  `docs/illustrations/render.sh` again and look at the four PNG files. Tile
  labels stay short (three tiles in a row).
- **Logo**: the Outlook icon (Wikimedia Commons, public domain, Microsoft
  trademark) names the service only; keep the "not affiliated" notices.
- **Versions belong to the Release workflow**: never edit `version` in
  `package.json`, the lockfile or the manifest, nor the `docker_image` tag.
  Below 1.0.0 until the maintainer runs Release with `release-as: 1.0.0`.
- **Changelog.** Ne jamais éditer CHANGELOG.md à la main, il est généré à la
  release. Écris le titre de PR comme la ligne que l'utilisateur lira.
- **Commits and PR titles** follow Conventional Commits, in English.
- **Forum drafts** (`.github/forum-*.md`) are never committed.
- **Style.** Prettier formats, ESLint catches mistakes. Comments explain why,
  in English. Tests never touch the network.

## Pitfalls

- `account_schema` refuses `oauth2` / `account_link` (milestone 1 of the
  core): that is why the code travels through a calendar. When the core gets a
  per-user Connect, switch to it with a migration of the stored tokens.
- The "Check the app registration" button and every startup request a device
  code that is thrown away: harmless, but the fake counts it (the illustration
  approves the last three codes).
- Error codes: AADSTS700016 unknown app, AADSTS7000218 public client flows off,
  AADSTS50020 / AADSTS9002331 account type vs tenant, `invalid_grant`
  (AADSTS70008…) revoked or expired refresh token.
- `node --test` also runs `test/helpers/`: nothing starts on import.
- A sync started by a button runs in the background: tests wait for
  `app.users.idle()` and `app.store.flush()` before deleting their folder.
- The validator of the Gladys core creates `gladys-development.duckdb` in the
  current folder: run it from elsewhere.
