# Changelog

All notable changes to this integration are documented in this file. The format
follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions
follow [Semantic Versioning](https://semver.org/). This file is generated at
each release from the Conventional Commits: do not edit it by hand.

## [0.1.0] - unreleased

### Added

- **calendar:** the Outlook calendars of each Gladys user in the Gladys calendar (Microsoft Graph, read-only): calendars shared with them included, recurrences expanded, all-day events, place, notes and online meeting link; events deleted in Outlook disappear
- **sign-in:** each user links their own Microsoft account with a code (device code flow) shown in a temporary private calendar of theirs, renewed for an hour; refresh tokens kept per user and rotated; a revoked sign-in brings the code back
- **privacy:** private events shown as "Busy", with their details or left out; declined invitations left out by default; read-only scopes
- **scenes:** triggers when a user must sign in, signed in, when a sync fails and works again; a "sync now" scene action
- **widgets:** linked users, calendars and events, the state of each user and a sync button
- **configuration:** "Check the app registration" and "Sync now" buttons; sync every 5 to 60 minutes, 3 to 12 months ahead
- **i18n:** English, French, German and Spanish for the manifest, the runtime texts, the documentation and the README

[0.1.0]: https://github.com/William-De71/gladys-outlook-calendar/releases/tag/v0.1.0
