// -----------------------------------------------------------------------------
// Entry point of the Outlook Calendar external integration.
//
// A calendar provider (manifest `type: "calendar"`): every Gladys user who
// enables "My calendars" signs in with their own Microsoft account (a code
// shown in a temporary calendar), and their Outlook calendars are kept in
// sync in Gladys — calendar view, `calendar.event-is-coming` scenes, the
// assistant. Read-only, through Microsoft Graph, with the administrator's own
// Microsoft Entra app registration. On top of that:
//   - scene triggers: a user must sign in / signed in / sync fails / recovers;
//   - scene action and button: sync now;
//   - dashboard widget: who is linked.
//
// This file only wires the SDK client: the handlers are in src/app.js.
//
// Environment variables provided by the Gladys supervisor:
//   GLADYS_HOST_API_URL / GLADYS_INTEGRATION_TOKEN / GLADYS_INTEGRATION_SELECTOR
//   TZ (the timezone of Gladys: times shown in the widget and the sign-in calendar)
// The SDK reads the first three: `new GladysIntegration()` is enough.
// -----------------------------------------------------------------------------

import { readFileSync } from 'node:fs';
import { GladysIntegration, logger } from '@gladysassistant/integration-sdk';
import { createApp } from './src/app.js';

const { version } = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8'));

const gladys = new GladysIntegration();
createApp(gladys);

// --- Startup -----------------------------------------------------------------------
logger.info(`Starting the Outlook Calendar integration ${version}...`);
gladys.connect().catch((err) => {
  logger.error('Initial connection failed', err);
  process.exit(1);
});
