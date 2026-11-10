// The whole integration against the fake Gladys of the SDK
// (`@gladysassistant/integration-sdk/testing`) and a fake Microsoft
// (test/helpers/fake-microsoft.js), a real HTTP server on 127.0.0.1. Timers
// are manual: `clock.run()` fires the sign-in polls. No test touches the
// network.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { setImmediate as tick } from 'node:timers/promises';
import { validateWidgetContent } from '@gladysassistant/integration-sdk';
import { createFakeGladys } from '@gladysassistant/integration-sdk/testing';
import { ACTIONS, createApp, SCENE_ACTIONS } from '../src/app.js';
import { TRIGGERS } from '../src/events.js';
import { WIDGET_ACTIONS, WIDGETS } from '../src/widgets.js';
import { CLIENT_ID, graphEvent, startFakeMicrosoft } from './helpers/fake-microsoft.js';

const NOW = Date.parse('2026-11-10T09:00:00Z');
const ALICE = { selector: 'alice', first_name: 'Alice', language: 'fr' };

const mailboxes = () => ({
  'alice-ms': {
    me: {
      displayName: 'Alice Martin',
      mail: 'alice@outlook.fr',
      userPrincipalName: 'alice@outlook.fr',
    },
    calendars: [
      {
        id: 'AAMk-calendar-main',
        name: 'Calendrier',
        hexColor: '#FF8C00',
        color: 'auto',
        owner: { address: 'alice@outlook.fr' },
      },
      {
        id: 'AAMk-calendar-family',
        name: 'Famille',
        hexColor: '',
        color: 'lightGreen',
        owner: { address: 'bob@outlook.fr' },
      },
    ],
    events: {
      'AAMk-calendar-main': [
        graphEvent('e1', {
          start: '2026-11-12T08:00:00',
          end: '2026-11-12T09:00:00',
          subject: 'Dentiste',
          location: { displayName: '12 rue de la Paix' },
        }),
        graphEvent('e2', {
          start: '2026-11-15T00:00:00',
          end: '2026-11-17T00:00:00',
          subject: 'Week-end',
          isAllDay: true,
        }),
        graphEvent('e3', {
          start: '2026-11-13T10:00:00',
          end: '2026-11-13T11:00:00',
          subject: 'Médecin',
          sensitivity: 'private',
          location: { displayName: 'Hôpital' },
        }),
        graphEvent('e4', {
          start: '2026-11-14T10:00:00',
          end: '2026-11-14T11:00:00',
          subject: 'Refusé',
          responseStatus: { response: 'declined' },
        }),
        graphEvent('e5', {
          start: '2026-11-14T12:00:00',
          end: '2026-11-14T13:00:00',
          subject: 'Annulé',
          isCancelled: true,
        }),
        graphEvent('e6', {
          start: '2028-01-01T10:00:00',
          end: '2028-01-01T11:00:00',
          subject: 'Trop loin',
        }),
      ],
      'AAMk-calendar-family': [
        graphEvent('f1', {
          start: '2026-11-11T17:00:00',
          end: '2026-11-11T18:00:00',
          subject: 'Piscine',
        }),
      ],
    },
  },
});

/** Manual timers: the sign-in polls run when the test says so. */
function manualTimers() {
  const pending = new Map();
  let id = 0;
  return {
    setTimer: (fn) => {
      id += 1;
      pending.set(id, fn);
      return id;
    },
    clearTimer: (handle) => pending.delete(handle),
    async run() {
      const due = [...pending.entries()];
      pending.clear();
      for (const [, fn] of due) {
        fn();
      }
      for (let i = 0; i < 20; i += 1) {
        await tick();
      }
    },
    size: () => pending.size,
  };
}

const start = async (
  t,
  {
    config = { client_id: CLIENT_ID },
    accounts = [{ user: ALICE, config: {} }],
    clock = { now: NOW },
  } = {},
) => {
  const ms = await startFakeMicrosoft({ mailboxes: mailboxes() });
  const dir = await mkdtemp(join(tmpdir(), 'outlook-app-'));
  const gladys = createFakeGladys({ config, calendarAccounts: accounts });
  const timers = manualTimers();
  const app = createApp(gladys, {
    loginUrl: ms.loginUrl,
    graphUrl: ms.graphUrl,
    sleep: async () => {},
    setTimer: timers.setTimer,
    clearTimer: timers.clearTimer,
    accountsFile: join(dir, 'accounts.json'),
    now: () => clock.now,
    timeZone: 'Europe/Paris',
    autoSync: false,
  });
  await gladys.connect();
  await app.whenReady();
  await app.users.idle();
  t.after(async () => {
    // A sync started by a button may still run: let it end before cleaning.
    await app.users.idle();
    await app.store.flush();
    app.stop();
    await gladys.disconnect?.();
    await ms.close();
    await rm(dir, { recursive: true, force: true });
  });
  return { ms, gladys, app, timers, dir, clock };
};

const settle = async (app) => {
  for (let i = 0; i < 20; i += 1) {
    await tick();
  }
  await app.users.idle();
};
const linkCalendar = (gladys) =>
  gladys.fake.calendars.find((c) => c.external_id.endsWith(':alice:link'));
const sceneEvents = (gladys, key) => gladys.fake.sceneEvents.filter((e) => e.key === key);
const lastStatus = (gladys) => gladys.fake.connectionStatuses.at(-1);

/**
 * @description Sign Alice in: approve the code, run the poll.
 * @returns {Promise<void>} Linked and synchronized.
 */
async function signIn({ ms, timers, app }) {
  ms.approve('alice-ms');
  await timers.run();
  await settle(app);
}

test('without the app id: the status asks for it and the user sees a waiting calendar', async (t) => {
  const { gladys, ms } = await start(t, { config: {} });
  assert.equal(lastStatus(gladys).connected, false);
  assert.match(JSON.stringify(lastStatus(gladys)), /application \(client\) ID/);
  assert.match(linkCalendar(gladys).name, /en attente de l’administrateur/);
  assert.equal(ms.requests.length, 0);
});

test('a user who enables gets a sign-in code in a temporary private calendar', async (t) => {
  const { gladys, timers } = await start(t);
  assert.equal(lastStatus(gladys).connected, true);
  const link = linkCalendar(gladys);
  assert.match(link.name, /^Microsoft : tapez CODE\d+ sur microsoft\.com\/devicelogin$/);
  assert.match(link.description, /valable jusqu’à 10:15/);
  assert.equal(link.shared, false, 'private until the user shares it');
  // The code also as an event, in the calendar view.
  const event = gladys.fake.calendarEvents.find((e) => e.calendar_external_id === link.external_id);
  assert.match(event.name, /CODE\d+/);
  assert.equal(timers.size(), 1, 'one poll scheduled');
  assert.equal(sceneEvents(gladys, TRIGGERS.LINK_REQUIRED)[0].data.reason, 'new');
});

test('once signed in: the temporary calendar goes, the Outlook calendars and events come', async (t) => {
  const ctx = await start(t);
  const { gladys, ms, dir } = ctx;
  await signIn(ctx);
  assert.equal(linkCalendar(gladys), undefined);
  assert.equal(sceneEvents(gladys, TRIGGERS.LINKED)[0].data.account, 'alice@outlook.fr');

  const calendars = gladys.fake.calendars.filter((c) => c.user === 'alice');
  assert.deepEqual(
    calendars.map((c) => [c.name, c.color, c.description]),
    [
      ['Calendrier', '#ff8c00', 'Calendrier Outlook de alice@outlook.fr'],
      ['Famille', '#87d28e', 'Calendrier Outlook partagé par bob@outlook.fr'],
    ],
  );
  for (const calendar of calendars) {
    assert.ok(calendar.external_id.length <= 255);
    assert.match(calendar.external_id, /:alice:cal-[0-9a-f]{24}$/);
  }

  const events = gladys.fake.calendarEvents;
  const byName = Object.fromEntries(events.map((e) => [e.name, e]));
  assert.deepEqual(Object.keys(byName).sort(), ['Dentiste', 'Occupé', 'Piscine', 'Week-end']);
  assert.equal(byName.Dentiste.location, '12 rue de la Paix');
  assert.equal(new Date(byName.Dentiste.start).toISOString(), '2026-11-12T08:00:00.000Z');
  assert.equal(byName['Week-end'].full_day, true);
  // A private event: "Busy", without its place.
  assert.equal(byName['Occupé'].location ?? null, null);
  // Read-only scopes, times in UTC, the token in a header.
  const view = ms.requests.find((r) => r.path.endsWith('/calendarView'));
  assert.equal(view.prefer, 'outlook.timezone="UTC"');
  assert.match(view.auth, /^Bearer at-alice-ms-/);
  const code = ms.requests.find((r) => r.path.endsWith('/devicecode'));
  assert.equal(code.form.scope, 'offline_access User.Read Calendars.Read');
  // The refresh token is kept, readable by the owner only.
  const file = join(dir, 'accounts.json');
  assert.match(await readFile(file, 'utf8'), /"refresh_token":"rt-alice-ms-/);
  assert.equal((await stat(file)).mode & 0o777, 0o600);
});

test('an unchanged calendar is not pushed again; a change is', async (t) => {
  const ctx = await start(t);
  const { gladys, ms, app } = ctx;
  await signIn(ctx);
  const before = gladys.fake.publishedCalendarEvents.length;
  await app.users.syncAll();
  assert.equal(gladys.fake.publishedCalendarEvents.length, before);
  ms.mailboxes['alice-ms'].events['AAMk-calendar-main'].splice(0, 1);
  await app.users.syncAll();
  assert.equal(gladys.fake.publishedCalendarEvents.length, before + 1);
  assert.equal(
    gladys.fake.calendarEvents.some((e) => e.name === 'Dentiste'),
    false,
    'deleted in Outlook, pruned by the window',
  );
});

test('a calendar turned off by the user is skipped; one removed in Outlook is deleted', async (t) => {
  const ctx = await start(t);
  const { gladys, ms, app } = ctx;
  await signIn(ctx);
  const family = gladys.fake.calendars.find((c) => c.name === 'Famille');
  family.sync = false;
  ms.mailboxes['alice-ms'].events['AAMk-calendar-family'].push(
    graphEvent('f2', { start: '2026-11-18T17:00:00', end: '2026-11-18T18:00:00', subject: 'Judo' }),
  );
  await app.users.syncAll();
  assert.equal(
    gladys.fake.calendarEvents.some((e) => e.name === 'Judo'),
    false,
  );
  ms.mailboxes['alice-ms'].calendars.pop();
  await app.users.syncAll();
  assert.ok(gladys.fake.deletedCalendars.includes(family.external_id));
});

test('a revoked sign-in brings the code back, with a trigger', async (t) => {
  const ctx = await start(t);
  const { gladys, ms, app, clock } = ctx;
  await signIn(ctx);
  ms.revoke('alice-ms');
  clock.now += 2 * 3600 * 1000; // the access token expired
  await app.users.syncAll();
  await settle(app);
  const required = sceneEvents(gladys, TRIGGERS.LINK_REQUIRED);
  assert.equal(required.at(-1).data.reason, 'revoked');
  assert.match(linkCalendar(gladys).name, /CODE\d+/);
  // Signing in again resumes the sync.
  await signIn(ctx);
  assert.equal(linkCalendar(gladys), undefined);
});

test('a Graph failure: one sync_error per kind, then sync_recovered', async (t) => {
  const ctx = await start(t);
  const { gladys, ms, app } = ctx;
  await signIn(ctx);
  ms.fail('/me/calendars', 500, { error: { message: 'boom' } });
  await app.users.syncAll();
  ms.fail('/me/calendars', 500, { error: { message: 'boom' } });
  await app.users.syncAll();
  await settle(app);
  assert.equal(sceneEvents(gladys, TRIGGERS.SYNC_ERROR).length, 1);
  assert.equal(sceneEvents(gladys, TRIGGERS.SYNC_ERROR)[0].data.kind, 'server');
  await app.users.syncAll();
  await settle(app);
  assert.equal(sceneEvents(gladys, TRIGGERS.SYNC_RECOVERED).length, 1);
});

test('an expired code is renewed four times, then waits for the user to save again', async (t) => {
  const ctx = await start(t);
  const { gladys, timers, clock, app } = ctx;
  for (let renewal = 1; renewal <= 5; renewal += 1) {
    clock.now += 16 * 60 * 1000;
    await timers.run();
    await settle(app);
  }
  assert.match(linkCalendar(gladys).name, /code expiré/);
  assert.equal(timers.size(), 0);
  // Saving "My calendars" again starts a new code.
  await gladys.fake.calendarAccountUpdated('alice');
  await settle(app);
  assert.match(linkCalendar(gladys).name, /CODE\d+/);
});

test('a declined sign-in stops until the user saves again', async (t) => {
  const ctx = await start(t);
  const { gladys, ms, timers, app } = ctx;
  ms.decline();
  await timers.run();
  await settle(app);
  assert.match(linkCalendar(gladys).name, /connexion refusée/);
});

test('a user who disables is forgotten', async (t) => {
  const ctx = await start(t);
  const { gladys, app, dir } = ctx;
  await signIn(ctx);
  gladys.fake.calendarAccounts.splice(0, 1);
  gladys.fake.calendars = gladys.fake.calendars.filter((c) => c.user !== 'alice');
  await gladys.fake.calendarAccountUpdated('alice');
  await settle(app);
  assert.doesNotMatch(await readFile(join(dir, 'accounts.json'), 'utf8'), /alice/);
  assert.deepEqual(app.users.snapshot(), []);
});

test('check button: an unknown app is explained with the Microsoft code', async (t) => {
  const { gladys } = await start(t, {
    config: { client_id: '99999999-2222-3333-4444-555555555555' },
  });
  const result = await gladys.fake.action(ACTIONS.CHECK_APP);
  assert.match(JSON.stringify(result), /AADSTS700016/);
  assert.equal(lastStatus(gladys).connected, false);
});

test('sync button, scene action and widget', async (t) => {
  const ctx = await start(t);
  const { gladys } = ctx;
  await signIn(ctx);
  assert.match(
    JSON.stringify(await gladys.fake.action(ACTIONS.SYNC_NOW)),
    /1 user\(s\): 2 calendar\(s\) and 4 event\(s\)/,
  );
  const scene = await gladys.fake.sceneAction(SCENE_ACTIONS.SYNC_NOW, {});
  assert.deepEqual(scene.data.outputs, { users: 1, events: 4 });
  for (const language of ['en', 'fr', 'de', 'es']) {
    const ack = await gladys.fake.widgetGet(WIDGETS.ACCOUNTS, { language });
    assert.deepEqual(validateWidgetContent(ack.data.content), [], language);
  }
  const ack = await gladys.fake.widgetGet(WIDGETS.ACCOUNTS, { language: 'fr' });
  const status = ack.data.content.components.find((c) => c.type === 'status');
  assert.deepEqual(
    status.items.map((i) => [i.label, i.value]),
    [['Alice', 'Synchro à 10:00']],
  );
  assert.doesNotMatch(
    JSON.stringify(ack.data.content),
    /outlook\.fr|Dentiste/,
    'no e-mail, no event',
  );
  const toast = await gladys.fake.widgetAction(WIDGETS.ACCOUNTS, WIDGET_ACTIONS.SYNC);
  assert.equal(toast.success, true);
});

test('the sign-in survives a restart', async (t) => {
  const ms = await startFakeMicrosoft({ mailboxes: mailboxes() });
  const dir = await mkdtemp(join(tmpdir(), 'outlook-restart-'));
  t.after(async () => {
    await ms.close();
    await rm(dir, { recursive: true, force: true });
  });
  const make = () => {
    const gladys = createFakeGladys({
      config: { client_id: CLIENT_ID },
      calendarAccounts: [{ user: ALICE, config: {} }],
    });
    const timers = manualTimers();
    const app = createApp(gladys, {
      loginUrl: ms.loginUrl,
      graphUrl: ms.graphUrl,
      sleep: async () => {},
      setTimer: timers.setTimer,
      clearTimer: timers.clearTimer,
      accountsFile: join(dir, 'accounts.json'),
      now: () => NOW,
      autoSync: false,
    });
    return { gladys, timers, app };
  };
  const first = make();
  await first.gladys.connect();
  await first.app.whenReady();
  await first.app.users.idle();
  await signIn({ ms, ...first });
  await first.app.store.flush();
  first.app.stop();

  const second = make();
  await second.gladys.connect();
  await second.app.whenReady();
  await second.app.users.idle();
  assert.equal(
    second.gladys.fake.calendars.some((c) => c.external_id.endsWith(':alice:link')),
    false,
  );
  assert.equal(second.gladys.fake.calendarEvents.length, 4);
  second.app.stop();
});
