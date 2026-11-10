// Widget contents for the generated illustrations, built by the real code of
// the integration: the app itself (src/app.js) wired to the fake Gladys of the
// SDK and the fake Microsoft of the tests (a real local HTTP server), on a
// made-up Tuesday, 10 November 2026 at 10:00 in Paris: three users enabled
// "My calendars", Alice and Chloé signed in (made-up mailboxes), Bob has not
// typed his code yet. Never real data.
// Written as JSON for render.js: node gen.mjs <repo> <out.json>

import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

const [repo, out] = process.argv.slice(2);
const load = (file) => import(pathToFileURL(join(repo, file)).href);
const { createApp } = await load('src/app.js');
const { LANGUAGES } = await load('src/i18n.js');
const { createFakeGladys } = await import(
  pathToFileURL(join(repo, 'node_modules/@gladysassistant/integration-sdk/lib/testing.js')).href
);
const { startFakeMicrosoft, graphEvent, CLIENT_ID } = await load('test/helpers/fake-microsoft.js');

const manifest = JSON.parse(readFileSync(join(repo, 'gladys-assistant-integration.json'), 'utf8'));
const WIDGET_KEYS = ['accounts'];
const NOW = Date.UTC(2026, 10, 10, 9, 0, 0);

// A small seeded generator (mulberry32): the same picture at every run.
function random(seed) {
  let state = seed;
  return () => {
    state = (state + 0x6d2b79f5) | 0;
    let t = Math.imul(state ^ (state >>> 15), 1 | state);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/**
 * @description One language: the accounts widget.
 * @param {string} language - Language.
 * @returns {Promise<object>} The widget contents.
 */
async function contents(language) {
  const next = random(20261110);
  const event = (id, day, hour, subject) =>
    graphEvent(id, {
      start: `2026-11-${String(day).padStart(2, '0')}T${String(hour).padStart(2, '0')}:00:00`,
      end: `2026-11-${String(day).padStart(2, '0')}T${String(hour + 1).padStart(2, '0')}:00:00`,
      subject,
    });
  const mailbox = (name, calendars) => ({
    me: { displayName: name, mail: `${name.toLowerCase()}@example.com` },
    calendars: calendars.map(([id, title]) => ({ id, name: title, hexColor: '#0f6cbd' })),
    events: Object.fromEntries(
      calendars.map(([id]) => [
        id,
        Array.from({ length: 20 + Math.floor(next() * 40) }, (_, i) =>
          event(`${id}-${i}`, 1 + (i % 28), 8 + (i % 10), `Event ${i}`),
        ),
      ]),
    ),
  });
  const ms = await startFakeMicrosoft({
    mailboxes: {
      alice: mailbox('Alice', [
        ['a1', 'Calendar'],
        ['a2', 'Family'],
        ['a3', 'Birthdays'],
      ]),
      chloe: mailbox('Chloe', [
        ['c1', 'Calendar'],
        ['c2', 'Work'],
      ]),
    },
  });
  const dir = mkdtempSync(join(tmpdir(), 'outlook-illustration-'));
  const users = [
    { selector: 'alice', first_name: 'Alice', language: 'en' },
    { selector: 'bob', first_name: 'Bob', language: 'en' },
    { selector: 'chloe', first_name: 'Chloé', language: 'en' },
  ];
  const gladys = createFakeGladys({
    config: { client_id: CLIENT_ID },
    calendarAccounts: users.map((user) => ({ user, config: {} })),
  });
  const polls = [];
  const app = createApp(gladys, {
    loginUrl: ms.loginUrl,
    graphUrl: ms.graphUrl,
    sleep: async () => {},
    setTimer: (fn) => polls.push(fn),
    clearTimer: () => {},
    accountsFile: join(dir, 'accounts.json'),
    now: () => NOW,
    timeZone: 'Europe/Paris',
    autoSync: false,
  });
  await gladys.connect();
  await app.whenReady();
  await app.users.idle();
  // Alice and Chloé type their codes; Bob has not yet. The first code of the
  // fake is the one the app check asked for and threw away.
  const [alice, , chloe] = ms.pending().slice(-3);
  ms.approve('alice', alice);
  ms.approve('chloe', chloe);
  for (const poll of polls.splice(0)) {
    poll();
  }
  await app.users.idle();
  const result = {};
  for (const key of WIDGET_KEYS) {
    result[key] = (await gladys.fake.widgetGet(key, { language })).data.content;
  }
  app.stop();
  await gladys.fake.shutdown();
  await ms.close();
  rmSync(dir, { recursive: true, force: true });
  return result;
}

const all = {};
for (const language of LANGUAGES) {
  all[language] = await contents(language);
}

const card = (key) => {
  const widget = manifest.widgets.find((item) => item.key === key);
  return { icon: widget.icon, label: widget.label };
};
writeFileSync(
  out,
  JSON.stringify({
    contents: all,
    images: {},
    cards: Object.fromEntries(WIDGET_KEYS.map((key) => [key, card(key)])),
  }),
);
