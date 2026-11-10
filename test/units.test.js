import { test } from 'node:test';
import assert from 'node:assert/strict';
import { setImmediate as tick } from 'node:timers/promises';
import { validateWidgetContent } from '@gladysassistant/integration-sdk';
import { normalizeAccount, normalizeConfig } from '../src/config.js';
import { createEvents, TRIGGERS } from '../src/events.js';
import { splitWindow, USER_STATES } from '../src/users.js';
import { buildAccountsWidget } from '../src/widgets.js';

test('configuration: valid ids only, selects as numbers', () => {
  const config = normalizeConfig({
    client_id: ' 11111111-2222-3333-4444-555555555555 ',
    tenant: 'contoso.onmicrosoft.com',
    sync_minutes: '5',
    horizon_months: '99',
  });
  assert.equal(config.client_id, '11111111-2222-3333-4444-555555555555');
  assert.equal(config.tenant, 'contoso.onmicrosoft.com');
  assert.equal(config.sync_minutes, 5);
  assert.equal(config.horizon_months, 12);
  const bad = normalizeConfig({ client_id: 'my app', tenant: 'https://evil' });
  assert.equal(bad.client_id, '');
  assert.equal(bad.typed_client_id, true);
  assert.equal(bad.tenant, 'common');
  assert.deepEqual(normalizeAccount({ private_events: 'x' }), {
    private_events: 'busy',
    declined_events: 'skip',
  });
});

test('more than 500 events: disjoint windows, a straddling event in each it overlaps', () => {
  const events = Array.from({ length: 1200 }, (_, i) => {
    const start = new Date(Date.UTC(2026, 10, 1) + i * 3600_000);
    return {
      external_id: `e${i}`,
      start: start.toISOString(),
      end: new Date(start.getTime() + 7200_000).toISOString(),
    };
  });
  const chunks = splitWindow(events, '2026-10-01T00:00:00.000Z', '2027-11-01T00:00:00.000Z');
  assert.equal(chunks.length, 3);
  assert.equal(chunks[0].from, '2026-10-01T00:00:00.000Z');
  assert.equal(chunks.at(-1).to, '2027-11-01T00:00:00.000Z');
  for (let i = 1; i < chunks.length; i += 1) {
    assert.equal(chunks[i].from, chunks[i - 1].to);
  }
  for (const chunk of chunks) {
    assert.ok(chunk.events.length <= 501, `${chunk.events.length}`);
  }
  // The event that starts an hour before a boundary and ends after it is in both.
  const boundary = chunks[1].from;
  const straddling = events.find((e) => e.start < boundary && e.end > boundary);
  assert.ok(chunks[0].events.includes(straddling) && chunks[1].events.includes(straddling));
  assert.equal(splitWindow(events.slice(0, 3), 'a', 'b').length, 1);
});

test('triggers: once per kind and user, recovery with the minutes', async () => {
  const sent = [];
  let clock = 0;
  const events = createEvents({
    publish: async (key, data) => sent.push([key, data]),
    now: () => clock,
  });
  const alice = { selector: 'alice', first_name: 'Alice' };
  const bob = { selector: 'bob', first_name: 'Bob' };
  events.failure(alice, { kind: 'server', message: 'x' });
  events.failure(alice, { kind: 'server', message: 'x' });
  events.failure(bob, { kind: 'throttled', message: 'y' });
  clock = 3 * 60_000;
  events.success(alice);
  events.success(alice);
  events.linkRequired(alice, 'revoked');
  await tick();
  assert.deepEqual(
    sent.map(([key, data]) => [
      key,
      data.first_name,
      data.kind ?? data.down_minutes ?? data.reason,
    ]),
    [
      [TRIGGERS.SYNC_ERROR, 'Alice', 'server'],
      [TRIGGERS.SYNC_ERROR, 'Bob', 'throttled'],
      [TRIGGERS.SYNC_RECOVERED, 'Alice', 3],
      [TRIGGERS.LINK_REQUIRED, 'Alice', 'revoked'],
    ],
  );
});

test('widget: valid for every state, never more than 8 components', () => {
  const users = [
    {
      selector: 'a',
      first_name: 'Alice',
      state: USER_STATES.LINKED,
      entry: { last_sync: 0, calendars: 2, events: 40 },
      failing: null,
    },
    { selector: 'b', first_name: 'Bob', state: USER_STATES.LINKING, entry: {}, failing: null },
    {
      selector: 'c',
      first_name: 'Chloé',
      state: USER_STATES.ERROR,
      entry: { calendars: 1, events: 3 },
      failing: { kind: 'throttled' },
    },
    { selector: 'd', first_name: '', state: USER_STATES.LINK_STOPPED, entry: {}, failing: null },
    { selector: 'e', first_name: 'Eve', state: USER_STATES.NO_APP, entry: {}, failing: null },
  ];
  for (const language of ['en', 'fr', 'de', 'es']) {
    for (const [config, list] of [
      [normalizeConfig({ client_id: '11111111-2222-3333-4444-555555555555' }), users],
      [normalizeConfig(), []],
      [normalizeConfig({ client_id: '11111111-2222-3333-4444-555555555555' }), []],
    ]) {
      const content = buildAccountsWidget({
        users: list,
        config,
        language,
        now: 0,
        timeZone: 'UTC',
      });
      assert.deepEqual(validateWidgetContent(content), [], language);
      assert.ok(content.components.length <= 8);
    }
  }
  const content = buildAccountsWidget({
    users,
    config: normalizeConfig({ client_id: '11111111-2222-3333-4444-555555555555' }),
    language: 'en',
    now: 0,
    timeZone: 'UTC',
  });
  assert.equal(content.components.find((c) => c.label === 'Linked users').value, 2);
  assert.equal(content.components.find((c) => c.label === 'Events').value, '43');
});
