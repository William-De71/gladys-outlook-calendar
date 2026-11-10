import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createGraphClient, GRAPH_ERRORS, kindOfGraphStatus, retryAfterMs } from '../src/graph.js';
import { graphEvent, startFakeMicrosoft } from './helpers/fake-microsoft.js';

test('status kinds and Retry-After', () => {
  assert.equal(kindOfGraphStatus(401), GRAPH_ERRORS.AUTH);
  assert.equal(kindOfGraphStatus(403), GRAPH_ERRORS.FORBIDDEN);
  assert.equal(kindOfGraphStatus(404), GRAPH_ERRORS.NOT_FOUND);
  assert.equal(kindOfGraphStatus(429), GRAPH_ERRORS.THROTTLED);
  assert.equal(retryAfterMs('3'), 3000);
  assert.equal(retryAfterMs(null), 1000);
  assert.equal(
    retryAfterMs(new Date(10_000).toUTCString(), () => 5000),
    5000,
  );
});

test('pages, Retry-After, one token refresh on 401', async (t) => {
  const events = Array.from({ length: 600 }, (_, i) =>
    graphEvent(`e${i}`, { start: '2026-11-12T08:00:00', end: '2026-11-12T09:00:00' }),
  );
  const ms = await startFakeMicrosoft({
    mailboxes: {
      m: {
        me: { displayName: 'M', userPrincipalName: 'm@x' },
        calendars: [{ id: 'c', name: 'C' }],
        events: { c: events },
      },
    },
  });
  t.after(() => ms.close());
  // A valid token for the mailbox, through the device flow of the fake.
  const code = await fetch(`${ms.loginUrl}/common/oauth2/v2.0/devicecode`, {
    method: 'POST',
    body: new URLSearchParams({ client_id: '11111111-2222-3333-4444-555555555555', scope: 'x' }),
  }).then((r) => r.json());
  ms.approve('m', code.device_code);
  const tokens = await fetch(`${ms.loginUrl}/common/oauth2/v2.0/token`, {
    method: 'POST',
    body: new URLSearchParams({
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      client_id: '11111111-2222-3333-4444-555555555555',
      device_code: code.device_code,
    }),
  }).then((r) => r.json());
  const asked = [];
  let current = 'stale';
  const waits = [];
  const graph = createGraphClient({
    baseUrl: ms.graphUrl,
    // Like the real cache: stale until refreshed once, then the fresh one.
    getToken: async (force) => {
      asked.push(force);
      if (force) {
        current = tokens.access_token;
      }
      return current;
    },
    sleep: async (ms_) => waits.push(ms_),
  });
  ms.fail('/calendarView', 429, { error: { message: 'slow' } }, { 'Retry-After': '2' });
  const list = await graph.calendarView('c', '2026-11-01T00:00:00Z', '2026-12-01T00:00:00Z');
  assert.equal(list.length, 600, 'three pages of 250');
  assert.deepEqual(asked.slice(0, 2), [false, true]);
  assert.equal(asked.filter(Boolean).length, 1, 'refreshed once');
  assert.deepEqual(waits, [2000]);
  assert.deepEqual(await graph.me(), { name: 'M', email: 'm@x' });
  await assert.rejects(
    graph.calendarView('missing', '2026-11-01T00:00:00Z', '2026-12-01T00:00:00Z'),
    {
      kind: GRAPH_ERRORS.NOT_FOUND,
    },
  );
});
