import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  calendarColor,
  calendarDate,
  shortId,
  toCalendar,
  toEvent,
  utcInstant,
} from '../src/convert.js';
import { normalizeAccount } from '../src/config.js';
import { graphEvent } from './helpers/fake-microsoft.js';

const gladys = { externalId: (id) => `ext:outlook:${id}` };
const account = normalizeAccount();

test('instants in UTC, full-day dates by calendar date even when shifted', () => {
  assert.equal(
    utcInstant({ dateTime: '2026-11-12T08:00:00.0000000', timeZone: 'UTC' }),
    '2026-11-12T08:00:00.000Z',
  );
  assert.equal(utcInstant({ dateTime: 'garbage' }), null);
  assert.equal(calendarDate({ dateTime: '2026-11-15T00:00:00.0000000' }), '2026-11-15');
  assert.equal(
    calendarDate({ dateTime: '2026-11-14T23:00:00.0000000' }),
    '2026-11-15',
    'midnight Paris in UTC',
  );
  assert.equal(calendarDate({ dateTime: '2026-12-31T22:00:00' }), '2027-01-01');
  assert.equal(calendarDate({ dateTime: '2026-11-15T05:00:00' }), '2026-11-15');
});

test('ids are short, stable and user-scoped', () => {
  const long = 'AAMkAGI2TGuLAAA='.repeat(12);
  assert.match(shortId(long), /^[0-9a-f]{24}$/);
  assert.equal(shortId(long), shortId(long));
  const event = toEvent(
    gladys,
    'alice',
    graphEvent(long, { start: '2026-11-12T08:00:00', end: '2026-11-12T09:00:00' }),
    account,
    'en',
  );
  assert.match(event.external_id, /^ext:outlook:alice:ev-[0-9a-f]{24}$/);
});

test('calendar colors: hex first, else the Outlook palette, else blue', () => {
  assert.equal(calendarColor({ hexColor: '#FF8C00' }), '#ff8c00');
  assert.equal(calendarColor({ hexColor: '', color: 'lightTeal' }), '#71c7bd');
  assert.equal(calendarColor({ color: 'auto' }), '#0078d4');
  const shared = toCalendar(
    gladys,
    'alice',
    { id: 'x', name: '', owner: { address: 'bob@x.fr' } },
    { email: 'alice@x.fr' },
    'en',
  );
  assert.equal(shared.name, 'Calendar');
  assert.equal(shared.description, 'Outlook calendar shared by bob@x.fr');
});

test('private, declined and cancelled events follow the choices of the user', () => {
  const secret = graphEvent('p', {
    start: '2026-11-13T10:00:00',
    end: '2026-11-13T11:00:00',
    subject: 'Médecin',
    sensitivity: 'private',
    location: { displayName: 'Hôpital' },
    bodyPreview: 'Bilan',
  });
  assert.deepEqual(Object.keys(toEvent(gladys, 'a', secret, account, 'fr')).sort(), [
    'end',
    'external_id',
    'full_day',
    'name',
    'start',
  ]);
  assert.equal(toEvent(gladys, 'a', secret, account, 'fr').name, 'Occupé');
  assert.equal(
    toEvent(gladys, 'a', secret, normalizeAccount({ private_events: 'show' }), 'fr').location,
    'Hôpital',
  );
  assert.equal(
    toEvent(gladys, 'a', secret, normalizeAccount({ private_events: 'skip' }), 'fr'),
    null,
  );
  const declined = graphEvent('d', {
    start: '2026-11-13T10:00:00',
    end: '2026-11-13T11:00:00',
    responseStatus: { response: 'declined' },
  });
  assert.equal(toEvent(gladys, 'a', declined, account, 'en'), null);
  assert.ok(toEvent(gladys, 'a', declined, normalizeAccount({ declined_events: 'show' }), 'en'));
  assert.equal(
    toEvent(
      gladys,
      'a',
      { ...declined, isCancelled: true },
      normalizeAccount({ declined_events: 'show' }),
      'en',
    ),
    null,
  );
});

test('links: the meeting link first, https only, 500 characters at most', () => {
  const meeting = graphEvent('m', {
    start: '2026-11-13T10:00:00',
    end: '2026-11-13T11:00:00',
    isOnlineMeeting: true,
    onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/abc' },
  });
  assert.equal(
    toEvent(gladys, 'a', meeting, account, 'en').url,
    'https://teams.microsoft.com/l/meetup-join/abc',
  );
  const long = graphEvent('l', {
    start: '2026-11-13T10:00:00',
    end: '2026-11-13T11:00:00',
    webLink: `https://x/${'a'.repeat(600)}`,
  });
  assert.equal(toEvent(gladys, 'a', long, account, 'en').url, undefined);
  const untitled = graphEvent('u', {
    start: '2026-11-13T10:00:00',
    end: '2026-11-13T11:00:00',
    subject: '  ',
  });
  assert.equal(toEvent(gladys, 'a', untitled, account, 'de').name, '(Ohne Titel)');
  const allDay = graphEvent('w', {
    start: '2026-11-15T00:00:00',
    end: '2026-11-16T00:00:00',
    isAllDay: true,
  });
  assert.deepEqual(
    [
      toEvent(gladys, 'a', allDay, account, 'en').start,
      toEvent(gladys, 'a', allDay, account, 'en').end,
    ],
    ['2026-11-15', '2026-11-16'],
  );
});
