// -----------------------------------------------------------------------------
// Outlook (Microsoft Graph) → Gladys calendars and events.
//
// Ids: Graph ids are long (~150 characters) while a Gladys external id is at
// most 255 with its `ext:<selector>:<user>:` prefix, so the integration keeps
// a hash of them: stable, short, and an occurrence of a recurring event (its
// own Graph id) stays its own event.
//
// Times: requested in UTC, so a timed event is `<dateTime>Z`. A full-day
// event is read by calendar date: Graph writes it at midnight; should a
// timezone conversion move it to 22:00 or 23:00 the evening before, the date
// is rounded to the next day.
// -----------------------------------------------------------------------------

import { createHash } from 'node:crypto';
import { t } from './i18n.js';

/** Colors of the `color` enum of Graph calendars (Outlook's palette). */
export const CALENDAR_COLORS = {
  lightBlue: '#a6d1f5',
  lightGreen: '#87d28e',
  lightOrange: '#fcab73',
  lightGray: '#c0c0c0',
  lightYellow: '#f4d07a',
  lightTeal: '#71c7bd',
  lightPink: '#ee9fbf',
  lightBrown: '#ca9d6f',
  lightRed: '#f08c84',
};
export const DEFAULT_COLOR = '#0078d4';

/**
 * @description A short stable hash of a provider id.
 * @param {string} id - A Graph id.
 * @returns {string} 24 hexadecimal characters.
 */
export function shortId(id) {
  return createHash('sha256').update(String(id)).digest('hex').slice(0, 24);
}

/**
 * @description Truncate a text, with an ellipsis.
 * @param {unknown} value - The text.
 * @param {number} max - Characters at most.
 * @returns {string} The text.
 */
export function clip(value, max) {
  const text = String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim();
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

/**
 * @description The color of a Graph calendar.
 * @param {object} calendar - `{ hexColor, color }`.
 * @returns {string} `#rrggbb`.
 */
export function calendarColor(calendar) {
  if (typeof calendar?.hexColor === 'string' && /^#[0-9a-f]{6}$/i.test(calendar.hexColor)) {
    return calendar.hexColor.toLowerCase();
  }
  return CALENDAR_COLORS[calendar?.color] ?? DEFAULT_COLOR;
}

/**
 * @description An instant of Graph, in UTC.
 * @param {{ dateTime: string, timeZone?: string }} value - `{ dateTime: '2026-11-10T08:00:00.0000000', timeZone: 'UTC' }`.
 * @returns {string|null} `2026-11-10T08:00:00.000Z`.
 */
export function utcInstant(value) {
  const text = String(value?.dateTime ?? '');
  const match = text.match(/^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})(\.\d+)?/);
  if (!match) {
    return null;
  }
  const date = new Date(`${match[1]}Z`);
  return Number.isNaN(date.getTime()) ? null : date.toISOString();
}

/**
 * @description The calendar date of a full-day bound.
 * @param {{ dateTime: string }} value - Graph bound, midnight (or shifted).
 * @returns {string|null} `2026-11-15`.
 */
export function calendarDate(value) {
  const text = String(value?.dateTime ?? '');
  const match = text.match(/^(\d{4}-\d{2}-\d{2})T(\d{2})/);
  if (!match) {
    return null;
  }
  const [, date, hour] = match;
  if (Number(hour) < 12) {
    return date;
  }
  // Midnight moved back a few hours by a conversion: the next day.
  const next = new Date(`${date}T00:00:00Z`);
  next.setUTCDate(next.getUTCDate() + 1);
  return next.toISOString().slice(0, 10);
}

/**
 * @description The Gladys calendar of a Graph calendar.
 * @param {object} gladys - SDK client (`externalId`).
 * @param {string} user - The user selector.
 * @param {object} calendar - Graph calendar.
 * @param {{ email: string }} account - Who signed in.
 * @param {string} language - The user language.
 * @returns {object} CalendarInput.
 */
export function toCalendar(gladys, user, calendar, account, language) {
  const owner =
    calendar?.owner?.address && calendar.owner.address !== account.email
      ? calendar.owner.address
      : '';
  return {
    external_id: calendarExternalId(gladys, user, calendar.id),
    name: clip(calendar.name || t('untitledCalendar', language), 100) || 'Outlook',
    description: clip(
      owner
        ? t('calendarSharedBy', language, { owner })
        : t('calendarOf', language, { account: account.email || account.name || 'Outlook' }),
      500,
    ),
    color: calendarColor(calendar),
  };
}

/**
 * @description The external id of a Graph calendar.
 * @param {object} gladys - SDK client.
 * @param {string} user - The user selector.
 * @param {string} calendarId - The Graph calendar id.
 * @returns {string} The external id.
 */
export function calendarExternalId(gladys, user, calendarId) {
  return gladys.externalId(`${user}:cal-${shortId(calendarId)}`);
}

/**
 * @description The Gladys event of a Graph event, or null when it is left out.
 * @param {object} gladys - SDK client.
 * @param {string} user - The user selector.
 * @param {object} event - Graph event.
 * @param {{ private_events: string, declined_events: string }} account - The user's choices.
 * @param {string} language - The user language.
 * @returns {object|null} CalendarEventInput.
 */
export function toEvent(gladys, user, event, account, language) {
  if (!event?.id || event.isCancelled) {
    return null;
  }
  if (event.responseStatus?.response === 'declined' && account.declined_events === 'skip') {
    return null;
  }
  const isPrivate = ['private', 'confidential'].includes(event.sensitivity);
  if (isPrivate && account.private_events === 'skip') {
    return null;
  }
  const hidden = isPrivate && account.private_events === 'busy';
  const external_id = gladys.externalId(`${user}:ev-${shortId(event.id)}`);
  const name = hidden
    ? t('privateEvent', language)
    : clip(event.subject, 200) || t('untitledEvent', language);
  const result = { external_id, name };
  if (event.isAllDay) {
    const start = calendarDate(event.start);
    if (!start) {
      return null;
    }
    const end = calendarDate(event.end);
    result.start = start;
    result.end = end && end > start ? end : null;
    result.full_day = true;
  } else {
    const start = utcInstant(event.start);
    if (!start) {
      return null;
    }
    const end = utcInstant(event.end);
    result.start = start;
    result.end = end && end >= start ? end : null;
    result.full_day = false;
  }
  if (!hidden) {
    const location = clip(event.location?.displayName, 500);
    if (location) {
      result.location = location;
    }
    const description = clip(event.bodyPreview, 1000);
    if (description) {
      result.description = description;
    }
    const link = event.onlineMeeting?.joinUrl || event.webLink;
    if (typeof link === 'string' && /^https:\/\//.test(link) && link.length <= 500) {
      result.url = link;
    }
  }
  return result;
}
