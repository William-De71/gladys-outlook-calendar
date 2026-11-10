// -----------------------------------------------------------------------------
// Every Gladys user who enabled "My calendars": their link with Microsoft and
// the sync of their Outlook calendars.
//
// LINKING. Gladys has no per-user "Connect" button for calendar integrations
// yet, and the connection status is shown to administrators only. So the
// sign-in code reaches the user through the one place that is theirs: a
// temporary private calendar, published in their "My calendars" block (and in
// their calendar view, as an event), named after the code:
//
//     "Microsoft: type ABCD-EFGH at microsoft.com/devicelogin"
//
// The user signs in on any device; the integration polls Microsoft, keeps the
// refresh token, deletes the temporary calendar and publishes the real ones.
// An expired code is renewed (4 times, about an hour); after that, saving
// "My calendars" again starts a new one. A revoked sign-in brings the
// temporary calendar back.
//
// SYNC. Per user, serially: the calendars (`/me/calendars`), those removed on
// the Outlook side deleted, then for each calendar the user did not turn off,
// the events of the window (30 days back, N months ahead), recurrences
// expanded by Microsoft. A window replaces its content, so an event deleted in
// Outlook disappears. A calendar whose events did not change is not pushed
// again. Calendar writes are spaced (the core allows 30 a minute).
// -----------------------------------------------------------------------------

import { createLogger } from '@gladysassistant/integration-sdk';
import { HISTORY_DAYS, normalizeAccount } from './config.js';
import { AUTH_ERRORS, AuthError } from './microsoft-auth.js';
import { createGraphClient, GRAPH_ERRORS, GraphError } from './graph.js';
import { clip, toCalendar, toEvent } from './convert.js';
import { formatTime } from './time.js';
import { t } from './i18n.js';

const logger = createLogger({ name: 'users' });

/** Spacing of calendar writes: at most ~27 a minute, under the 30 of the core. */
export const WRITE_SPACING_MS = 2200;
/** Events per publishCalendarEvents request, and per calendar. */
export const MAX_EVENTS_PER_REQUEST = 500;
export const MAX_EVENTS_PER_CALENDAR = 10_000;
/** Times a sign-in code is renewed on its own after expiring. */
export const MAX_CODE_RENEWALS = 4;
/** The access token is renewed this long before it expires. */
const TOKEN_MARGIN_MS = 2 * 60 * 1000;

/** The states of a user, as the widget shows them. */
export const USER_STATES = {
  LINKING: 'linking',
  LINK_STOPPED: 'link_stopped',
  NO_APP: 'no_app',
  LINKED: 'linked',
  ERROR: 'error',
};

/**
 * @description The instant of an event bound, in ms (full-day dates at UTC midnight).
 * @param {string|null} value - ISO instant or date.
 * @returns {number} Milliseconds.
 */
function instant(value) {
  return Date.parse(
    /^\d{4}-\d{2}-\d{2}$/.test(String(value)) ? `${value}T00:00:00Z` : String(value),
  );
}

/**
 * @description Split a window so that no request carries more than `max`
 * events: disjoint sub-windows, an event that straddles a boundary in every
 * sub-window it overlaps (the pruning rule of the core).
 * @param {object[]} events - Gladys events.
 * @param {string} from - ISO start of the window.
 * @param {string} to - ISO end of the window.
 * @param {number} [max] - Events per request.
 * @returns {{ from: string, to: string, events: object[] }[]} The requests.
 */
export function splitWindow(events, from, to, max = MAX_EVENTS_PER_REQUEST) {
  if (events.length <= max) {
    return [{ from, to, events }];
  }
  const starts = events.map((event) => instant(event.start)).sort((a, b) => a - b);
  const bounds = [Date.parse(from)];
  for (let index = max; index < starts.length; index += max) {
    // A boundary at an instant where an event starts; never the same twice.
    if (starts[index] > bounds.at(-1)) {
      bounds.push(starts[index]);
    }
  }
  bounds.push(Date.parse(to));
  const chunks = [];
  for (let index = 0; index < bounds.length - 1; index += 1) {
    const chunkFrom = bounds[index];
    const chunkTo = bounds[index + 1];
    chunks.push({
      from: new Date(chunkFrom).toISOString(),
      to: new Date(chunkTo).toISOString(),
      events: events.filter((event) => {
        const start = instant(event.start);
        const end = event.end ? instant(event.end) : start;
        return start < chunkTo && (end > chunkFrom || start >= chunkFrom);
      }),
    });
  }
  return chunks;
}

/**
 * @description Build the user manager.
 * @param {object} options - Options.
 * @param {object} options.gladys - SDK client.
 * @param {object} options.auth - Microsoft sign-in client.
 * @param {object} options.store - The per-user store.
 * @param {object} options.events - The trigger state machine.
 * @param {() => object} options.getConfig - The normalized configuration.
 * @param {string} [options.graphUrl] - Graph root (tests).
 * @param {typeof fetch} [options.fetchImpl] - fetch.
 * @param {(ms: number) => Promise<void>} [options.sleep] - Waits.
 * @param {(fn: Function, ms: number) => unknown} [options.setTimer] - Timers (tests).
 * @param {(handle: unknown) => void} [options.clearTimer] - Timers (tests).
 * @param {() => number} [options.now] - Clock.
 * @param {string} [options.timeZone] - The timezone of Gladys.
 * @param {() => void} [options.onChange] - Something the widget shows changed.
 * @returns {object} The manager.
 */
export function createUsers({
  gladys,
  auth,
  store,
  events,
  getConfig,
  graphUrl,
  fetchImpl = globalThis.fetch,
  sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
  setTimer = (fn, ms) => {
    const handle = setTimeout(fn, ms);
    handle.unref?.();
    return handle;
  },
  clearTimer = (handle) => clearTimeout(handle),
  now = Date.now,
  timeZone,
  onChange = () => {},
}) {
  // selector -> { accessToken, expiresAt }
  const tokens = new Map();
  // selector -> { timer, renewals, code, state }
  const links = new Map();
  // external id -> key of the last content pushed
  const pushed = new Map();
  // selector -> the account of the last call (user + values)
  const accounts = new Map();
  // Sign-in polls running (tests wait for them).
  const inflight = new Set();
  let lastWrite = 0;
  let queue = Promise.resolve();

  // --- Calendar writes, spaced -------------------------------------------------
  async function write(call) {
    const wait = lastWrite + WRITE_SPACING_MS - now();
    if (wait > 0) {
      await sleep(wait);
    }
    lastWrite = now();
    return call();
  }

  const linkCalendarId = (selector) => gladys.externalId(`${selector}:link`);
  const calendarPrefix = (selector) => gladys.externalId(`${selector}:cal-`);
  const languageOf = (user) => user?.language || 'en';

  // --- The temporary calendar that carries the sign-in code --------------------
  async function showLink(user, kind, values = {}) {
    const language = languageOf(user);
    const id = linkCalendarId(user.selector);
    const name = clip(t(`link_${kind}`, language, values), 100);
    const description = clip(t(`link_${kind}_description`, language, values), 500);
    const key = JSON.stringify([name, description, values.expiresAt ?? null]);
    if (pushed.get(id) === key) {
      return;
    }
    await write(() =>
      gladys.publishCalendars(user.selector, [
        { external_id: id, name, description, color: '#d13438' },
      ]),
    );
    if (kind === 'code') {
      // The code also in the calendar view, as an event until it expires.
      const from = new Date(now() - 60_000).toISOString();
      const to = new Date(values.expiresAt).toISOString();
      await write(() =>
        gladys.publishCalendarEvents(
          id,
          [
            {
              external_id: gladys.externalId(`${user.selector}:link-code`),
              name: clip(name, 200),
              start: from,
              end: to,
              description,
            },
          ],
          { from, to },
        ),
      ).catch((err) =>
        logger.warn(`Sign-in event not published for ${user.selector}: ${err?.message ?? err}`),
      );
    }
    pushed.set(id, key);
  }

  async function hideLink(selector) {
    const id = linkCalendarId(selector);
    const exists = (await gladys.getCalendars(selector)).some(
      (calendar) => calendar.external_id === id,
    );
    if (exists) {
      await write(() => gladys.deleteCalendar(id)).catch((err) => {
        if (err?.status !== 404) {
          throw err;
        }
      });
    }
    pushed.delete(id);
  }

  function stopLink(selector) {
    const link = links.get(selector);
    if (link?.timer) {
      clearTimer(link.timer);
    }
    links.delete(selector);
  }

  /**
   * @description Start (or restart) the sign-in of a user.
   * @param {object} user - LinkedUser.
   * @param {{ renewal?: boolean }} [options] - A renewal after expiry keeps the count.
   */
  async function startLink(user, { renewal = false } = {}) {
    const previous = links.get(user.selector);
    stopLink(user.selector);
    const config = getConfig();
    if (!config.client_id) {
      links.set(user.selector, { state: USER_STATES.NO_APP, renewals: 0 });
      await showLink(user, 'no_app');
      onChange();
      return;
    }
    let code;
    try {
      code = await auth.requestDeviceCode(config);
    } catch (err) {
      logger.warn(`Sign-in code refused for ${user.selector}: ${err?.message ?? err}`);
      const kind = err?.kind === AUTH_ERRORS.APP ? 'app' : 'unavailable';
      links.set(user.selector, { state: USER_STATES.LINK_STOPPED, renewals: 0, error: kind });
      await showLink(user, kind);
      onChange();
      return;
    }
    const link = {
      state: USER_STATES.LINKING,
      renewals: renewal ? (previous?.renewals ?? 0) + 1 : 0,
      code,
      intervalMs: code.intervalMs,
      timer: null,
    };
    links.set(user.selector, link);
    await showLink(user, 'code', {
      code: code.userCode,
      url: code.verificationUri.replace(/^https:\/\//, '').replace(/\/$/, ''),
      time: formatTime(code.expiresAt, { language: languageOf(user), timeZone, now: now() }),
      expiresAt: code.expiresAt,
    });
    logger.info(
      `Sign-in code published for ${user.selector} (valid until ${new Date(code.expiresAt).toISOString()})`,
    );
    schedulePoll(user, link);
    onChange();
  }

  function schedulePoll(user, link) {
    link.timer = setTimer(() => {
      const running = poll(user, link)
        .catch((err) =>
          logger.warn(`Sign-in poll failed for ${user.selector}: ${err?.message ?? err}`),
        )
        .finally(() => inflight.delete(running));
      inflight.add(running);
    }, link.intervalMs);
  }

  async function poll(user, link) {
    if (links.get(user.selector) !== link) {
      return;
    }
    if (now() >= link.code.expiresAt) {
      await expired(user, link);
      return;
    }
    let tokenSet;
    try {
      tokenSet = await auth.pollDeviceCode(getConfig(), link.code.deviceCode);
    } catch (err) {
      if (links.get(user.selector) !== link) {
        return;
      }
      if (
        err?.kind === AUTH_ERRORS.PENDING ||
        err?.kind === AUTH_ERRORS.UNREACHABLE ||
        err?.kind === AUTH_ERRORS.SERVER
      ) {
        schedulePoll(user, link);
        return;
      }
      if (err?.kind === AUTH_ERRORS.SLOW_DOWN) {
        link.intervalMs += 5000;
        schedulePoll(user, link);
        return;
      }
      if (err?.kind === AUTH_ERRORS.EXPIRED) {
        await expired(user, link);
        return;
      }
      stopLink(user.selector);
      const kind = err?.kind === AUTH_ERRORS.DECLINED ? 'declined' : 'app';
      links.set(user.selector, { state: USER_STATES.LINK_STOPPED, renewals: 0, error: kind });
      await showLink(user, kind);
      onChange();
      return;
    }
    if (links.get(user.selector) !== link) {
      return;
    }
    stopLink(user.selector);
    tokens.set(user.selector, { accessToken: tokenSet.accessToken, expiresAt: tokenSet.expiresAt });
    await store.update(user.selector, {
      refresh_token: tokenSet.refreshToken,
      linked_at: now(),
      last_error: undefined,
    });
    let who = { name: '', email: '' };
    try {
      who = await graphFor(user.selector).me();
      await store.update(user.selector, { account: who });
    } catch (err) {
      logger.warn(`Profile not read for ${user.selector}: ${err?.message ?? err}`);
    }
    logger.info(`${user.selector} signed in to Microsoft`);
    events.linked(user, who.email || who.name);
    await hideLink(user.selector);
    onChange();
    const account = accounts.get(user.selector) ?? { user, config: {} };
    await enqueue(() => syncAccount(account, { force: true }));
  }

  async function expired(user, link) {
    if ((link.renewals ?? 0) < MAX_CODE_RENEWALS) {
      await startLink(user, { renewal: true });
      return;
    }
    stopLink(user.selector);
    links.set(user.selector, {
      state: USER_STATES.LINK_STOPPED,
      renewals: link.renewals,
      error: 'expired',
    });
    await showLink(user, 'expired');
    onChange();
  }

  // --- Tokens ----------------------------------------------------------------------
  async function accessToken(selector, force) {
    const cached = tokens.get(selector);
    if (!force && cached && cached.expiresAt - TOKEN_MARGIN_MS > now()) {
      return cached.accessToken;
    }
    const refreshToken = store.get(selector).refresh_token;
    if (!refreshToken) {
      throw new AuthError(AUTH_ERRORS.REVOKED, 'Not signed in');
    }
    const fresh = await auth.refresh(getConfig(), refreshToken);
    tokens.set(selector, { accessToken: fresh.accessToken, expiresAt: fresh.expiresAt });
    // Microsoft rotates the refresh token: keep the new one.
    await store.update(selector, { refresh_token: fresh.refreshToken });
    return fresh.accessToken;
  }

  function graphFor(selector) {
    return createGraphClient({
      getToken: (force) => accessToken(selector, force),
      baseUrl: graphUrl,
      fetchImpl,
      sleep,
      now,
    });
  }

  // --- Sync --------------------------------------------------------------------------
  async function syncAccount(account, { force = false } = {}) {
    const { user } = account;
    const selector = user.selector;
    accounts.set(selector, account);
    const entry = store.get(selector);
    if (!entry.refresh_token) {
      if (!links.has(selector)) {
        events.linkRequired(user, 'new');
        await startLink(user);
      }
      return { calendars: 0, events: 0 };
    }
    const config = getConfig();
    const values = normalizeAccount(account.config);
    const language = languageOf(user);
    const graph = graphFor(selector);
    try {
      const who = entry.account ?? { name: '', email: '' };
      const outlook = await graph.listCalendars();
      const calendars = outlook.map((calendar) => ({
        source: calendar,
        input: toCalendar(gladys, selector, calendar, who, language),
      }));
      const ids = new Set(calendars.map((calendar) => calendar.input.external_id));

      // Calendars gone on the Outlook side, and the sign-in calendar.
      for (const existing of await gladys.getCalendars(selector)) {
        if (existing.external_id === linkCalendarId(selector)) {
          await write(() => gladys.deleteCalendar(existing.external_id)).catch(() => {});
          pushed.delete(existing.external_id);
        } else if (
          existing.external_id.startsWith(calendarPrefix(selector)) &&
          !ids.has(existing.external_id)
        ) {
          await write(() => gladys.deleteCalendar(existing.external_id)).catch(() => {});
          pushed.delete(existing.external_id);
        }
      }

      const inputs = calendars.map((calendar) => calendar.input).slice(0, 50);
      const calendarsKey = JSON.stringify(inputs);
      if (inputs.length > 0 && (force || pushed.get(`calendars:${selector}`) !== calendarsKey)) {
        await write(() => gladys.publishCalendars(selector, inputs));
        pushed.set(`calendars:${selector}`, calendarsKey);
      }

      const skipped = new Set(
        (await gladys.getCalendars(selector))
          .filter((calendar) => !calendar.sync)
          .map((calendar) => calendar.external_id),
      );
      const fromDate = new Date(now() - HISTORY_DAYS * 86_400_000);
      fromDate.setUTCHours(0, 0, 0, 0);
      const toDate = new Date(now());
      toDate.setUTCMonth(toDate.getUTCMonth() + config.horizon_months);
      const from = fromDate.toISOString();
      const to = toDate.toISOString();
      // The window moves once a day, not at every sync: the same content stays unchanged.
      const windowKey = from.slice(0, 10);

      let total = 0;
      for (const { source, input } of calendars.slice(0, 50)) {
        if (skipped.has(input.external_id)) {
          continue;
        }
        let list;
        try {
          list = await graph.calendarView(source.id, from, to);
        } catch (err) {
          if (
            err instanceof GraphError &&
            [GRAPH_ERRORS.FORBIDDEN, GRAPH_ERRORS.NOT_FOUND].includes(err.kind)
          ) {
            logger.info(
              `Calendar "${input.name}" of ${selector} not readable (${err.kind}): skipped`,
            );
            continue;
          }
          throw err;
        }
        const seen = new Set();
        const converted = [];
        for (const event of list) {
          const result = toEvent(gladys, selector, event, values, language);
          if (result && !seen.has(result.external_id)) {
            seen.add(result.external_id);
            converted.push(result);
          }
        }
        converted.sort((a, b) => instant(a.start) - instant(b.start));
        const kept = converted.slice(0, MAX_EVENTS_PER_CALENDAR);
        total += kept.length;
        const key = JSON.stringify([windowKey, kept]);
        if (!force && pushed.get(input.external_id) === key) {
          continue;
        }
        try {
          for (const chunk of splitWindow(kept, from, to)) {
            await write(() =>
              gladys.publishCalendarEvents(input.external_id, chunk.events, {
                from: chunk.from,
                to: chunk.to,
              }),
            );
          }
          pushed.set(input.external_id, key);
        } catch (err) {
          // The user turned this calendar off meanwhile.
          if (err?.status === 403) {
            continue;
          }
          throw err;
        }
      }
      await store.update(selector, {
        last_sync: now(),
        calendars: inputs.length,
        events: total,
        last_error: undefined,
      });
      events.success(user);
      onChange();
      return { calendars: inputs.length, events: total };
    } catch (err) {
      const revoked =
        (err instanceof AuthError && err.kind === AUTH_ERRORS.REVOKED) ||
        (err instanceof GraphError && err.kind === GRAPH_ERRORS.AUTH);
      if (revoked) {
        logger.warn(`Microsoft no longer accepts the sign-in of ${selector}: ${err.message}`);
        tokens.delete(selector);
        await store.update(selector, { refresh_token: undefined });
        events.linkRequired(user, 'revoked');
        await startLink(user);
        return { calendars: 0, events: 0 };
      }
      const kind = err?.kind ?? 'server';
      logger.warn(`Sync of ${selector} failed (${kind}): ${err?.message ?? err}`);
      await store.update(selector, {
        last_error: { kind, message: clip(err?.message, 200), at: now() },
      });
      events.failure(user, { kind, message: err?.message ?? String(err) });
      onChange();
      return { calendars: 0, events: 0, error: kind };
    }
  }

  function enqueue(job) {
    const result = queue.then(job);
    queue = result.catch((err) => logger.warn(`Calendar job failed: ${err?.message ?? err}`));
    return result;
  }

  return {
    /**
     * @description Sync every enabled user (the periodic loop, a button).
     * @param {{ force?: boolean }} [options] - Push even what did not change.
     * @returns {Promise<{ users: number, calendars: number, events: number }>} The totals.
     */
    syncAll(options = {}) {
      return enqueue(async () => {
        const list = await gladys.getCalendarAccounts();
        const enabled = new Set(list.map((account) => account.user.selector));
        // Users who disabled while the integration was away.
        for (const selector of store.selectors()) {
          if (!enabled.has(selector)) {
            await forgetUser(selector);
          }
        }
        let calendars = 0;
        let count = 0;
        for (const account of list) {
          const result = await syncAccount(account, options);
          calendars += result.calendars;
          count += result.events;
        }
        return { users: list.length, calendars, events: count };
      });
    },

    /**
     * @description A user enabled, disabled, changed their values or a toggle.
     * @param {string} selector - The user selector.
     * @returns {Promise<void>} Done.
     */
    accountUpdated(selector) {
      return enqueue(async () => {
        const account = (await gladys.getCalendarAccounts()).find(
          (candidate) => candidate.user.selector === selector,
        );
        if (!account) {
          await forgetUser(selector);
          return;
        }
        // Saving "My calendars" again restarts a stopped sign-in.
        if (
          links.get(selector)?.state === USER_STATES.LINK_STOPPED ||
          links.get(selector)?.state === USER_STATES.NO_APP
        ) {
          stopLink(selector);
        }
        await syncAccount(account, { force: true });
      });
    },

    /**
     * @description The configuration changed: a sign-in waiting for an app starts.
     * @returns {Promise<void>} Done.
     */
    configChanged() {
      return enqueue(async () => {
        for (const [selector, link] of links) {
          if (link.state === USER_STATES.NO_APP || link.state === USER_STATES.LINK_STOPPED) {
            stopLink(selector);
          }
        }
        tokens.clear();
      });
    },

    /**
     * @description What the widget shows, per user.
     * @returns {Array<{ selector: string, first_name: string, state: string, entry: object, failing: object|null, code?: object }>} The users.
     */
    snapshot() {
      return [...accounts.values()].map(({ user }) => {
        const entry = store.get(user.selector);
        const link = links.get(user.selector);
        let state = USER_STATES.LINKED;
        if (link) {
          state = link.state;
        } else if (!entry.refresh_token) {
          state = USER_STATES.LINKING;
        } else if (events.state(user.selector)) {
          state = USER_STATES.ERROR;
        }
        return {
          selector: user.selector,
          first_name: user.first_name ?? '',
          state,
          entry: {
            last_sync: entry.last_sync ?? null,
            calendars: entry.calendars ?? 0,
            events: entry.events ?? 0,
            last_error: entry.last_error ?? null,
          },
          failing: events.state(user.selector),
        };
      });
    },

    /** @description Stop every timer (shutdown, tests). */
    stop() {
      for (const selector of [...links.keys()]) {
        stopLink(selector);
      }
    },

    /** @description Forget what was pushed (reconnection: the core may have changed). */
    forgetPushed() {
      pushed.clear();
    },

    /** @returns {Promise<void>} The polls and the queue drained (tests). */
    async idle() {
      do {
        await Promise.all([...inflight]);
        await queue;
      } while (inflight.size > 0);
    },
  };

  async function forgetUser(selector) {
    stopLink(selector);
    tokens.delete(selector);
    accounts.delete(selector);
    events.forget(selector);
    for (const key of [...pushed.keys()]) {
      if (key === `calendars:${selector}` || key.includes(`:${selector}:`)) {
        pushed.delete(key);
      }
    }
    await store.remove(selector);
    onChange();
  }
}
