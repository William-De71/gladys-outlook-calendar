// A fake Microsoft on 127.0.0.1: the identity platform (device code, token
// endpoint with the device grant and the refresh grant, refresh tokens
// rotated) and Microsoft Graph (/me, /me/calendars, calendarView with pages),
// close to the real answers. Mailboxes are plain data; `approve`,
// `decline`, `revoke` and `fail` drive it. Nothing starts on import.

import { createServer } from 'node:http';

export const CLIENT_ID = '11111111-2222-3333-4444-555555555555';

/**
 * @description An identity platform error.
 * @param {string} error - OAuth error.
 * @param {string} [description] - `AADSTS…: text`.
 * @returns {object} The body.
 */
export function oauthError(error, description = '') {
  return { error, error_description: description, error_codes: [] };
}

/**
 * @description Start the fake.
 * @param {object} [options] - `{ mailboxes: { name: { me, calendars, events: { calendarId: [event] } } } }`.
 * @returns {Promise<object>} The server and its controls.
 */
export async function startFakeMicrosoft({ mailboxes = {} } = {}) {
  const requests = [];
  const failures = [];
  const devices = new Map(); // device_code -> { mailbox | null, declined }
  const refreshTokens = new Map(); // refresh token -> mailbox
  const accessTokens = new Map(); // access token -> mailbox
  let serial = 0;

  const issue = (mailbox) => {
    serial += 1;
    const access = `at-${mailbox}-${serial}`;
    const refresh = `rt-${mailbox}-${serial}`;
    accessTokens.set(access, mailbox);
    refreshTokens.set(refresh, mailbox);
    return {
      token_type: 'Bearer',
      scope: 'Calendars.Read User.Read',
      expires_in: 3600,
      access_token: access,
      refresh_token: refresh,
    };
  };

  const server = createServer((req, res) => {
    let raw = '';
    req.on('data', (chunk) => {
      raw += chunk;
    });
    req.on('end', () => {
      const url = new URL(req.url, 'http://127.0.0.1');
      const form = Object.fromEntries(new URLSearchParams(raw));
      requests.push({
        method: req.method,
        path: url.pathname,
        query: Object.fromEntries(url.searchParams),
        form,
        auth: req.headers.authorization,
        prefer: req.headers.prefer,
      });
      const send = (status, body, headers = {}) => {
        res.writeHead(status, { 'Content-Type': 'application/json', ...headers });
        res.end(JSON.stringify(body));
      };
      const failure = failures.findIndex((entry) => url.pathname.includes(entry.match));
      if (failure >= 0) {
        const [entry] = failures.splice(failure, 1);
        send(entry.status, entry.body ?? {}, entry.headers);
        return;
      }

      // --- identity platform -------------------------------------------------
      const oauth = url.pathname.match(/^\/([^/]+)\/oauth2\/v2\.0\/(devicecode|token)$/);
      if (oauth && req.method === 'POST') {
        if (form.client_id !== CLIENT_ID) {
          send(
            400,
            oauthError(
              'unauthorized_client',
              `AADSTS700016: Application with identifier '${form.client_id}' was not found in the directory.`,
            ),
          );
          return;
        }
        if (oauth[2] === 'devicecode') {
          serial += 1;
          const deviceCode = `device-${serial}`;
          devices.set(deviceCode, { mailbox: null, declined: false });
          send(200, {
            device_code: deviceCode,
            user_code: `CODE${serial}`,
            verification_uri: 'https://microsoft.com/devicelogin',
            expires_in: 900,
            interval: 5,
            message: 'To sign in…',
          });
          return;
        }
        if (form.grant_type === 'urn:ietf:params:oauth:grant-type:device_code') {
          const device = devices.get(form.device_code);
          if (!device) {
            send(400, oauthError('expired_token', 'AADSTS70019: Verification code expired.'));
          } else if (device.declined) {
            send(400, oauthError('authorization_declined', 'AADSTS70000: declined'));
          } else if (!device.mailbox) {
            send(
              400,
              oauthError(
                'authorization_pending',
                'AADSTS70016: OAuth 2.0 device flow error. Authorization is pending.',
              ),
            );
          } else {
            devices.delete(form.device_code);
            send(200, issue(device.mailbox));
          }
          return;
        }
        if (form.grant_type === 'refresh_token') {
          const mailbox = refreshTokens.get(form.refresh_token);
          if (!mailbox) {
            send(
              400,
              oauthError(
                'invalid_grant',
                'AADSTS70008: The refresh token has expired due to inactivity.',
              ),
            );
            return;
          }
          refreshTokens.delete(form.refresh_token);
          send(200, issue(mailbox));
          return;
        }
        send(400, oauthError('unsupported_grant_type'));
        return;
      }

      // --- Microsoft Graph ------------------------------------------------------
      const token = String(req.headers.authorization ?? '').replace(/^Bearer /, '');
      const mailbox = accessTokens.get(token);
      if (!mailbox) {
        send(401, {
          error: { code: 'InvalidAuthenticationToken', message: 'Access token is empty.' },
        });
        return;
      }
      const box = mailboxes[mailbox];
      if (url.pathname === '/v1.0/me') {
        send(200, box.me);
        return;
      }
      if (url.pathname === '/v1.0/me/calendars') {
        send(200, { value: box.calendars });
        return;
      }
      const view = url.pathname.match(/^\/v1\.0\/me\/calendars\/([^/]+)\/calendarView$/);
      if (view) {
        const calendarId = decodeURIComponent(view[1]);
        if (!box.calendars.some((calendar) => calendar.id === calendarId)) {
          send(404, {
            error: {
              code: 'ErrorItemNotFound',
              message: 'The specified object was not found in the store.',
            },
          });
          return;
        }
        const from = Date.parse(url.searchParams.get('startDateTime'));
        const to = Date.parse(url.searchParams.get('endDateTime'));
        const list = (box.events[calendarId] ?? []).filter((event) => {
          const start = Date.parse(`${event.start.dateTime.slice(0, 19)}Z`);
          const end = Date.parse(`${event.end.dateTime.slice(0, 19)}Z`);
          return start < to && end > from;
        });
        const top = Number(url.searchParams.get('$top') ?? 10);
        const skip = Number(url.searchParams.get('$skip') ?? 0);
        const page = list.slice(skip, skip + top);
        const body = { value: page };
        if (skip + top < list.length) {
          const next = new URL(url.href);
          next.searchParams.set('$skip', String(skip + top));
          body['@odata.nextLink'] = `${baseUrl}${next.pathname}${next.search}`;
        }
        send(200, body);
        return;
      }
      send(404, { error: { code: 'NotFound', message: 'Unknown path' } });
    });
  });
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  const baseUrl = `http://127.0.0.1:${server.address().port}`;
  return {
    loginUrl: baseUrl,
    graphUrl: `${baseUrl}/v1.0`,
    mailboxes,
    requests,
    /** @returns {string[]} The device codes still waiting. */
    pending: () =>
      [...devices.entries()].filter(([, device]) => !device.mailbox).map(([code]) => code),
    /** @description The user of the latest device code signs in as `mailbox`. */
    approve(mailbox, deviceCode = [...devices.keys()].at(-1)) {
      devices.get(deviceCode).mailbox = mailbox;
    },
    decline(deviceCode = [...devices.keys()].at(-1)) {
      devices.get(deviceCode).declined = true;
    },
    /** @description Microsoft forgets every token of a mailbox (password changed…). */
    revoke(mailbox) {
      for (const map of [refreshTokens, accessTokens]) {
        for (const [key, value] of [...map]) {
          if (value === mailbox) {
            map.delete(key);
          }
        }
      }
    },
    /** @description The next request whose path contains `match` answers this. */
    fail(match, status, body, headers) {
      failures.push({ match, status, body, headers });
    },
    close: () =>
      new Promise((resolve) => {
        server.closeAllConnections?.();
        server.close(() => resolve());
      }),
  };
}

/**
 * @description A Graph event.
 * @param {string} id - Id.
 * @param {object} fields - Overrides; `start` / `end` as `YYYY-MM-DDTHH:mm:ss`.
 * @returns {object} The event.
 */
export function graphEvent(id, { start, end, ...rest } = {}) {
  return {
    id,
    subject: `Event ${id}`,
    isAllDay: false,
    isCancelled: false,
    sensitivity: 'normal',
    showAs: 'busy',
    responseStatus: { response: 'organizer' },
    location: { displayName: '' },
    bodyPreview: '',
    webLink: `https://outlook.live.com/calendar/item/${id}`,
    start: { dateTime: `${start}.0000000`, timeZone: 'UTC' },
    end: { dateTime: `${end}.0000000`, timeZone: 'UTC' },
    ...rest,
  };
}
