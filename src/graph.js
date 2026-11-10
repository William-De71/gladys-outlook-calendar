// -----------------------------------------------------------------------------
// Client of Microsoft Graph v1.0 (graph.microsoft.com), read-only, with the
// access token of one user:
//
//   - GET /me                                   who signed in (shown to them)
//   - GET /me/calendars                         their calendars (shared ones too)
//   - GET /me/calendars/{id}/calendarView       the events of a window, with
//                                               the occurrences of recurring
//                                               events expanded by Microsoft
//
// Times come back in UTC (`Prefer: outlook.timezone="UTC"`). Pages follow
// `@odata.nextLink`. A 429 or a 503 waits as `Retry-After` asks (bounded); a
// 401 asks for a fresh token once. Errors are typed (`kind`).
// -----------------------------------------------------------------------------

export const GRAPH_URL = 'https://graph.microsoft.com/v1.0';

export const GRAPH_ERRORS = {
  /** The token is refused even fresh: sign in again. */
  AUTH: 'auth',
  /** No access to this calendar (a shared one withdrawn…): skip it. */
  FORBIDDEN: 'forbidden',
  /** The calendar or the mailbox does not exist (any more). */
  NOT_FOUND: 'not_found',
  THROTTLED: 'throttled',
  SERVER: 'server',
  UNREACHABLE: 'unreachable',
};

const TIMEOUT_MS = 30_000;
const MAX_RETRY_WAIT_MS = 30_000;
const MAX_PAGES = 60;
const PAGE_SIZE = 250;

export class GraphError extends Error {
  /**
   * @param {string} kind - One of GRAPH_ERRORS.
   * @param {string} message - What went wrong (never a token).
   * @param {number|null} [status] - HTTP status.
   */
  constructor(kind, message, status = null) {
    super(message);
    this.name = 'GraphError';
    this.kind = kind;
    this.status = status;
  }
}

/**
 * @description The kind of an HTTP failure.
 * @param {number} status - HTTP status.
 * @returns {string} One of GRAPH_ERRORS.
 */
export function kindOfGraphStatus(status) {
  if (status === 401) {
    return GRAPH_ERRORS.AUTH;
  }
  if (status === 403) {
    return GRAPH_ERRORS.FORBIDDEN;
  }
  if (status === 404) {
    return GRAPH_ERRORS.NOT_FOUND;
  }
  if (status === 429) {
    return GRAPH_ERRORS.THROTTLED;
  }
  return GRAPH_ERRORS.SERVER;
}

/**
 * @description The wait asked by a `Retry-After` header.
 * @param {string|null} value - Seconds, or an HTTP date.
 * @param {() => number} now - Clock.
 * @returns {number} Milliseconds (1 s when absent).
 */
export function retryAfterMs(value, now = Date.now) {
  if (!value) {
    return 1000;
  }
  const seconds = Number(value);
  if (Number.isFinite(seconds)) {
    return Math.max(0, seconds * 1000);
  }
  const date = Date.parse(value);
  return Number.isFinite(date) ? Math.max(0, date - now()) : 1000;
}

/**
 * @description Build the client.
 * @param {object} options - Options.
 * @param {(force: boolean) => Promise<string>} options.getToken - The access token of the user (`force`: refresh it).
 * @param {string} [options.baseUrl] - The Graph root.
 * @param {typeof fetch} [options.fetchImpl] - fetch.
 * @param {(ms: number) => Promise<void>} [options.sleep] - Waits.
 * @param {() => number} [options.now] - Clock.
 * @returns {{ me: Function, listCalendars: Function, calendarView: Function }} The client.
 */
export function createGraphClient({
  getToken,
  baseUrl = GRAPH_URL,
  fetchImpl = globalThis.fetch,
  sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
  now = Date.now,
}) {
  async function once(url, token) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
    try {
      const response = await fetchImpl(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          Prefer: 'outlook.timezone="UTC"',
        },
        signal: controller.signal,
      });
      const text = await response.text();
      let body = null;
      try {
        body = text ? JSON.parse(text) : null;
      } catch {
        body = null;
      }
      return { response, body };
    } catch (err) {
      throw new GraphError(
        GRAPH_ERRORS.UNREACHABLE,
        controller.signal.aborted
          ? 'Microsoft Graph did not answer in time'
          : `Cannot reach Microsoft Graph: ${err?.cause?.code ?? err?.message ?? err}`,
      );
    } finally {
      clearTimeout(timer);
    }
  }

  async function get(url) {
    let token = await getToken(false);
    let refreshed = false;
    for (let attempt = 1; attempt <= 4; attempt += 1) {
      const { response, body } = await once(url, token);
      if (response.ok) {
        return body ?? {};
      }
      const kind = kindOfGraphStatus(response.status);
      if (kind === GRAPH_ERRORS.AUTH && !refreshed) {
        refreshed = true;
        token = await getToken(true);
        continue;
      }
      if ((kind === GRAPH_ERRORS.THROTTLED || response.status === 503) && attempt < 4) {
        const wait = retryAfterMs(response.headers.get('retry-after'), now);
        if (wait <= MAX_RETRY_WAIT_MS) {
          await sleep(wait);
          continue;
        }
      }
      const message = String(body?.error?.message ?? response.statusText ?? '').slice(0, 300);
      throw new GraphError(kind, `HTTP ${response.status}: ${message || 'error'}`, response.status);
    }
    throw new GraphError(GRAPH_ERRORS.THROTTLED, 'Microsoft Graph kept asking to wait', 429);
  }

  async function all(url) {
    const items = [];
    let next = url;
    for (let page = 0; next && page < MAX_PAGES; page += 1) {
      // Only follow links back to Graph itself.
      if (!next.startsWith(baseUrl)) {
        break;
      }
      const body = await get(next);
      if (Array.isArray(body.value)) {
        items.push(...body.value);
      }
      next = typeof body['@odata.nextLink'] === 'string' ? body['@odata.nextLink'] : null;
    }
    return items;
  }

  return {
    /**
     * @description Who signed in.
     * @returns {Promise<{ name: string, email: string }>} The account.
     */
    async me() {
      const body = await get(`${baseUrl}/me?$select=displayName,mail,userPrincipalName`);
      return {
        name: String(body.displayName ?? '').slice(0, 100),
        email: String(body.mail || body.userPrincipalName || '').slice(0, 200),
      };
    },

    /**
     * @description The calendars of the user.
     * @returns {Promise<Array<object>>} Graph calendar objects.
     */
    listCalendars() {
      return all(
        `${baseUrl}/me/calendars?$select=id,name,color,hexColor,isDefaultCalendar,owner&$top=100`,
      );
    },

    /**
     * @description The events of a calendar between two instants, recurrences expanded.
     * @param {string} calendarId - The Graph calendar id.
     * @param {string} from - ISO instant.
     * @param {string} to - ISO instant.
     * @returns {Promise<Array<object>>} Graph event objects.
     */
    calendarView(calendarId, from, to) {
      const select = [
        'id',
        'subject',
        'start',
        'end',
        'isAllDay',
        'isCancelled',
        'location',
        'webLink',
        'sensitivity',
        'showAs',
        'responseStatus',
        'bodyPreview',
        'isOnlineMeeting',
        'onlineMeeting',
      ].join(',');
      const query = new URLSearchParams({
        startDateTime: from,
        endDateTime: to,
        $select: select,
        $top: String(PAGE_SIZE),
      });
      return all(`${baseUrl}/me/calendars/${encodeURIComponent(calendarId)}/calendarView?${query}`);
    },
  };
}
