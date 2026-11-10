// -----------------------------------------------------------------------------
// Dashboard widget of the Outlook Calendar integration.
//
//   - `accounts`  who is linked: the users with a Microsoft account linked,
//                 the calendars and events synchronized (tiles); one row per
//                 user who enabled "My calendars", with their state (synced
//                 at…, waiting for sign-in, code expired, error); a "Sync now"
//                 button.
//
// Only first names and states: never an e-mail, an event or a calendar name
// (a dashboard is seen by the whole household). Every text exists in the four
// languages (the content is cached per language). Budget of a card: 8
// components, 1 main one, 6 tiles, 2 texts (1 body), 1 status, 4 buttons.
// -----------------------------------------------------------------------------

import { WIDGET_COLORS } from '@gladysassistant/integration-sdk';
import { USER_STATES } from './users.js';
import { formatNumber, t } from './i18n.js';
import { formatTime } from './time.js';

export const WIDGETS = { ACCOUNTS: 'accounts' };
export const WIDGET_ACTIONS = { SYNC: 'sync' };
const TTL_SECONDS = 120;
const MAX_ROWS = 10;

/**
 * @description Truncate a text to a bound, with an ellipsis.
 * @param {string} value - The text.
 * @param {number} max - The bound.
 * @returns {string} The text.
 */
function clip(value, max) {
  const text = String(value ?? '');
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

/**
 * @description The row of one user.
 * @param {object} user - From `users.snapshot()`.
 * @param {object} options - `{ language, timeZone, now }`.
 * @returns {{ label: string, value: string, color: string }} The status item.
 */
export function userRow(user, { language, timeZone, now }) {
  const label = clip(user.first_name || t('someone', language), 40);
  switch (user.state) {
    case USER_STATES.LINKING:
      return { label, value: clip(t('rowLinking', language), 40), color: WIDGET_COLORS.INFO };
    case USER_STATES.NO_APP:
      return { label, value: clip(t('rowNoApp', language), 40), color: WIDGET_COLORS.WARNING };
    case USER_STATES.LINK_STOPPED:
      return {
        label,
        value: clip(t('rowLinkStopped', language), 40),
        color: WIDGET_COLORS.WARNING,
      };
    case USER_STATES.ERROR:
      return {
        label,
        value: clip(
          t('rowError', language, { kind: t(`kind_${user.failing?.kind ?? 'server'}`, language) }),
          40,
        ),
        color: WIDGET_COLORS.DANGER,
      };
    default:
      return {
        label,
        value: clip(
          user.entry.last_sync
            ? t('rowSynced', language, {
                time: formatTime(user.entry.last_sync, { language, timeZone, now }),
              })
            : t('rowLinked', language),
          40,
        ),
        color: WIDGET_COLORS.SUCCESS,
      };
  }
}

/**
 * @description The "accounts" widget.
 * @param {object} options - Options.
 * @param {Array<object>} options.users - `users.snapshot()`.
 * @param {object} options.config - The normalized configuration.
 * @param {string} options.language - The viewer language.
 * @param {number} options.now - Now (ms).
 * @param {string} [options.timeZone] - The timezone of Gladys.
 * @returns {object} The content.
 */
export function buildAccountsWidget({ users, config, language, now, timeZone }) {
  const components = [];
  if (!config.client_id) {
    components.push({ type: 'text', variant: 'body', text: clip(t('widgetNoApp', language), 300) });
  } else if (users.length === 0) {
    components.push({
      type: 'text',
      variant: 'body',
      text: clip(t('widgetNoUser', language), 300),
    });
  }
  const linked = users.filter((user) =>
    [USER_STATES.LINKED, USER_STATES.ERROR].includes(user.state),
  );
  components.push(
    {
      type: 'value',
      label: clip(t('tileLinked', language), 24),
      value: linked.length,
      icon: 'users',
      color: WIDGET_COLORS.PRIMARY,
    },
    {
      type: 'value',
      label: clip(t('tileCalendars', language), 24),
      value: linked.reduce((sum, user) => sum + (user.entry.calendars ?? 0), 0),
      icon: 'calendar',
      color: WIDGET_COLORS.INFO,
    },
    {
      type: 'value',
      label: clip(t('tileEvents', language), 24),
      value: clip(
        formatNumber(
          linked.reduce((sum, user) => sum + (user.entry.events ?? 0), 0),
          language,
          0,
        ),
        12,
      ),
      icon: 'list',
      color: WIDGET_COLORS.NEUTRAL,
    },
  );
  if (users.length > 0) {
    const sorted = [...users].sort((a, b) => a.first_name.localeCompare(b.first_name));
    components.push({
      type: 'status',
      items: sorted.slice(0, MAX_ROWS).map((user) => userRow(user, { language, timeZone, now })),
    });
  }
  if (config.client_id && users.length > 0) {
    components.push({
      type: 'button',
      label: clip(t('buttonSync', language), 24),
      icon: 'refresh-cw',
      style: 'secondary',
      action: { key: WIDGET_ACTIONS.SYNC },
    });
  }
  return { ttl_seconds: TTL_SECONDS, components };
}
