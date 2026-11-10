// -----------------------------------------------------------------------------
// Scene triggers, per Gladys user, once per transition:
//
//   - `link_required`   a user must sign in to Microsoft: they just enabled
//                       "My calendars" (`new`), or Microsoft no longer accepts
//                       their sign-in (`revoked`: password changed, access
//                       removed, 90 days unused);
//   - `linked`          a user signed in: their calendars are coming;
//   - `sync_error`      reading Outlook failed for a user (once per kind);
//   - `sync_recovered`  the first good sync after a `sync_error`.
//
// The core's own `calendar.event-is-coming` trigger already covers the
// events themselves: these are about the link with Microsoft.
// -----------------------------------------------------------------------------

import { createLogger } from '@gladysassistant/integration-sdk';

const logger = createLogger({ name: 'events' });

export const TRIGGERS = {
  LINK_REQUIRED: 'link_required',
  LINKED: 'linked',
  SYNC_ERROR: 'sync_error',
  SYNC_RECOVERED: 'sync_recovered',
};

/**
 * @description Build the trigger state machine.
 * @param {object} options - Options.
 * @param {(key: string, data: object) => Promise<unknown>} options.publish - `gladys.publishSceneEvent`.
 * @param {() => number} [options.now] - Clock.
 * @returns {object} The hooks.
 */
export function createEvents({ publish, now = Date.now }) {
  // selector -> { kind, since }
  const failing = new Map();

  const send = (key, data) => {
    Promise.resolve()
      .then(() => publish(key, data))
      .catch((err) => logger.warn(`Scene event ${key} not published: ${err?.message ?? err}`));
  };

  return {
    /**
     * @description A user must sign in.
     * @param {{ selector: string, first_name: string }} user - The user.
     * @param {'new'|'revoked'} reason - Why.
     */
    linkRequired(user, reason) {
      send(TRIGGERS.LINK_REQUIRED, {
        user: user.selector,
        first_name: user.first_name ?? '',
        reason,
      });
    },

    /**
     * @description A user signed in.
     * @param {{ selector: string, first_name: string }} user - The user.
     * @param {string} account - The Microsoft account (e-mail).
     */
    linked(user, account) {
      failing.delete(user.selector);
      send(TRIGGERS.LINKED, {
        user: user.selector,
        first_name: user.first_name ?? '',
        account: account ?? '',
      });
    },

    /**
     * @description A sync failed.
     * @param {{ selector: string, first_name: string }} user - The user.
     * @param {{ kind: string, message: string }} error - The error.
     */
    failure(user, error) {
      const current = failing.get(user.selector);
      if (current && current.kind === error.kind) {
        return;
      }
      failing.set(user.selector, { kind: error.kind, since: current?.since ?? now() });
      send(TRIGGERS.SYNC_ERROR, {
        user: user.selector,
        first_name: user.first_name ?? '',
        kind: error.kind,
        message: String(error.message ?? '').slice(0, 300),
      });
    },

    /**
     * @description A sync succeeded.
     * @param {{ selector: string, first_name: string }} user - The user.
     */
    success(user) {
      const current = failing.get(user.selector);
      if (!current) {
        return;
      }
      failing.delete(user.selector);
      send(TRIGGERS.SYNC_RECOVERED, {
        user: user.selector,
        first_name: user.first_name ?? '',
        down_minutes: Math.max(0, Math.round((now() - current.since) / 60000)),
      });
    },

    /** @description Forget a user (disabled). */
    forget(selector) {
      failing.delete(selector);
    },

    /**
     * @description The failure of a user, for the widget.
     * @param {string} selector - The user selector.
     * @returns {object|null} `{ kind, since }`.
     */
    state(selector) {
      return failing.get(selector) ?? null;
    },
  };
}
