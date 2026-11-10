// -----------------------------------------------------------------------------
// SDK wiring: every handler of the integration, registered on the Gladys
// client given, before `connect()`. Kept out of index.js so the tests run it
// against the fake Gladys of the SDK (`createFakeGladys`) and fake Microsoft
// servers.
// -----------------------------------------------------------------------------

import { createLogger } from '@gladysassistant/integration-sdk';
import { normalizeConfig } from './config.js';
import { AUTH_ERRORS, createAuthClient, LOGIN_URL } from './microsoft-auth.js';
import { GRAPH_URL } from './graph.js';
import { createStore, ACCOUNTS_FILE } from './store.js';
import { createEvents } from './events.js';
import { createUsers } from './users.js';
import { buildAccountsWidget, WIDGET_ACTIONS, WIDGETS } from './widgets.js';
import { all } from './i18n.js';
import { usableTimeZone } from './time.js';

const logger = createLogger({ name: 'app' });

/** Keys of the manifest `actions` and `scene_actions`: stored by users, forever. */
export const ACTIONS = { CHECK_APP: 'check_app', SYNC_NOW: 'sync_now' };
export const SCENE_ACTIONS = { SYNC_NOW: 'sync_now' };

/** A widget is nudged at most this often (the core allows 1 / 10 s). */
const WIDGET_NUDGE_MS = 15_000;

/**
 * @description Register the handlers on a Gladys client, before `connect()`.
 * @param {object} gladys - A `GladysIntegration` (or the fake one of the SDK).
 * @param {object} [options] - Test seams.
 * @returns {object} Handles for the tests.
 */
export function createApp(
  gladys,
  {
    loginUrl = LOGIN_URL,
    graphUrl = GRAPH_URL,
    fetchImpl = globalThis.fetch,
    sleep,
    setTimer,
    clearTimer,
    accountsFile = ACCOUNTS_FILE,
    now = Date.now,
    timeZone = usableTimeZone(),
    autoSync = true,
  } = {},
) {
  let config = normalizeConfig();
  const getConfig = () => config;

  const auth = createAuthClient({ baseUrl: loginUrl, fetchImpl, now });
  const store = createStore({
    file: accountsFile,
    onError: (err) => logger.warn(`Accounts not saved: ${err?.message ?? err}`),
  });
  const events = createEvents({ publish: (key, data) => gladys.publishSceneEvent(key, data), now });

  // --- Widget nudges ----------------------------------------------------------
  let lastNudge = 0;
  const nudgeWidget = () => {
    if (now() - lastNudge < WIDGET_NUDGE_MS) {
      return;
    }
    lastNudge = now();
    Promise.resolve()
      .then(() => gladys.requestWidgetRefresh(WIDGETS.ACCOUNTS))
      .catch(() => {});
  };

  const users = createUsers({
    gladys,
    auth,
    store,
    events,
    getConfig,
    graphUrl,
    fetchImpl,
    sleep,
    setTimer,
    clearTimer,
    now,
    timeZone,
    onChange: nudgeWidget,
  });

  const setStatus = (ok, message) =>
    Promise.resolve()
      .then(() => gladys.setConnectionStatus(ok, ok ? undefined : message))
      .catch((err) => logger.warn(`Status not updated: ${err?.message ?? err}`));

  /**
   * @description Ask Microsoft for a sign-in code with the app of the
   * configuration (thrown away): the app exists, takes the device flow and
   * the account types of the tenant.
   * @returns {Promise<{ ok: boolean, message?: object }>} The result.
   */
  async function checkApp() {
    if (!config.client_id) {
      return {
        ok: false,
        message: all(config.typed_client_id ? 'statusBadClientId' : 'statusNoApp'),
      };
    }
    try {
      await auth.requestDeviceCode(config);
      return { ok: true };
    } catch (err) {
      logger.warn(`App check failed: ${err?.message ?? err}`);
      if (err?.kind === AUTH_ERRORS.UNREACHABLE || err?.kind === AUTH_ERRORS.SERVER) {
        return { ok: false, message: all('statusUnreachable') };
      }
      return {
        ok: false,
        message: all('statusBadApp', { detail: String(err?.message ?? err).slice(0, 140) }),
      };
    }
  }

  // --- Periodic sync ---------------------------------------------------------------
  let loop = null;
  function schedule() {
    if (loop) {
      clearInterval(loop);
    }
    loop = null;
    if (!autoSync) {
      return;
    }
    loop = setInterval(
      () => {
        users.syncAll().catch((err) => logger.warn(`Periodic sync failed: ${err?.message ?? err}`));
      },
      config.sync_minutes * 60 * 1000,
    );
    loop.unref?.();
  }

  // --- Calendar accounts: enable, disable, values, toggles ---------------------------
  gladys.onCalendarAccountUpdated(async (userSelector) => {
    logger.info(`Calendar account updated: ${userSelector}`);
    await users.accountUpdated(userSelector);
    nudgeWidget();
  });

  // --- Configuration buttons -------------------------------------------------------
  gladys.onAction(ACTIONS.CHECK_APP, async () => {
    logger.info('Action check_app');
    const { ok, message } = await checkApp();
    await setStatus(ok, message);
    return ok ? all('checkOk') : message;
  });
  gladys.onAction(ACTIONS.SYNC_NOW, async () => {
    logger.info('Action sync_now');
    const result = await users.syncAll({ force: true });
    return all('syncDone', result);
  });

  // --- Scene action -------------------------------------------------------------------
  gladys.onSceneAction(SCENE_ACTIONS.SYNC_NOW, async () => {
    const result = await users.syncAll();
    return { users: result.users, events: result.events };
  });

  // --- Dashboard widget ------------------------------------------------------------
  gladys.onWidgetGet(WIDGETS.ACCOUNTS, async ({ language } = {}) =>
    buildAccountsWidget({ users: users.snapshot(), config, language, now: now(), timeZone }),
  );
  gladys.onWidgetAction(WIDGETS.ACCOUNTS, async (actionKey) => {
    if (actionKey !== WIDGET_ACTIONS.SYNC) {
      throw new Error(`Unknown action ${actionKey}`);
    }
    users
      .syncAll({ force: true })
      .then(() => gladys.requestWidgetRefresh(WIDGETS.ACCOUNTS))
      .catch((err) => logger.warn(`Sync from the widget failed: ${err?.message ?? err}`));
    return all('toastSync');
  });

  // --- Configuration ---------------------------------------------------------------
  async function applyConfig(raw) {
    const previous = config;
    config = normalizeConfig(raw);
    logger.info(
      `Configuration: app ${config.client_id ? 'set' : 'missing'}, tenant ${config.tenant}, ` +
        `every ${config.sync_minutes} min, ${config.horizon_months} months ahead`,
    );
    if (previous.client_id !== config.client_id || previous.tenant !== config.tenant) {
      await users.configChanged();
    }
    const { ok, message } = await checkApp();
    await setStatus(ok, message);
    schedule();
  }

  gladys.onConfigUpdated(async (raw) => {
    await applyConfig(raw);
    await users.syncAll({ force: true });
    nudgeWidget();
  });

  // Each (re)connection re-reads the configuration and every account, which
  // may have changed while the integration was away (the account events are
  // lost while disconnected). The promise lets the tests wait for it.
  let startup = Promise.resolve();
  gladys.on('connected', () => {
    startup = (async () => {
      await store.load();
      await applyConfig(await gladys.getConfig());
      users.forgetPushed();
      await users.syncAll({ force: true });
      nudgeWidget();
    })().catch(async (err) => {
      logger.error('Post-connection initialization failed', err);
      await setStatus(false, all('statusInitFailed'));
    });
  });

  // --- Graceful shutdown -------------------------------------------------------------
  gladys.handleShutdown(async (signal) => {
    logger.info(`Received ${signal} -> graceful shutdown`);
    if (loop) {
      clearInterval(loop);
    }
    users.stop();
    await store.flush();
  });

  return {
    getConfig,
    users,
    store,
    events,
    whenReady: () => startup,
    stop: () => {
      if (loop) {
        clearInterval(loop);
      }
      users.stop();
    },
  };
}
