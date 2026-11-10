// -----------------------------------------------------------------------------
// Configuration, at two levels:
//
//   - the integration (`config_schema`, the administrator): the Microsoft
//     Entra application every user signs in through, how often to sync, how
//     far ahead;
//   - each user (`account_schema`, "My calendars"): what to do with their
//     private and declined events.
//
// This module provides the defaults and normalizes what arrives, so the rest
// of the code never deals with `undefined` or a value a select no longer
// offers. Every select value is stored by users: never remove one.
// -----------------------------------------------------------------------------

/** How often the calendars of every linked user are read again (minutes). */
export const SYNC_MINUTES = ['5', '15', '30', '60'];
/** How far ahead events are read (months). */
export const HORIZON_MONTHS = ['3', '6', '12'];
/** What a private Outlook event becomes in Gladys. */
export const PRIVATE_EVENTS = ['busy', 'show', 'skip'];
/** What an event the user declined becomes. */
export const DECLINED_EVENTS = ['skip', 'show'];

/** Days of the past kept in Gladys (the calendar view of last weeks). */
export const HISTORY_DAYS = 30;

/** A tenant: `common`, `consumers`, `organizations`, a GUID or a domain. */
const TENANT_PATTERN =
  /^(common|consumers|organizations|[0-9a-f-]{36}|[a-z0-9-]+(\.[a-z0-9-]+)+)$/i;
/** An application (client) id: a GUID. */
const CLIENT_ID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// Defaults: they MUST stay consistent with the `default` values declared in
// the manifest (test/manifest.test.js checks it).
export const DEFAULT_CONFIG = {
  client_id: '',
  tenant: 'common',
  sync_minutes: '15',
  horizon_months: '12',
};

export const DEFAULT_ACCOUNT = {
  private_events: 'busy',
  declined_events: 'skip',
};

/**
 * @description Normalize the configuration of the integration.
 * @param {Record<string, unknown>} [raw] - Config returned by the SDK.
 * @returns {object} `{ client_id, tenant, sync_minutes, horizon_months, valid_client_id }`.
 */
export function normalizeConfig(raw = {}) {
  const config = { ...DEFAULT_CONFIG, ...(raw ?? {}) };
  const clientId = String(config.client_id ?? '').trim();
  const tenant = String(config.tenant ?? '').trim();
  return {
    client_id: CLIENT_ID_PATTERN.test(clientId) ? clientId.toLowerCase() : '',
    typed_client_id: clientId !== '',
    tenant: TENANT_PATTERN.test(tenant) ? tenant.toLowerCase() : DEFAULT_CONFIG.tenant,
    sync_minutes: Number(
      SYNC_MINUTES.includes(String(config.sync_minutes))
        ? config.sync_minutes
        : DEFAULT_CONFIG.sync_minutes,
    ),
    horizon_months: Number(
      HORIZON_MONTHS.includes(String(config.horizon_months))
        ? config.horizon_months
        : DEFAULT_CONFIG.horizon_months,
    ),
  };
}

/**
 * @description Normalize the values a user entered in "My calendars".
 * @param {Record<string, unknown>} [raw] - The account values.
 * @returns {{ private_events: string, declined_events: string }} The values.
 */
export function normalizeAccount(raw = {}) {
  const account = { ...DEFAULT_ACCOUNT, ...(raw ?? {}) };
  return {
    private_events: PRIVATE_EVENTS.includes(account.private_events)
      ? account.private_events
      : DEFAULT_ACCOUNT.private_events,
    declined_events: DECLINED_EVENTS.includes(account.declined_events)
      ? account.declined_events
      : DEFAULT_ACCOUNT.declined_events,
  };
}
