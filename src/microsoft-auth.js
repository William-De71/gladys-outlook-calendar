// -----------------------------------------------------------------------------
// Sign-in with a Microsoft account: the OAuth 2.0 device authorization grant
// of the Microsoft identity platform (learn.microsoft.com, "OAuth 2.0 device
// authorization grant"), against the administrator's own app registration
// (a public client: no secret, no redirect URI).
//
//   1. POST /{tenant}/oauth2/v2.0/devicecode → a short `user_code` to type at
//      `verification_uri` (microsoft.com/devicelogin), valid ~15 minutes;
//   2. POST /{tenant}/oauth2/v2.0/token with the device code, every
//      `interval` seconds, until the user signed in (or declined, or the code
//      expired);
//   3. the refresh token (scope `offline_access`) renews the one-hour access
//      token; Microsoft rotates it at every use, so the new one is kept.
//
// Plain fetch, form-encoded bodies. Errors are typed (`kind`) and never quote
// a token.
// -----------------------------------------------------------------------------

export const LOGIN_URL = 'https://login.microsoftonline.com';
export const DEFAULT_VERIFICATION_URI = 'https://microsoft.com/devicelogin';
/** Read-only: Gladys never writes back to the provider. */
export const SCOPES = ['offline_access', 'User.Read', 'Calendars.Read'];
const DEVICE_GRANT = 'urn:ietf:params:oauth:grant-type:device_code';
const TIMEOUT_MS = 20_000;

export const AUTH_ERRORS = {
  /** The app registration is missing, unknown or refuses the device flow. */
  APP: 'app',
  /** Still waiting for the user. */
  PENDING: 'pending',
  /** Poll less often. */
  SLOW_DOWN: 'slow_down',
  DECLINED: 'declined',
  EXPIRED: 'expired',
  /** The refresh token no longer works: the user must sign in again. */
  REVOKED: 'revoked',
  UNREACHABLE: 'unreachable',
  SERVER: 'server',
};

export class AuthError extends Error {
  /**
   * @param {string} kind - One of AUTH_ERRORS.
   * @param {string} message - What went wrong (never a token).
   * @param {string} [code] - The OAuth `error`, or the AADSTS code.
   */
  constructor(kind, message, code = null) {
    super(message);
    this.name = 'AuthError';
    this.kind = kind;
    this.code = code;
  }
}

/**
 * @description The AADSTS code of an error description.
 * @param {string} description - `AADSTS700016: Application with identifier…`.
 * @returns {string} `AADSTS700016`, or ''.
 */
export function aadstsOf(description) {
  return String(description ?? '').match(/AADSTS\d+/)?.[0] ?? '';
}

/**
 * @description The kind of an OAuth error answer.
 * @param {number} status - HTTP status.
 * @param {object} body - `{ error, error_description }`.
 * @returns {string} One of AUTH_ERRORS.
 */
export function kindOfAuthError(status, body) {
  const error = String(body?.error ?? '');
  const code = aadstsOf(body?.error_description);
  if (error === 'authorization_pending') {
    return AUTH_ERRORS.PENDING;
  }
  if (error === 'slow_down') {
    return AUTH_ERRORS.SLOW_DOWN;
  }
  if (error === 'authorization_declined' || error === 'access_denied') {
    return AUTH_ERRORS.DECLINED;
  }
  if (error === 'expired_token' || error === 'bad_verification_code') {
    return AUTH_ERRORS.EXPIRED;
  }
  // The app: unknown in the directory (700016), public client flows off
  // (7000218), unknown tenant (90002), personal accounts not allowed
  // (9002331, 50020), invalid client id format (700038).
  if (
    error === 'invalid_client' ||
    error === 'unauthorized_client' ||
    [
      'AADSTS700016',
      'AADSTS7000218',
      'AADSTS90002',
      'AADSTS9002331',
      'AADSTS50020',
      'AADSTS700038',
    ].includes(code)
  ) {
    return AUTH_ERRORS.APP;
  }
  if (
    error === 'invalid_grant' ||
    error === 'interaction_required' ||
    error === 'consent_required'
  ) {
    return AUTH_ERRORS.REVOKED;
  }
  return status >= 500 ? AUTH_ERRORS.SERVER : AUTH_ERRORS.APP;
}

/**
 * @description The verification address, only a Microsoft one.
 * @param {unknown} value - `verification_uri` of the answer.
 * @returns {string} The address to show.
 */
export function safeVerificationUri(value) {
  try {
    const url = new URL(String(value));
    const host = url.hostname.toLowerCase();
    const microsoft =
      host === 'microsoft.com' ||
      host.endsWith('.microsoft.com') ||
      host.endsWith('.microsoftonline.com');
    return url.protocol === 'https:' && microsoft ? url.href : DEFAULT_VERIFICATION_URI;
  } catch {
    return DEFAULT_VERIFICATION_URI;
  }
}

/**
 * @description Build the client.
 * @param {object} [options] - Options.
 * @param {string} [options.baseUrl] - The login root.
 * @param {typeof fetch} [options.fetchImpl] - fetch.
 * @param {() => number} [options.now] - Clock.
 * @returns {{ requestDeviceCode: Function, pollDeviceCode: Function, refresh: Function }} The client.
 */
export function createAuthClient({
  baseUrl = LOGIN_URL,
  fetchImpl = globalThis.fetch,
  now = Date.now,
} = {}) {
  async function post(tenant, endpoint, form) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
    let response;
    let text;
    try {
      response = await fetchImpl(
        `${baseUrl}/${encodeURIComponent(tenant)}/oauth2/v2.0/${endpoint}`,
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
            Accept: 'application/json',
          },
          body: new URLSearchParams(form).toString(),
          signal: controller.signal,
        },
      );
      text = await response.text();
    } catch (err) {
      throw new AuthError(
        AUTH_ERRORS.UNREACHABLE,
        controller.signal.aborted
          ? 'Microsoft did not answer in time'
          : `Cannot reach Microsoft: ${err?.cause?.code ?? err?.message ?? err}`,
      );
    } finally {
      clearTimeout(timer);
    }
    let body;
    try {
      body = text ? JSON.parse(text) : {};
    } catch {
      body = {};
    }
    if (!response.ok) {
      const kind = kindOfAuthError(response.status, body);
      const code = aadstsOf(body.error_description) || String(body.error ?? '') || null;
      const description = String(body.error_description ?? body.error ?? `HTTP ${response.status}`)
        .split(/\r?\n/)[0]
        .slice(0, 300);
      throw new AuthError(kind, description, code);
    }
    return body;
  }

  /**
   * @description The tokens of an answer.
   * @param {object} body - The token answer.
   * @returns {{ accessToken: string, refreshToken: string, expiresAt: number }} The tokens.
   */
  function tokensOf(body) {
    if (typeof body.access_token !== 'string' || typeof body.refresh_token !== 'string') {
      throw new AuthError(
        AUTH_ERRORS.SERVER,
        'Microsoft answered without tokens (is offline_access granted?)',
      );
    }
    const seconds = Number(body.expires_in);
    return {
      accessToken: body.access_token,
      refreshToken: body.refresh_token,
      expiresAt: now() + (Number.isFinite(seconds) && seconds > 0 ? seconds : 3600) * 1000,
    };
  }

  return {
    /**
     * @description Start a sign-in.
     * @param {{ client_id: string, tenant: string }} config - The configuration.
     * @returns {Promise<{ deviceCode: string, userCode: string, verificationUri: string, expiresAt: number, intervalMs: number }>} The code.
     */
    async requestDeviceCode(config) {
      if (!config.client_id) {
        throw new AuthError(AUTH_ERRORS.APP, 'No application (client) ID in the configuration');
      }
      const body = await post(config.tenant, 'devicecode', {
        client_id: config.client_id,
        scope: SCOPES.join(' '),
      });
      if (typeof body.device_code !== 'string' || typeof body.user_code !== 'string') {
        throw new AuthError(AUTH_ERRORS.SERVER, 'Microsoft answered without a device code');
      }
      const expires = Number(body.expires_in);
      const interval = Number(body.interval);
      return {
        deviceCode: body.device_code,
        userCode: body.user_code.slice(0, 20),
        verificationUri: safeVerificationUri(body.verification_uri),
        expiresAt: now() + (Number.isFinite(expires) && expires > 0 ? expires : 900) * 1000,
        intervalMs: (Number.isFinite(interval) && interval > 0 ? interval : 5) * 1000,
      };
    },

    /**
     * @description Ask once whether the user signed in.
     * @param {{ client_id: string, tenant: string }} config - The configuration.
     * @param {string} deviceCode - The device code.
     * @returns {Promise<{ accessToken: string, refreshToken: string, expiresAt: number }>} The tokens;
     * throws `pending` / `slow_down` while waiting.
     */
    async pollDeviceCode(config, deviceCode) {
      return tokensOf(
        await post(config.tenant, 'token', {
          grant_type: DEVICE_GRANT,
          client_id: config.client_id,
          device_code: deviceCode,
        }),
      );
    },

    /**
     * @description Renew the access token.
     * @param {{ client_id: string, tenant: string }} config - The configuration.
     * @param {string} refreshToken - The current refresh token.
     * @returns {Promise<{ accessToken: string, refreshToken: string, expiresAt: number }>} The new tokens.
     */
    async refresh(config, refreshToken) {
      return tokensOf(
        await post(config.tenant, 'token', {
          grant_type: 'refresh_token',
          client_id: config.client_id,
          refresh_token: refreshToken,
          scope: SCOPES.join(' '),
        }),
      );
    },
  };
}
