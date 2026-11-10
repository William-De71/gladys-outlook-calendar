import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  aadstsOf,
  AUTH_ERRORS,
  createAuthClient,
  kindOfAuthError,
  safeVerificationUri,
} from '../src/microsoft-auth.js';
import { CLIENT_ID, startFakeMicrosoft } from './helpers/fake-microsoft.js';

test('error kinds of the identity platform', () => {
  const kind = (error, description = '') =>
    kindOfAuthError(400, { error, error_description: description });
  assert.equal(kind('authorization_pending'), AUTH_ERRORS.PENDING);
  assert.equal(kind('slow_down'), AUTH_ERRORS.SLOW_DOWN);
  assert.equal(kind('authorization_declined'), AUTH_ERRORS.DECLINED);
  assert.equal(kind('expired_token'), AUTH_ERRORS.EXPIRED);
  assert.equal(kind('invalid_grant', 'AADSTS70008: expired'), AUTH_ERRORS.REVOKED);
  assert.equal(
    kind('invalid_request', 'AADSTS7000218: The request body must contain client_assertion'),
    AUTH_ERRORS.APP,
  );
  assert.equal(kind('unauthorized_client', 'AADSTS700016: not found'), AUTH_ERRORS.APP);
  assert.equal(kindOfAuthError(503, {}), AUTH_ERRORS.SERVER);
  assert.equal(
    aadstsOf('AADSTS9002331: Application is configured for use by Microsoft Account users only'),
    'AADSTS9002331',
  );
});

test('the verification address is only ever a Microsoft one', () => {
  assert.equal(
    safeVerificationUri('https://microsoft.com/devicelogin'),
    'https://microsoft.com/devicelogin',
  );
  assert.equal(
    safeVerificationUri('https://login.microsoftonline.com/common/oauth2/deviceauth'),
    'https://login.microsoftonline.com/common/oauth2/deviceauth',
  );
  assert.equal(
    safeVerificationUri('https://evil.example/devicelogin'),
    'https://microsoft.com/devicelogin',
  );
  assert.equal(
    safeVerificationUri('http://microsoft.com/devicelogin'),
    'https://microsoft.com/devicelogin',
  );
});

test('device code, polling and refresh with rotation, against the fake', async (t) => {
  const ms = await startFakeMicrosoft({ mailboxes: {} });
  t.after(() => ms.close());
  const auth = createAuthClient({ baseUrl: ms.loginUrl, now: () => 0 });
  const config = { client_id: CLIENT_ID, tenant: 'common' };
  const code = await auth.requestDeviceCode(config);
  assert.equal(code.expiresAt, 900_000);
  assert.equal(code.intervalMs, 5000);
  await assert.rejects(auth.pollDeviceCode(config, code.deviceCode), { kind: AUTH_ERRORS.PENDING });
  ms.approve('x');
  const first = await auth.pollDeviceCode(config, code.deviceCode);
  const second = await auth.refresh(config, first.refreshToken);
  assert.notEqual(second.refreshToken, first.refreshToken);
  await assert.rejects(auth.refresh(config, first.refreshToken), { kind: AUTH_ERRORS.REVOKED });
  await assert.rejects(auth.requestDeviceCode({ client_id: '', tenant: 'common' }), {
    kind: AUTH_ERRORS.APP,
  });
  const offline = createAuthClient({ baseUrl: 'http://127.0.0.1:1' });
  await assert.rejects(offline.requestDeviceCode(config), { kind: AUTH_ERRORS.UNREACHABLE });
});
