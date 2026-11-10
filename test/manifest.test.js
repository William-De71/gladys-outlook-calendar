import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, statSync } from 'node:fs';
import { TRIGGERS } from '../src/events.js';
import { ACTIONS, SCENE_ACTIONS } from '../src/app.js';
import { WIDGETS } from '../src/widgets.js';
import {
  DECLINED_EVENTS,
  DEFAULT_ACCOUNT,
  DEFAULT_CONFIG,
  HORIZON_MONTHS,
  PRIVATE_EVENTS,
  SYNC_MINUTES,
} from '../src/config.js';
import { SCOPES } from '../src/microsoft-auth.js';

const read = (path) => JSON.parse(readFileSync(new URL(path, import.meta.url), 'utf8'));
const manifest = read('../gladys-assistant-integration.json');
const pkg = read('../package.json');
const byKey = (list) => new Map((list ?? []).map((entry) => [entry.key, entry]));
const values = (field) => field.options.map((option) => option.value);
const source = (path) => readFileSync(new URL(path, import.meta.url), 'utf8');
const REPO = 'William-De71/gladys-outlook-calendar';

test('a calendar provider, in lockstep with package.json, at 0.x until 1.0.0 is decided', () => {
  assert.equal(manifest.type, 'calendar');
  assert.equal(manifest.name, 'Outlook Calendar');
  assert.equal(manifest.version, pkg.version);
  assert.match(pkg.version, /^0\./);
  assert.equal(
    manifest.docker_image,
    `ghcr.io/william-de71/gladys-outlook-calendar:${pkg.version}`,
  );
  assert.equal(manifest.cover_image, `https://raw.githubusercontent.com/${REPO}/main/cover.jpg`);
  assert.equal(pkg.license, 'Apache-2.0');
  assert.equal(pkg.name, 'gladys-outlook-calendar');
  assert.deepEqual(Object.keys(pkg.dependencies), ['@gladysassistant/integration-sdk']);
});

test('declares the cloud transport: Microsoft only, two modules talk to it', () => {
  assert.deepEqual(manifest.transports, ['cloud']);
  assert.match(source('../src/microsoft-auth.js'), /https:\/\/login\.microsoftonline\.com/);
  assert.match(source('../src/graph.js'), /https:\/\/graph\.microsoft\.com\/v1\.0/);
  for (const file of [
    'app',
    'users',
    'convert',
    'store',
    'events',
    'widgets',
    'config',
    'time',
    'i18n',
  ]) {
    assert.doesNotMatch(
      source(`../src/${file}.js`),
      /fetch\(|node:https?|node:net|node:dgram/,
      file,
    );
  }
  assert.doesNotMatch(source('../src/app.js'), /publishTransports|GLADYS_PREFER_LOCAL/);
});

test('read-only: only the scopes needed, never a write scope', () => {
  assert.deepEqual(SCOPES, ['offline_access', 'User.Read', 'Calendars.Read']);
  for (const file of ['graph', 'users']) {
    assert.doesNotMatch(source(`../src/${file}.js`), /method: '(POST|PATCH|PUT|DELETE)'/, file);
  }
});

test('requires Gladys 5.3.0, on the services shelf, asks no permission', () => {
  assert.equal(manifest.gladys_version, '>=5.3.0');
  assert.deepEqual(manifest.categories, ['services']);
  for (const key of [
    'location',
    'network_wake',
    'containers',
    'network_discovery',
    'webhooks',
    'messaging',
    'energy_contracts',
    'music',
  ]) {
    assert.equal(manifest[key], undefined, key);
  }
});

test('the configuration matches the code: keys, select values, defaults', () => {
  const config = byKey(manifest.config_schema);
  assert.equal(config.get('client_id').type, 'string');
  assert.equal(config.get('client_id').required, true);
  assert.equal(config.get('tenant').default, DEFAULT_CONFIG.tenant);
  assert.deepEqual(values(config.get('sync_minutes')), SYNC_MINUTES);
  assert.equal(config.get('sync_minutes').default, DEFAULT_CONFIG.sync_minutes);
  assert.deepEqual(values(config.get('horizon_months')), HORIZON_MONTHS);
  assert.equal(config.get('horizon_months').default, DEFAULT_CONFIG.horizon_months);
  assert.deepEqual(
    manifest.config_schema.filter((f) => f.type !== 'section').map((f) => f.key),
    Object.keys(DEFAULT_CONFIG),
  );
  // Per user: no oauth2 / account_link (not allowed there), no secret needed.
  const account = byKey(manifest.account_schema);
  assert.deepEqual(values(account.get('private_events')), PRIVATE_EVENTS);
  assert.equal(account.get('private_events').default, DEFAULT_ACCOUNT.private_events);
  assert.deepEqual(values(account.get('declined_events')), DECLINED_EVENTS);
  assert.equal(account.get('declined_events').default, DEFAULT_ACCOUNT.declined_events);
  for (const field of [...manifest.config_schema, ...manifest.account_schema]) {
    assert.ok(!['text', 'password', 'oauth2', 'account_link'].includes(field.type), `${field.key}`);
    assert.equal(field.display_if, undefined, field.key);
  }
});

test('declares exactly the buttons, widget, triggers and scene action the code handles', () => {
  const app = source('../src/app.js');
  assert.deepEqual(
    manifest.actions.map((a) => a.key),
    Object.values(ACTIONS),
  );
  for (const name of Object.keys(ACTIONS)) {
    assert.match(app, new RegExp(`onAction\\(ACTIONS\\.${name}`), name);
  }
  assert.deepEqual(
    manifest.widgets.map((w) => w.key),
    Object.values(WIDGETS),
  );
  assert.match(app, /onWidgetGet\(WIDGETS\.ACCOUNTS/);
  assert.match(app, /onWidgetAction\(WIDGETS\.ACCOUNTS/);
  assert.deepEqual(
    manifest.scene_triggers.map((t) => t.key),
    Object.values(TRIGGERS),
  );
  assert.deepEqual(
    manifest.scene_actions.map((a) => a.key),
    Object.values(SCENE_ACTIONS),
  );
  assert.match(app, /onSceneAction\(SCENE_ACTIONS\.SYNC_NOW/);
  assert.match(app, /onCalendarAccountUpdated\(/);
  for (const trigger of manifest.scene_triggers) {
    for (const field of trigger.fields ?? []) {
      assert.ok(
        trigger.variables.some((variable) => variable.key === field.key),
        `${trigger.key}.${field.key}`,
      );
    }
  }
  assert.deepEqual(
    manifest.scene_actions[0].outputs.map((o) => o.key),
    ['users', 'events'],
  );
});

test('limits of the store: counts, timeouts, sizes', () => {
  assert.ok(manifest.actions.length <= 10);
  assert.ok(manifest.widgets.length >= 1 && manifest.widgets.length <= 5);
  assert.ok(manifest.scene_triggers.length <= 20);
  for (const action of [...manifest.actions, ...manifest.scene_actions]) {
    assert.ok(action.timeout_seconds >= 5 && action.timeout_seconds <= 120, action.key);
  }
  const size = statSync(new URL('../gladys-assistant-integration.json', import.meta.url)).size;
  assert.ok(size < 100 * 1024, `${size} bytes`);
  const cover = readFileSync(new URL('../cover.jpg', import.meta.url));
  assert.ok(cover.length < 150 * 1024, `cover.jpg: ${cover.length} bytes`);
  assert.equal(cover.readUInt16BE(0), 0xffd8, 'cover.jpg is a JPEG');
});

test('catalog of Gladys 5.3.0: search keywords within the store bounds, worldwide', () => {
  const { keywords } = manifest;
  assert.ok(Array.isArray(keywords) && keywords.length >= 1 && keywords.length <= 10);
  const normalized = keywords.map((keyword) => keyword.trim().toLowerCase());
  for (const keyword of normalized) {
    assert.ok(keyword.length >= 2 && keyword.length <= 30, keyword);
  }
  assert.equal(new Set(normalized).size, normalized.length, 'unique once trimmed and lowercased');
  assert.ok(!normalized.includes(manifest.name.toLowerCase()), 'never the name itself');
  assert.equal(manifest.countries, undefined, 'usable worldwide: no country list');
});
