// -----------------------------------------------------------------------------
// What the integration keeps per Gladys user, in /data/accounts.json (the only
// writable folder, file mode 600): the Microsoft refresh token of the user,
// who signed in, and the state of the last sync (for the widget). Access
// tokens stay in memory only.
//
// A user who disables "My calendars" is forgotten here (their calendars were
// already destroyed by the core). A missing or broken file starts empty: the
// users then sign in again.
// -----------------------------------------------------------------------------

import { chmod, mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';

export const ACCOUNTS_FILE = '/data/accounts.json';

/**
 * @description Build the store.
 * @param {object} [options] - Options.
 * @param {string} [options.file] - The JSON file.
 * @param {(err: Error) => void} [options.onError] - Write failures.
 * @returns {object} The store.
 */
export function createStore({ file = ACCOUNTS_FILE, onError = () => {} } = {}) {
  let users = {};
  let writing = Promise.resolve();

  async function write() {
    await mkdir(dirname(file), { recursive: true });
    const temp = `${file}.tmp`;
    await writeFile(temp, JSON.stringify({ users }), { mode: 0o600 });
    await chmod(temp, 0o600).catch(() => {});
    await rename(temp, file);
  }

  const save = () => {
    writing = writing.then(write).catch(onError);
    return writing;
  };

  return {
    /** @description Read the file (missing or broken: start empty). */
    async load() {
      try {
        const parsed = JSON.parse(await readFile(file, 'utf8'));
        users = parsed && typeof parsed.users === 'object' && parsed.users ? parsed.users : {};
      } catch {
        users = {};
      }
    },

    /**
     * @description One user.
     * @param {string} selector - The user selector.
     * @returns {object} Their entry (a copy), `{}` when unknown.
     */
    get(selector) {
      return { ...(users[selector] ?? {}) };
    },

    /**
     * @description Change one user and save.
     * @param {string} selector - The user selector.
     * @param {object} patch - Fields to set (`undefined` removes one).
     * @returns {Promise<void>} Saved.
     */
    update(selector, patch) {
      const entry = { ...(users[selector] ?? {}) };
      for (const [key, value] of Object.entries(patch)) {
        if (value === undefined) {
          delete entry[key];
        } else {
          entry[key] = value;
        }
      }
      users[selector] = entry;
      return save();
    },

    /**
     * @description Forget a user.
     * @param {string} selector - The user selector.
     * @returns {Promise<void>} Saved.
     */
    remove(selector) {
      delete users[selector];
      return save();
    },

    /** @returns {string[]} The users known. */
    selectors() {
      return Object.keys(users);
    },

    /** @returns {Promise<void>} The pending write. */
    flush() {
      return writing;
    },
  };
}
