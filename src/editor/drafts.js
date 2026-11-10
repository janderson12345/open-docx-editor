// Unsaved changes kept in the browser (IndexedDB, this site's own storage), so a closed tab or a crash doesn't lose them:
// the edited .docx a moment after each change, cleared when the document is saved. Opt in with mountDocxEditor's
// `drafts` option (it keeps document content on this computer). Every call fails quietly where storage isn't
// available (a private window, storage turned off): there is then simply no draft.

const DB = 'open-docx-editor';
const STORE = 'drafts';

function db() {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') return reject(new Error('no IndexedDB'));
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: 'key' });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function run(mode, fn) {
  const d = await db();
  try {
    return await new Promise((resolve, reject) => {
      const tx = d.transaction(STORE, mode);
      const req = fn(tx.objectStore(STORE));
      tx.oncomplete = () => resolve(req?.result);
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  } finally { d.close(); }
}

/** Keeps `bytes` (the .docx) as the draft for `key`. */
export async function saveDraft(key, bytes, name) {
  try { await run('readwrite', (s) => s.put({ key, bytes, name, savedAt: Date.now() })); } catch { /* no storage: no draft */ }
}

/** The draft for `key`: { key, bytes, name, savedAt } or null. */
export async function loadDraft(key) {
  try { return (await run('readonly', (s) => s.get(key))) || null; } catch { return null; }
}

/** Forgets the draft for `key`. */
export async function clearDraft(key) {
  try { await run('readwrite', (s) => s.delete(key)); } catch { /* nothing to clear */ }
}

/** Forgets every draft older than `maxAge` ms (done as an editor with drafts starts, so they don't pile up). */
export async function pruneDrafts(maxAge) {
  const cutoff = Date.now() - maxAge;
  try {
    await run('readwrite', (s) => {
      const req = s.openCursor();
      req.onsuccess = () => {
        const c = req.result;
        if (!c) return;
        if (!(c.value.savedAt > cutoff)) c.delete();
        c.continue();
      };
      return req;
    });
  } catch { /* nothing to prune */ }
}
