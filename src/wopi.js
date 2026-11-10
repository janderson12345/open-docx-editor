// WOPI client: opening, locking and saving a document a WOPI host serves (the protocol SharePoint, OneDrive and other
// document stores use to let a web editor work on their files; Microsoft documents it openly as [MS-WOPI]). The host
// gives the page the file's WOPI URL (…/wopi/files/<id>) and an access token; everything else goes through these calls:
//
//   CheckFileInfo  GET  <file>                        the file's name, size, version, and what this user may do
//   GetFile        GET  <file>/contents               the .docx
//   Lock           POST <file>  X-WOPI-Override: LOCK          while it's open for editing (30 minutes, refreshed)
//   RefreshLock    POST <file>  X-WOPI-Override: REFRESH_LOCK  every 15 minutes, as Word for the web does
//   Unlock         POST <file>  X-WOPI-Override: UNLOCK        on closing
//   PutFile        POST <file>/contents  X-WOPI-Override: PUT  saving, under the lock
//
// The access token goes in the access_token query parameter, as the protocol says. The host must allow this page's
// origin (CORS) on those URLs, with the X-WOPI-* request and response headers. Not covered: proof keys (a host that
// requires them checks a signature a browser can't hold the key for), PutRelativeFile (Save As), and the host frame's
// PostMessage interface.
import { mountDocxEditor } from './index.js';

const LOCK_REFRESH_MS = 15 * 60 * 1000;

/** A WOPI error, with the host's status and, for a lock conflict, the lock the file is held under. */
export class WopiError extends Error {
  constructor(message, { status = 0, lock = null, reason = null } = {}) {
    super(message);
    this.name = 'WopiError';
    this.status = status;
    this.lock = lock;
    this.reason = reason;
  }
}

/**
 * Opens a file a WOPI host serves: reads its information and contents, and takes the lock when this user may edit it.
 * `src` is the file's WOPI URL, `accessToken` the host's token for this user. Options: `fetch` (a fetch function),
 * `lock` (false to open without locking, as for viewing), `lockRefreshMs`, `onLockError(error)` (a lock refresh refused).
 * Returns { info, bytes, canWrite, locked, lockedBy, save(bytes), refreshLock(), close() }.
 */
export async function openWopi(src, accessToken, { fetch: f = globalThis.fetch.bind(globalThis), lock = true, lockRefreshMs = LOCK_REFRESH_MS, onLockError = null } = {}) {
  const base = new URL(src);
  if (base.protocol !== 'https:' && !/^(localhost|127\.0\.0\.1|\[::1\])$/.test(base.hostname)) throw new WopiError('The WOPI address must use https.');
  if (!accessToken) throw new WopiError('A WOPI access token is needed.');
  const url = (suffix = '') => {
    const u = new URL(base);
    u.pathname = u.pathname.replace(/\/$/, '') + suffix;
    u.searchParams.set('access_token', accessToken);
    return u.href;
  };
  const call = async (what, suffix, init) => {
    let res;
    try { res = await f(url(suffix), init); } catch (e) { throw new WopiError(`${what}: the document store couldn't be reached (${e.message}).`); }
    if (res.ok) return res;
    const held = res.headers.get('X-WOPI-Lock');
    const reason = res.headers.get('X-WOPI-LockFailureReason');
    const why = { 401: 'the access token was refused', 404: 'the file was not found', 409: 'the file is locked by someone else', 412: 'the file changed on the server', 413: 'the file is too large for the document store', 500: 'the document store failed', 501: 'the document store does not support it' }[res.status] || `the document store answered ${res.status}`;
    throw new WopiError(`${what}: ${why}.`, { status: res.status, lock: held, reason });
  };
  const override = (op, headers = {}) => ({ method: 'POST', headers: { 'X-WOPI-Override': op, ...headers } });

  const info = await (await call('Opening', '', { method: 'GET' })).json();
  const bytes = new Uint8Array(await (await call('Reading the file', '/contents', { method: 'GET' })).arrayBuffer());
  const mayWrite = !!info.UserCanWrite && !info.ReadOnly && info.SupportsUpdate !== false;
  const lockId = mayWrite && info.SupportsLocks ? newLockId() : null;
  lockRefreshMs = Math.max(1000, Number(lockRefreshMs) || LOCK_REFRESH_MS); // (never a stream of refreshes)
  let locked = false;
  let lockedBy = null;
  let timer = null;
  if (lockId && lock) {
    try {
      await call('Locking the file', '', override('LOCK', { 'X-WOPI-Lock': lockId }));
      locked = true;
      // (A refresh refused, the lock lost to someone else, is reported: saves will fail until it's sorted out.)
      timer = setInterval(() => doc.refreshLock().catch((e) => onLockError && onLockError(e)), lockRefreshMs);
    } catch (e) {
      if (!(e instanceof WopiError) || e.status !== 409) throw e;
      lockedBy = e.lock || '';
    }
  }
  // Editable when the host allows it and, where it locks files, we hold the lock.
  const canWrite = mayWrite && (!lockId || locked) && lock !== false;
  let closed = false;
  const doc = {
    info,
    bytes,
    canWrite,
    get locked() { return locked; },
    lockedBy,
    /** Saves the document's bytes to the host (PutFile). Resolves with the host's new version, if it gives one. */
    async save(data) {
      if (!canWrite) throw new WopiError('This document is open read-only.');
      if (closed) throw new WopiError('This document has been closed.');
      const headers = { 'X-WOPI-Override': 'PUT', 'Content-Type': 'application/octet-stream' };
      if (locked) headers['X-WOPI-Lock'] = lockId;
      const res = await call('Saving', '/contents', { method: 'POST', headers, body: data });
      return res.headers.get('X-WOPI-ItemVersion');
    },
    /** Renews the lock (done every 15 minutes while open). */
    async refreshLock() {
      if (!locked || closed) return;
      await call('Keeping the file locked', '', override('REFRESH_LOCK', { 'X-WOPI-Lock': lockId }));
    },
    /** Takes the lock again after close() (a page come back from the browser's back/forward cache). */
    async reopen() {
      if (!closed) return;
      closed = false;
      if (!lockId || !lock) return;
      await call('Locking the file', '', override('LOCK', { 'X-WOPI-Lock': lockId }));
      locked = true;
      clearInterval(timer);
      timer = setInterval(() => doc.refreshLock().catch((e) => onLockError && onLockError(e)), lockRefreshMs);
    },
    /** Releases the lock. `keepalive` lets the request finish as the page closes. */
    async close({ keepalive = false } = {}) {
      if (closed) return;
      closed = true;
      clearInterval(timer);
      if (locked) {
        locked = false;
        await call('Unlocking the file', '', { ...override('UNLOCK', { 'X-WOPI-Lock': lockId }), keepalive });
      }
    },
  };
  return doc;
}

/** A lock id of our own: random, so two editors of the same file never share one. */
function newLockId() {
  const b = new Uint8Array(16);
  crypto.getRandomValues(b);
  return 'odx-' + [...b].map((x) => x.toString(16).padStart(2, '0')).join('');
}

/**
 * Opens a WOPI host's file in an editor in `container`: read-only when the user may not edit it or someone else has it
 * locked. Changes are saved back to the host as Word for the web does, on their own a moment after editing stops
 * (`autosaveMs`, 2 seconds; 0 for never), at once with Ctrl+S, and with `save()`; File > Create a Copy still downloads
 * a copy. The lock is released when the page closes.
 * Options: `src`, `accessToken`, `autosaveMs`, `onSaved(version)`, `onError(error)`, `fetch`, `lockRefreshMs`, and any
 * of mountDocxEditor's (the author defaults to the host's UserFriendlyName).
 * Returns { editor, wopi, save(), close() }.
 */
export async function mountWopiEditor(container, { src, accessToken, autosaveMs = 2000, onSaved, onError, onChange, fetch, lockRefreshMs, ...options } = {}) {
  if (autosaveMs > 0) autosaveMs = Math.max(500, autosaveMs); // (not a save a keystroke)
  const fail = (e) => (onError ? onError(e) : console.error(e));
  const wopi = await openWopi(src, accessToken, { fetch, lockRefreshMs, lock: !options.readOnly, onLockError: fail });
  const name = wopi.info.BaseFileName || 'document.docx';
  let saving = null;
  let again = false;
  let timer = null;
  let edits = 0; // edits made, so a save knows whether any came while its bytes were on the way
  const save = async () => {
    clearTimeout(timer);
    if (saving) { again = true; return saving; } // (one save at a time; changes made during it are saved after it)
    saving = (async () => {
      try {
        // (Only the host's own document: never one opened or made in its place.)
        if (editor.fileName !== name) throw new WopiError(`Not saved: ${editor.fileName} isn't the document this page opened.`);
        const at = edits;
        const version = await wopi.save(editor.save());
        // Saved only once the host says so, and only if nothing changed meanwhile (marking it saved also forgets
        // a draft kept in the browser, which must outlast a failed upload).
        if (edits === at) editor.markSaved();
        if (onSaved) onSaved(version);
        return version;
      } finally {
        saving = null;
        if (again) { again = false; save().catch(fail); }
      }
    })();
    return saving;
  };
  const editor = mountDocxEditor(container, {
    author: wopi.info.UserFriendlyName || 'Author',
    ...options,
    fileName: name,
    readOnly: options.readOnly || !wopi.canWrite,
    // One file: no New or Open in the File menu, so nothing else can be saved over it.
    features: { ...options.features, newDocument: false, openFile: false },
    onError: fail,
    onChange: (view) => {
      edits++;
      if (onChange) onChange(view);
      if (wopi.canWrite && autosaveMs > 0) { clearTimeout(timer); timer = setTimeout(() => save().catch(fail), autosaveMs); }
    },
  });
  try {
    await editor.open(wopi.bytes, name);
  } catch (e) {
    // (A file that won't open: the lock taken for it released, not left held for everyone else.)
    editor.destroy();
    await wopi.close().catch(() => {});
    throw e;
  }
  const onKey = (e) => {
    if ((e.ctrlKey || e.metaKey) && !e.altKey && e.key.toLowerCase() === 's') {
      e.preventDefault();
      if (wopi.canWrite) save().catch(fail);
    }
  };
  container.addEventListener('keydown', onKey);
  // Leaving the page: with unsaved edits, the browser asks first while they're saved; the lock is released as it goes,
  // and taken again if the page comes back from the back/forward cache (read-only if someone else has it by then).
  const onLeave = (e) => {
    if (!editor.dirty || !wopi.canWrite) return;
    save().catch(fail);
    e.preventDefault();
    e.returnValue = '';
  };
  const onHide = () => { wopi.close({ keepalive: true }).catch(() => {}); };
  const onShow = (e) => {
    if (!e.persisted) return;
    wopi.reopen().catch((err) => { editor.setReadOnly(true); fail(err); });
  };
  window.addEventListener('beforeunload', onLeave);
  window.addEventListener('pagehide', onHide);
  window.addEventListener('pageshow', onShow);
  return {
    editor,
    wopi,
    save,
    /** Saves what's unsaved, releases the lock and removes the editor. */
    async close() {
      clearTimeout(timer);
      window.removeEventListener('beforeunload', onLeave);
      window.removeEventListener('pagehide', onHide);
      window.removeEventListener('pageshow', onShow);
      container.removeEventListener('keydown', onKey);
      try {
        editor.view?.odx.closeStories?.(); // (a header, note or text box still being edited, put back first)
        // Every save through, the one under way and any that edits during it queued, before the lock goes.
        while (wopi.canWrite && (saving || editor.dirty)) await (saving || save());
      } finally {
        editor.destroy();
        await wopi.close();
      }
    },
  };
}
