// A WOPI host for tests: files kept in memory, CheckFileInfo, GetFile, PutFile, Lock, RefreshLock and Unlock as
// [MS-WOPI] describes them (a lock lasts 30 minutes; a call under the wrong lock is refused with 409 and the lock it's
// held under), access tokens checked, CORS for an editor page on another origin, and a log of every call.
import http from 'node:http';

/**
 * Starts a host. `files`: { id: { name, bytes, token, user?, canWrite?, readOnly? } }. Resolves with
 * { url(id), files, log, lockOf(id), setLock(id, lockId), close() }.
 */
export function startWopiHost(files) {
  const locks = new Map(); // id -> { id, until }
  const log = [];
  const lockOf = (id) => { const l = locks.get(id); return l && l.until > Date.now() ? l.id : null; };
  const server = http.createServer((req, res) => {
    const u = new URL(req.url, 'http://host');
    const cors = { 'Access-Control-Allow-Origin': req.headers.origin || '*', 'Access-Control-Allow-Headers': 'X-WOPI-Override, X-WOPI-Lock, Content-Type', 'Access-Control-Expose-Headers': 'X-WOPI-Lock, X-WOPI-LockFailureReason, X-WOPI-ItemVersion', 'Access-Control-Allow-Methods': 'GET, POST' };
    if (req.method === 'OPTIONS') { res.writeHead(204, cors); return res.end(); }
    const m = /^\/wopi\/files\/([^/]+)(\/contents)?$/.exec(u.pathname);
    const send = (status, body = '', headers = {}) => { res.writeHead(status, { ...cors, ...headers }); res.end(body); };
    const file = m && files[m[1]];
    const op = req.headers['x-wopi-override'] || (req.method === 'GET' ? (m?.[2] ? 'GET_FILE' : 'CHECK_FILE_INFO') : '?');
    const lockHeader = req.headers['x-wopi-lock'] || null;
    log.push({ op, id: m?.[1], lock: lockHeader, token: u.searchParams.get('access_token') === file?.token });
    if (!file) return send(404);
    if (u.searchParams.get('access_token') !== file.token) return send(401);
    const id = m[1];
    const current = lockOf(id);
    const conflict = (reason) => send(409, '', { 'X-WOPI-Lock': current || '', 'X-WOPI-LockFailureReason': reason });
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => {
      if (op === 'CHECK_FILE_INFO') {
        return send(200, JSON.stringify({ BaseFileName: file.name, Size: file.bytes.length, Version: String(file.version || 1), OwnerId: 'owner', UserId: file.user || 'user', UserFriendlyName: file.user || 'Test User', UserCanWrite: file.canWrite !== false, ReadOnly: !!file.readOnly, SupportsLocks: true, SupportsUpdate: true }), { 'Content-Type': 'application/json' });
      }
      if (op === 'GET_FILE') return send(200, file.bytes, { 'Content-Type': 'application/octet-stream', 'X-WOPI-ItemVersion': String(file.version || 1) });
      if (op === 'LOCK') {
        if (current && current !== lockHeader) return conflict('Locked by another editor');
        locks.set(id, { id: lockHeader, until: Date.now() + 30 * 60 * 1000 });
        return send(200);
      }
      if (op === 'REFRESH_LOCK' || op === 'UNLOCK') {
        if (current !== lockHeader) return conflict(current ? 'Locked by another editor' : 'Not locked');
        if (op === 'UNLOCK') locks.delete(id); else locks.get(id).until = Date.now() + 30 * 60 * 1000;
        return send(200);
      }
      if (op === 'PUT') {
        if (file.canWrite === false || file.readOnly) return send(401);
        // (An unlocked file may be written without a lock only while it's empty, as the protocol says.)
        if (current !== lockHeader && !(current == null && file.bytes.length === 0)) return conflict(current ? 'Locked by another editor' : 'Not locked');
        file.bytes = Buffer.concat(chunks);
        file.version = (file.version || 1) + 1;
        return send(200, '', { 'X-WOPI-ItemVersion': String(file.version) });
      }
      send(501);
    });
  });
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve({
    url: (id) => `http://127.0.0.1:${server.address().port}/wopi/files/${id}`,
    files,
    log,
    lockOf,
    setLock: (id, lockId) => locks.set(id, { id: lockId, until: Date.now() + 30 * 60 * 1000 }),
    dropLock: (id) => locks.delete(id),
    close: () => new Promise((r) => server.close(r)),
  })));
}
