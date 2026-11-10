import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { openWopi, WopiError } from '../src/wopi.js';

/** A fake host's fetch: `info` for CheckFileInfo, and per operation a status (and headers) to answer with. */
function host({ info = {}, answers = {} } = {}) {
  const calls = [];
  const res = (status, { body = '', headers = {} } = {}) => ({
    ok: status >= 200 && status < 300,
    status,
    headers: { get: (k) => headers[k] ?? null },
    json: async () => JSON.parse(body),
    arrayBuffer: async () => new TextEncoder().encode(body).buffer,
  });
  const fetch = async (url, init = {}) => {
    const u = new URL(url);
    const op = init.headers?.['X-WOPI-Override'] || (u.pathname.endsWith('/contents') ? 'GET_FILE' : 'CHECK_FILE_INFO');
    calls.push({ op, token: u.searchParams.get('access_token'), lock: init.headers?.['X-WOPI-Lock'] ?? null });
    const a = answers[op];
    if (a) return res(a.status, a);
    if (op === 'CHECK_FILE_INFO') return res(200, { body: JSON.stringify({ BaseFileName: 'A.docx', UserCanWrite: true, SupportsLocks: true, SupportsUpdate: true, ...info }) });
    if (op === 'GET_FILE') return res(200, { body: 'PK' });
    return res(200, { headers: { 'X-WOPI-ItemVersion': '2' } });
  };
  return { fetch, calls };
}
const SRC = 'https://files.example.com/wopi/files/1';

test('a writable file is read and locked; a save goes under the lock; closing unlocks', async () => {
  const h = host();
  const doc = await openWopi(SRC, 'tok', { fetch: h.fetch });
  assert.equal(doc.canWrite, true);
  assert.equal(doc.locked, true);
  assert.deepEqual(h.calls.map((c) => c.op), ['CHECK_FILE_INFO', 'GET_FILE', 'LOCK']);
  assert.ok(h.calls.every((c) => c.token === 'tok'));
  assert.equal(await doc.save(new Uint8Array([1])), '2');
  const lock = h.calls[2].lock;
  assert.match(lock, /^odx-[0-9a-f]{32}$/);
  assert.equal(h.calls[3].op, 'PUT');
  assert.equal(h.calls[3].lock, lock);
  await doc.close();
  assert.deepEqual(h.calls.at(-1), { op: 'UNLOCK', token: 'tok', lock });
});

test('someone else\'s lock: read-only, their lock reported; a read-only user takes no lock', async () => {
  const taken = host({ answers: { LOCK: { status: 409, headers: { 'X-WOPI-Lock': 'theirs' } } } });
  const doc = await openWopi(SRC, 'tok', { fetch: taken.fetch });
  assert.equal(doc.canWrite, false);
  assert.equal(doc.lockedBy, 'theirs');
  await assert.rejects(doc.save(new Uint8Array([1])), WopiError);
  const reader = host({ info: { UserCanWrite: false } });
  const view = await openWopi(SRC, 'tok', { fetch: reader.fetch });
  assert.equal(view.canWrite, false);
  assert.ok(!reader.calls.some((c) => c.op === 'LOCK'));
  const viewing = await openWopi(SRC, 'tok', { fetch: host().fetch, lock: false });
  assert.equal(viewing.canWrite, false);
});

test('the host\'s refusals are WopiErrors with its status; addresses must be https (or this computer)', async () => {
  await assert.rejects(openWopi(SRC, 'bad', { fetch: host({ answers: { CHECK_FILE_INFO: { status: 401 } } }).fetch }), (e) => e instanceof WopiError && e.status === 401 && /access token was refused/.test(e.message) && !/bad/.test(e.message));
  const lost = host({ answers: { PUT: { status: 409, headers: { 'X-WOPI-Lock': 'someone' } } } });
  const doc = await openWopi(SRC, 'tok', { fetch: lost.fetch });
  await assert.rejects(doc.save(new Uint8Array([1])), (e) => e.status === 409 && e.lock === 'someone');
  await doc.close();
  await assert.rejects(openWopi('http://files.example.com/wopi/files/1', 'tok', { fetch: host().fetch }), /https/);
  await assert.doesNotReject(async () => (await openWopi('http://localhost:8080/wopi/files/1', 'tok', { fetch: host().fetch })).close());
  await assert.rejects(openWopi(SRC, '', { fetch: host().fetch }), /access token/);
});
