// A page that edits a document a WOPI host serves: open it as wopi.html#src=<file's WOPI URL>&access_token=<token>.
// For trying it out: the fragment isn't sent to servers, and it's taken out of the address bar here, but the visit as
// first opened stays in the browser's history. A host builds a page of its own and hands it the token the way the
// protocol suggests (a form POST, or a message from the host's frame), with a short lifetime.
import { mountWopiEditor } from '../src/wopi.js';
import { exportPdf } from '../src/pdf.js';

const status = document.getElementById('status');
const params = new URLSearchParams(location.hash.slice(1));
history.replaceState(null, '', location.pathname); // (the token out of the address bar)
const say = (t) => { status.textContent = t; };
mountWopiEditor(document.getElementById('app'), {
  src: params.get('src'),
  accessToken: params.get('access_token'),
  pdf: exportPdf,
  autosaveMs: Number(params.get('autosave') ?? 2000),
  lockRefreshMs: params.get('refresh') ? Number(params.get('refresh')) : undefined,
  onChange: () => say('Editing…'),
  onSaved: () => say('Saved to the document store.'),
  onError: (e) => say(e.message),
}).then((app) => {
  window.wopiApp = app; // (for tests)
  say(app.wopi.canWrite ? `${app.wopi.info.BaseFileName}: editing, saved as you go.` : app.wopi.lockedBy != null ? `${app.wopi.info.BaseFileName}: someone else is editing it, so it's open read-only.` : `${app.wopi.info.BaseFileName}: read-only.`);
}, (e) => { say(e.message); window.wopiError = e.message; });
