// open-docx-editor public API
export { DocxSession } from './ooxml/docx.js';
export { blankDocxBytes } from './ooxml/template.js';
export { setXmlImpl } from './ooxml/xml.js';
export { schema } from './editor/schema.js';
export { createEditor, saveEditor } from './editor/editor.js';
export { createToolbar } from './editor/toolbar.js';
export * as commands from './editor/commands.js';
export * as track from './editor/track.js';
export { listFields, fieldValues, setFieldValues, setForms } from './editor/fields.js';
export { readFieldValues, fillFields } from './fieldsNode.js';
export { mailMerge, mergeFieldNames, formatMergeValue } from './merge.js';

/**
 * One-call setup: toolbar + editor in `container`.
 * `author` names tracked changes made here. `onOpened(timings, fileName)` runs after each open; `onError(error)` when a file from the Open button can't be opened. `onSave(bytes, fileName)` replaces the default download (e.g. upload to your server).
 * @returns {{ open(bytes|File|Blob|Response, name?), newDocument(), save(): Uint8Array, download(name?), dirty: boolean, destroy(), view }}
 */
import { DocxSession as Session } from './ooxml/docx.js';
import { loadEmbeddedFonts } from './ooxml/fonts.js';
import { createEditor as make } from './editor/editor.js';
import { createToolbar as makeToolbar } from './editor/toolbar.js';
import { listFields, fieldValues, setFieldValues as setFV, setForms as setFormsMode } from './editor/fields.js';
import { missingFonts, fontsSettled } from './editor/fontcheck.js';
import { saveDraft, loadDraft, clearDraft, pruneDrafts } from './editor/drafts.js';
import { updateFields } from './editor/refs.js';
import { editorAgent } from './agent.js';
import { setProtection } from './editor/protect.js';
import { announce } from './editor/live.js';
import * as mm from './editor/mailmerge.js';
import { mailMerge as mergeDocx, mergeFieldNames } from './merge.js';

const DRAFT_DAYS = 14;
/** A fingerprint of a file's bytes (FNV-1a and the length): which file a draft was made from. */
function printOf(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  let h = 0x811c9dc5;
  for (let i = 0; i < b.length; i++) h = Math.imul(h ^ b[i], 0x01000193);
  return `${(h >>> 0).toString(16)}-${b.length}`;
}

/**
 * Options beyond the callbacks:
 * - `readOnly`: show the document without editing (no ribbon, nothing editable; comments still show).
 * - `features`: turn parts off, e.g. `{ review: false }` (a whole tab), `{ comments: false, tracking: false }`
 *   (groups by their label on any tab; `comments: false` also hides comments in the margin), `{ autocorrect: false }`
 *   (AutoCorrect as you type), `{ newDocument: false, openFile: false }` (File > New and Open).
 * - `toolbar`: false to hide the ribbon entirely.
 * - `forms`: fill-in mode: no ribbon, only content controls (form fields) can be changed.
 * - `pdf`: the PDF exporter, `exportPdf` from 'open-docx-editor/pdf' (kept out of the main bundle): with it, File > Export
 *   downloads a PDF drawn from the page layout and `api.exportPdf()` returns one; without it, Export prints.
 *   `onSavePdf(bytes, fileName)` replaces that download, as `onSave` does the .docx's.
 * - `collab`: a room from 'open-docx-editor/collab' (collaboration(ydoc, { awareness, user })): the body is edited
 *   together. Mount after your Yjs provider has synced: the editor opens the room's document, or makes the room's
 *   first one; a document opened here becomes the room's, and one opened elsewhere opens here.
 * - `onMissingFonts(fonts)`: after each open, the fonts the document asks for that this computer lacks, as
 *   [{ font, substitute, standIn }] (substitute: the font drawn instead, or null for a generic one; standIn: one of
 *   the editor's own fonts sized to its widths), when there are any. Lines
 *   in a substitute may break differently from Word's; 'open-docx-editor/fonts.css' supplies metric-compatible ones.
 * - `ai`: your AI model, for Home > AI Writing: async ({ action, instruction, prompt, text, before, after, document,
 *   signal }) => text. `action` is 'rewrite', 'shorter', 'longer', 'formal', 'casual' or 'proofread' (rewrite `text`,
 *   the selection or the cursor's paragraph; `instruction` says how), 'summarize' (`text`, or `document` when nothing is
 *   selected) or 'ask' (answer `prompt` about them). The editor sends nothing anywhere itself. A rewrite goes in as a
 *   tracked change under `aiAuthor` ('AI'); an answer can be inserted below or put in place of the selection, tracked too.
 * - `drafts`: true, or { key, delayMs }, to keep unsaved changes in this browser (IndexedDB) a moment after each edit
 *   (delayMs, 3000) under `key` (the file's name if not given) and the file's fingerprint, cleared when the document is
 *   saved (`markSaved()`), forgotten after two weeks. Opening the same file again with a draft left
 *   offers it back: a notice above the editor with Restore and Discard, or `onDraft(draft, { restore, discard })` if
 *   given (draft: { name, savedAt, bytes }). Off by default: it keeps the document's content on this computer.
 * - `editors`: this user's names (besides `author`) as Word names editors in a protected document's editable parts
 *   (an e-mail address, DOMAIN\name): those parts stay editable for them. `features.protection: false` ignores
 *   Restrict Editing (and Mark as Final, Always open read-only) for a host that enforces its own rights.
 * Change them later with `api.setReadOnly(on)` and `api.setFeatures({...})`.
 */
export function mountDocxEditor(container, { onChange, onSave, onOpened, onError, fileName = 'document.docx', author = 'Author', readOnly = false, features = {}, toolbar: showToolbar = true, forms = false, pdf = null, onSavePdf, collab = null, onMissingFonts, drafts = false, onDraft, editors = [], mergeData = null, onMerge, ai = null, aiAuthor = 'AI' } = {}) {
  const bar = document.createElement('div');
  const host = document.createElement('div');
  container.replaceChildren(bar, host);
  let view = null;
  let name = fileName;
  const urls = [];
  const imageUrl = (bytes, mime) => {
    const u = URL.createObjectURL(new Blob([bytes], { type: mime }));
    urls.push(u);
    return u;
  };
  // Mail merge's recipients (mergeData: records, or { records, fields }), kept across documents.
  const merge = { records: null, fields: null };
  const takeMerge = (data, fields) => {
    const records = Array.isArray(data) ? data : data?.records;
    merge.records = Array.isArray(records) ? records : null;
    merge.fields = fields || (Array.isArray(data) ? null : data?.fields) || null;
  };
  takeMerge(mergeData);
  const api = {
    get view() { return view; },
    /**
     * Mail merge's recipients: `records` (objects, field name -> value), and the field names Insert Merge Field offers
     * (`fields`, else the records' keys). The Mailings tab previews and merges them; null clears them.
     */
    setMergeData(records, fields = null) {
      takeMerge(records, fields);
      if (view) mm.setRecipients(view, merge.records || [], merge.fields);
    },
    /** The document's merge field names. */
    get mergeFields() { return view ? mergeFieldNames(view.odx.session) : []; },
    /** Preview Results: each recipient's values in place of the fields (on screen), at record `index`; false: off. */
    previewMerge(on = true, index) { if (view) mm.setPreview(view.odx.main || view, on, index); },
    /**
     * Finish & Merge: the document merged with the recipients (or `records`): { bytes } (one document, each record
     * on a new page) or, with { combine: false }, { files }; with `fields` and `missing` (names no record has).
     */
    mailMerge({ records = merge.records, combine = true } = {}) {
      if (!records?.length) throw new Error('mailMerge: no recipients (api.setMergeData)');
      return mergeDocx(api.save(), records, { combine });
    },
    get fileName() { return name; },
    load(session) {
      if (view) view.odx.destroy();
      // (The recipients stay with the editor from one document to the next, as a host gave them: api.setMergeData.)
      if (merge.records) session.merge = { records: merge.records, fields: merge.fields, index: 0, preview: false };
      view = make(host, session, {
        author,
        collab,
        onChange: (v, tr) => {
          toolbar.update(v);
          if (!tr.docChanged) return;
          api.dirty = true;
          keepDraft();
          onChange && onChange(v);
        },
      });
      // A document Word protects for filling in forms opens in forms mode, unless features.autoForms is false.
      const locked = features.autoForms !== false && session.formsProtected;
      if (locked !== autoForms) { forms = locked; autoForms = locked; } // back to normal editing after the protected form
      softLock = softLockOf(session);
      applyMode();
      // The zoom carries over from the document before (a percentage, or Page Width / One Page kept fitting).
      view.odx.onZoom = (z) => { zoomWas = z.mode || z.value; toolbar.update(view); };
      if (zoomWas !== 1) view.odx.zoom.set(zoomWas, true);
      // (Edits in a header, footer, note or text box being edited count as changes too.)
      view.odx.notify = (v) => { toolbar.update(v); if (v !== view) { api.dirty = true; keepDraft(); onChange && onChange(v); } };
      api.dirty = false;
      toolbar.setDocument(session);
      toolbar.update(view);
      view.focus();
    },
    /** Read-only: the document is shown but can't be changed (handy for displaying filled-in forms). */
    get readOnly() { return readOnly; },
    setReadOnly(on) { readOnly = !!on; applyMode(); },
    /** Forms mode: only content controls can be filled in; the rest of the document is locked. */
    get forms() { return forms; },
    /** The open document is protected for filling in forms (it opens in forms mode). */
    get formsProtected() { return !!view?.odx.session.formsProtected; },
    setForms(on) { forms = !!on; applyMode(); },
    /** Every content control: [{ tag, alias, kind, value, checked?, items? }] in document order. */
    getFields() { return listFields(view.state.doc).map(({ tag, alias, kind, value, checked, items, block }) => ({ tag, alias, kind, value, checked, items, block })); },
    /** Field values as { name: value } (name = the control's Tag, else its Title). Checkboxes give true/false. */
    getFieldValues() { return fieldValues(view.state.doc); },
    /** Fill fields by Tag or Title: { CustomerName: 'Acme', Agree: true }. Returns names that matched no field. */
    setFieldValues(values) { return setFV(view, values); },
    /** Turn features on/off: ribbon tabs and groups, AutoCorrect, File > New and Open (see mountDocxEditor's `features`). */
    setFeatures(f) { features = { ...features, ...f }; toolbar.setFeatures(f); applyMode(); },
    /** True once the document has been edited since it was opened or last saved with markSaved(). */
    dirty: false,
    markSaved() {
      api.dirty = false;
      if (drafts) { clearTimeout(draftTimer); clearDraft(draftKey()); }
    },
    /** The draft of unsaved changes kept for the open document ({ key, name, savedAt, bytes }), or null (see `drafts`). */
    draft() { return drafts ? loadDraft(draftKey()) : Promise.resolve(null); },
    /** Remove the editor and toolbar (e.g. when a Blazor component is disposed). */
    destroy() {
      clearTimeout(draftTimer);
      stopCollab();
      if (view) view.odx.destroy();
      view = null;
      urls.splice(0).forEach((u) => URL.revokeObjectURL(u));
      container.replaceChildren();
    },
    /** The fonts the open document asks for that this computer lacks: [{ font, substitute }] (see onMissingFonts). */
    async missingFonts() { await fontsSettled(); return view ? missingFonts(view.dom) : []; },
    /** Timings of the last open, in ms: read (file to bytes), parse (.docx to model), editor (first render), layout (first pagination). */
    timings: null,
    /** Open a document (with `collab`, it becomes the room's unless `share` is false). */
    async open(src, fname, { share = true, reopen = false } = {}) {
      const t0 = performance.now();
      keepDraftNow(); // (the document open now keeps its unsaved changes under its own key)
      // A .docx is a zip: check the first bytes before reading a large file whole.
      if (src?.slice && !(src instanceof Uint8Array) && !(src instanceof ArrayBuffer)) {
        const head = new Uint8Array(await src.slice(0, 2).arrayBuffer());
        if (head[0] !== 0x50 || head[1] !== 0x4b) throw new Error(`${src.name || 'This file'} is not a Word document (.docx files are zip packages).`);
      }
      const bytes = src instanceof Uint8Array || src instanceof ArrayBuffer ? src : await src.arrayBuffer();
      const t1 = performance.now();
      // (Set once it has opened: a file that fails to open, or edits while it loads, leave the open document's.)
      const newName = fname || src.name || name;
      // Reopened (a page setup change, a restored draft): still the same file, with its unsaved changes.
      const print = reopen ? openedPrint : printOf(bytes);
      const wasDirty = api.dirty, wasLock = softLock;
      urls.splice(0).forEach((u) => URL.revokeObjectURL(u));
      const session = new Session(bytes, { imageUrl });
      await loadEmbeddedFonts(session.pkg); // laid out in the document's own fonts, as Word does
      const t2 = performance.now();
      if (collab && share) { dropView(); collab.publish(bytes, session.doc); } // (the old view mustn't follow the new body)
      api.load(session);
      name = newName;
      openedPrint = print;
      if (reopen) { api.dirty = wasDirty; softLock = wasLock; applyProtection(); }
      const t3 = performance.now();
      // pagination runs on the next animation frame
      await new Promise((r) => requestAnimationFrame(() => setTimeout(r)));
      const t4 = performance.now();
      api.timings = { read: t1 - t0, parse: t2 - t1, editor: t3 - t2, layout: t4 - t3, total: t4 - t0 };
      if (onOpened) onOpened(api.timings, name);
      if (onMissingFonts) api.missingFonts().then((m) => m.length && onMissingFonts(m), () => {});
      if (drafts && !reopen) offerDraft();
      return api.timings;
    },
    newDocument({ share = true } = {}) {
      keepDraftNow();
      name = 'Document.docx';
      openedPrint = 'new';
      const session = Session.blank({ imageUrl });
      if (collab && share) { dropView(); collab.publish(session.save(session.doc), session.doc); }
      api.load(session);
    },
    /**
     * Word's Update Fields: page numbers, cross-references (PAGEREF, REF, STYLEREF) and the table of contents in the
     * body, from the pages as laid out now. One undoable step. Returns how many were updated. (Headers and footers
     * work theirs out on every page by themselves.)
     */
    updateFields() { return view ? updateFields(view, { all: true }) : 0; },
    /**
     * An agent's view of the open document ('open-docx-editor/agent''s AgentDocument): read it as paragraphs, Markdown
     * or an outline, and edit it in plain terms, each edit a tracked change under `author` ('AI') that the person at
     * the editor sees, accepts or rejects.
     */
    agent({ author = 'AI' } = {}) { if (!view) throw new Error('No document is open.'); return editorAgent(view, { author }); },
    /**
     * The open document's Restrict Editing: { edit ('readOnly' | 'comments' | 'trackedChanges' | 'forms' | null),
     * enforced, formatting, password (one is set), readOnlyRecommended, passwordToModify, final, readOnlyNow (the
     * editor holds it read-only: Mark as Final or Always open read-only, until Edit Anyway) }.
     */
    get protection() {
      if (!view) return null;
      const p = view.odx.session.protection;
      return { edit: p.edit, enforced: p.enforced, formatting: p.formatting, password: p.password, readOnlyRecommended: !!p.writeProtection?.recommended, passwordToModify: !!p.writeProtection?.password, final: p.final, readOnlyNow: !!softLock };
    },
    /**
     * Word's Restrict Editing > Start Enforcing Protection: `edit` 'readOnly' (no changes), 'comments', 'trackedChanges'
     * or 'forms', or null; `formatting`: only the styles not locked (w:locked in styles.xml) and no direct formatting;
     * `password` (optional) to stop it. Saved with the document; Word enforces it too.
     */
    async protect({ edit = null, formatting = false, password = '' } = {}) {
      if (!view) throw new Error('No document is open.');
      await view.odx.session.protect({ edit, formatting, password, enforced: true });
      if (edit === 'forms' && features.autoForms !== false) forms = autoForms = true;
      changedProtection();
    },
    /** Stop Protection: resolves to false (still protected) when the document's password isn't `password`. */
    async unprotect(password = '') {
      if (!view) return true;
      if (!(await view.odx.session.unprotect(password))) return false;
      if (autoForms) forms = autoForms = false;
      changedProtection();
      return true;
    },
    /**
     * Edit Anyway for a document marked final or opened read-only as its author recommends; with a password to modify,
     * `password` must be it. Resolves to whether it can be edited now.
     */
    async editAnyway(password = '') {
      if (!view || !softLock) return true;
      if (softLock === 'password' && !(await view.odx.session.checkWritePassword(password))) return false;
      if (softLock === 'final') { view.odx.session.clearFinal(); api.dirty = true; }
      softLock = null;
      applyProtection();
      return true;
    },
    save() {
      view.odx.flushStories?.(); // (with the text of a header, footer or note still being edited)
      return view.odx.session.save(view.state.doc);
    },
    /** The zoom the pages are shown at (1 is 100%); the layout, printing and PDF export are always at 100%. */
    get zoom() { return view?.odx.zoom.value ?? zoomWas; },
    /** Show the pages at `zoom` (0.1 to 5), or 'pageWidth' / 'onePage' to fit the window as it resizes. */
    setZoom(zoom) { if (view) view.odx.zoom.set(zoom); else zoomWas = zoom; },
    /** The pages as a PDF (with the `pdf` option): Promise of its bytes. */
    async exportPdf() {
      if (!pdf) throw new Error('PDF export needs the pdf option: mountDocxEditor(el, { pdf: exportPdf }), exportPdf from open-docx-editor/pdf.');
      // (Measured at 100%, whatever the zoom on screen.)
      const z = view?.odx.zoom;
      z?.suspend();
      try { return await pdf(api, { title: name.replace(/\.docx$/i, ''), author }); } finally { z?.resume(); }
    },
    async downloadPdf(fname = name.replace(/\.docx$/i, '') + '.pdf') {
      const bytes = await api.exportPdf();
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }));
      a.download = fname;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 1000);
      return bytes;
    },
    download(fname = name) {
      const blob = new Blob([api.save()], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = fname;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    },
  };
  // Drafts (src/editor/drafts.js): unsaved changes kept in this browser, offered back on opening.
  let draftTimer = null;
  let draftNote = null;
  // A draft belongs to the file it was made from: keyed by its name (or the host's key) and a fingerprint of the file as
  // opened, so another file of the same name, or a newer copy of this one, is never offered it.
  let openedPrint = '';
  const draftKey = () => `${(typeof drafts === 'object' && drafts.key) || name}#${openedPrint}`;
  if (drafts) pruneDrafts(DRAFT_DAYS * 864e5);
  // (Not a new document's: with no file to open again, it could never be offered back.)
  const keeping = () => drafts && !readOnly && openedPrint && openedPrint !== 'new';
  const keepDraft = () => {
    if (!keeping()) return;
    clearTimeout(draftTimer);
    // (For the document open now, under its key: if another has opened by then, nothing is kept.)
    const key = draftKey(), v = view;
    draftTimer = setTimeout(() => { if (view === v && api.dirty) saveDraft(key, api.save(), name); }, (typeof drafts === 'object' && drafts.delayMs) || 3000);
  };
  /** Keeps the open document's unsaved changes at once (before another document takes its place). */
  const keepDraftNow = () => {
    clearTimeout(draftTimer);
    if (keeping() && view && api.dirty) saveDraft(draftKey(), api.save(), name);
  };
  async function offerDraft() {
    const key = draftKey();
    const draft = await loadDraft(key);
    if (draftKey() !== key) return; // (another document opened meanwhile)
    draftNote?.remove();
    if (!draft) return;
    if (!(Date.now() - draft.savedAt < DRAFT_DAYS * 864e5)) { await clearDraft(key); return; } // (kept two weeks)
    // (Each acts on the document the draft was offered for, and only while it's the one open.)
    const restore = async () => {
      draftNote?.remove();
      if (draftKey() !== key) return;
      await api.open(draft.bytes, draft.name || name, { reopen: true });
      api.dirty = true; // (still unsaved: the draft stays until it's saved)
    };
    const discard = async () => { draftNote?.remove(); await clearDraft(key); };
    if (onDraft) return onDraft(draft, { restore, discard });
    // A notice above the ribbon; it doesn't take the focus from the document.
    draftNote = document.createElement('div');
    draftNote.className = 'odx-draftnote';
    draftNote.setAttribute('role', 'status');
    setTimeout(() => announce(`${text.textContent} Restore or Discard are above the ribbon.`), 600);
    const text = document.createElement('span');
    text.textContent = `Unsaved changes to ${draft.name || name} from ${new Date(draft.savedAt).toLocaleString()} were kept in this browser.`;
    const button = (label, fn) => { const b = document.createElement('button'); b.type = 'button'; b.textContent = label; b.addEventListener('click', () => fn().catch(fail)); return b; };
    draftNote.append(text, button('Restore', restore), button('Discard', discard));
    container.prepend(draftNote);
  }
  let autoForms = false;
  let zoomWas = 1;
  // Mark as Final, Always open read-only, a password to modify: the document stays read-only until Edit Anyway.
  let softLock = null;
  const softLockOf = (session) => {
    if (features.protection === false) return null;
    const p = session.protection;
    return p.final ? 'final' : p.writeProtection?.password ? 'password' : p.writeProtection?.recommended ? 'recommended' : null;
  };
  let protectNote = null;
  let saidNote = '';
  /** The editor enforces the document's protection as it is now; a notice above the ribbon says what it allows. */
  function applyProtection() {
    if (!view) return;
    const p = view.odx.session.protection;
    const on = features.protection !== false && p.enforced;
    const edit = softLock ? 'readOnly' : on && p.edit !== 'forms' ? p.edit : null;
    const formatting = on && p.formatting;
    setProtection(view, edit || formatting ? { edit, formatting, locked: p.locked, editors: [author, ...editors] } : null);
    protectNote?.remove();
    protectNote = null;
    const say = readOnly ? null
      : softLock === 'final' ? 'An author has marked this document as final to discourage editing.'
        : softLock === 'recommended' ? 'The author would like you to open this document as read-only, unless you need to make changes.'
          : softLock === 'password' ? `${name} is reserved: enter the password to modify it, or read it as it is.`
            : !on ? null
              : NOTE[p.edit] || (formatting ? 'Restricted editing: only some styles can be used.' : null);
    if (!say) { saidNote = ''; return; }
    // (Said once as it appears: a status region added with its text isn't read by screen readers on its own.)
    if (say !== saidNote) { saidNote = say; setTimeout(() => announce(softLock === 'password' ? `${say} The password box is above the ribbon.` : softLock ? `${say} Edit Anyway is above the ribbon.` : say), 600); }
    protectNote = document.createElement('div');
    protectNote.className = 'odx-protectnote';
    protectNote.setAttribute('role', 'status');
    const text = document.createElement('span');
    text.textContent = say;
    protectNote.append(text);
    if (softLock) {
      const pw = softLock === 'password' ? Object.assign(document.createElement('input'), { type: 'password', autocomplete: 'off' }) : null;
      if (pw) { pw.setAttribute('aria-label', 'Password to modify'); protectNote.append(pw); }
      const b = document.createElement('button');
      b.type = 'button';
      b.textContent = pw ? 'Modify' : 'Edit Anyway';
      const go = () => api.editAnyway(pw ? pw.value : '').then((ok) => { if (!ok) { text.textContent = 'That password isn’t right. Try again, or read the document as it is.'; pw?.focus(); } else view?.focus(); }).catch(fail);
      b.addEventListener('click', go);
      pw?.addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); });
      protectNote.append(b);
    }
    container.prepend(protectNote);
  }
  const NOTE = {
    readOnly: 'Restricted editing: this document can only be read, apart from any highlighted parts.',
    comments: 'Restricted editing: only comments can be added, apart from any highlighted parts.',
    trackedChanges: 'Restricted editing: every change is tracked.',
    forms: 'Restricted editing: only the form fields can be filled in.',
  };
  function changedProtection() {
    api.dirty = true;
    keepDraft();
    applyMode();
    toolbar.setDocument(view.odx.session);
    toolbar.update(view);
    onChange && onChange(view);
  }
  const dropView = () => { if (view) view.odx.destroy(); view = null; };
  const fail = (e) => (onError ? onError(e) : console.error(e));
  const toolbar = makeToolbar(bar, () => view?.odx?.hfView || view, { onOpen: (f) => api.open(f).catch(fail), onNew: () => api.newDocument(), onReopen: (bytes) => api.open(bytes, name, { reopen: true }), onSave: () => (onSave ? onSave(api.save(), name) : api.download()), onMerge: (records) => {
    try {
      const { bytes } = mergeDocx(api.save(), records);
      const merged = name.replace(/\.docx$/i, '') + ' (merged).docx';
      if (onMerge) onMerge(bytes, merged);
      else {
        const a = document.createElement('a');
        a.href = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
        a.download = merged;
        a.click();
        setTimeout(() => URL.revokeObjectURL(a.href), 1000);
      }
      announce(`Merged ${records.length} record${records.length === 1 ? '' : 's'}`);
    } catch (e) { fail(e); }
  }, ai, aiAuthor, onExportPdf: pdf ? () => (onSavePdf ? api.exportPdf().then((b) => onSavePdf(b, name.replace(/\.docx$/i, '') + '.pdf')) : api.downloadPdf()).catch(fail) : null, protection: { get: () => api.protection, protect: (o) => api.protect(o).catch(fail), unprotect: (pw) => api.unprotect(pw) } });
  function applyMode() {
    bar.style.display = readOnly || forms || !showToolbar ? 'none' : '';
    container.classList.toggle('odx-readonly', readOnly);
    container.classList.toggle('odx-no-markup', readOnly || forms);
    container.classList.toggle('odx-no-comments', features.comments === false);
    if (!view) return;
    view.odx.readOnly = readOnly;
    view.odx.features = features;
    view.setProps({ editable: () => !readOnly });
    setFormsMode(view, forms && !readOnly);
    container.classList.toggle('odx-no-ffshade', !view.odx.session.shadeFormData);
    applyProtection();
    view.odx.pages.update();
  }
  toolbar.setFeatures(features);
  // Co-editing: the room's document, or a new one as the room's first; one opened elsewhere opens here too.
  let stopCollab = () => {};
  if (collab) {
    stopCollab = collab.onDocument((bytes) => api.open(bytes, undefined, { share: false }).catch(fail));
    // (A blank page meanwhile, not shared: the room's own document may still be on its way, collab.ready.)
    api.newDocument({ share: false });
    let stopped = false;
    const stop0 = stopCollab;
    stopCollab = () => { stopped = true; stop0(); };
    Promise.resolve(collab.ready).catch(() => {}).then(() => {
      if (stopped) return;
      if (collab.bytes) api.open(collab.bytes, name, { share: false }).catch(fail);
      else api.newDocument();
    });
  } else api.newDocument();
  return api;
}
