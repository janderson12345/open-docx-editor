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

/**
 * Options beyond the callbacks:
 * - `readOnly`: show the document without editing (no ribbon, nothing editable; comments still show).
 * - `features`: turn parts off, e.g. `{ review: false }` (a whole tab), `{ comments: false, tracking: false }`
 *   (groups by their label on any tab; `comments: false` also hides comments in the margin).
 * - `toolbar`: false to hide the ribbon entirely.
 * - `forms`: fill-in mode: no ribbon, only content controls (form fields) can be changed.
 * - `pdf`: the PDF exporter, `exportPdf` from 'open-docx-editor/pdf' (kept out of the main bundle): with it, File > Export
 *   downloads a PDF drawn from the page layout and `api.exportPdf()` returns one; without it, Export prints.
 *   `onSavePdf(bytes, fileName)` replaces that download, as `onSave` does the .docx's.
 * - `collab`: a room from 'open-docx-editor/collab' (collaboration(ydoc, { awareness, user })): the body is edited
 *   together. Mount after your Yjs provider has synced: the editor opens the room's document, or makes the room's
 *   first one; a document opened here becomes the room's, and one opened elsewhere opens here.
 * Change them later with `api.setReadOnly(on)` and `api.setFeatures({...})`.
 */
export function mountDocxEditor(container, { onChange, onSave, onOpened, onError, fileName = 'document.docx', author = 'Author', readOnly = false, features = {}, toolbar: showToolbar = true, forms = false, pdf = null, onSavePdf, collab = null } = {}) {
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
  const api = {
    get view() { return view; },
    get fileName() { return name; },
    load(session) {
      if (view) view.odx.destroy();
      view = make(host, session, {
        author,
        collab,
        onChange: (v, tr) => {
          toolbar.update(v);
          if (!tr.docChanged) return;
          api.dirty = true;
          onChange && onChange(v);
        },
      });
      // A document Word protects for filling in forms opens in forms mode, unless features.autoForms is false.
      const locked = features.autoForms !== false && session.formsProtected;
      if (locked !== autoForms) { forms = locked; autoForms = locked; } // back to normal editing after the protected form
      applyMode();
      view.odx.notify = (v) => { toolbar.update(v); if (v !== view) { api.dirty = true; } };
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
    /** Turn ribbon features on/off (see mountDocxEditor's `features`). */
    setFeatures(f) { features = { ...features, ...f }; toolbar.setFeatures(f); applyMode(); },
    /** True once the document has been edited since it was opened or last saved with markSaved(). */
    dirty: false,
    markSaved() { api.dirty = false; },
    /** Remove the editor and toolbar (e.g. when a Blazor component is disposed). */
    destroy() {
      stopCollab();
      if (view) view.odx.destroy();
      view = null;
      urls.splice(0).forEach((u) => URL.revokeObjectURL(u));
      container.replaceChildren();
    },
    /** Timings of the last open, in ms: read (file to bytes), parse (.docx to model), editor (first render), layout (first pagination). */
    timings: null,
    /** Open a document (with `collab`, it becomes the room's unless `share` is false). */
    async open(src, fname, { share = true } = {}) {
      const t0 = performance.now();
      // A .docx is a zip: check the first bytes before reading a large file whole.
      if (src?.slice && !(src instanceof Uint8Array) && !(src instanceof ArrayBuffer)) {
        const head = new Uint8Array(await src.slice(0, 2).arrayBuffer());
        if (head[0] !== 0x50 || head[1] !== 0x4b) throw new Error(`${src.name || 'This file'} is not a Word document (.docx files are zip packages).`);
      }
      const bytes = src instanceof Uint8Array || src instanceof ArrayBuffer ? src : await src.arrayBuffer();
      const t1 = performance.now();
      name = fname || src.name || name;
      urls.splice(0).forEach((u) => URL.revokeObjectURL(u));
      const session = new Session(bytes, { imageUrl });
      await loadEmbeddedFonts(session.pkg); // laid out in the document's own fonts, as Word does
      const t2 = performance.now();
      if (collab && share) { dropView(); collab.publish(bytes, session.doc); } // (the old view mustn't follow the new body)
      api.load(session);
      const t3 = performance.now();
      // pagination runs on the next animation frame
      await new Promise((r) => requestAnimationFrame(() => setTimeout(r)));
      const t4 = performance.now();
      api.timings = { read: t1 - t0, parse: t2 - t1, editor: t3 - t2, layout: t4 - t3, total: t4 - t0 };
      if (onOpened) onOpened(api.timings, name);
      return api.timings;
    },
    newDocument({ share = true } = {}) {
      name = 'Document.docx';
      const session = Session.blank({ imageUrl });
      if (collab && share) { dropView(); collab.publish(session.save(session.doc), session.doc); }
      api.load(session);
    },
    save() {
      return view.odx.session.save(view.state.doc);
    },
    /** The pages as a PDF (with the `pdf` option): Promise of its bytes. */
    exportPdf() {
      if (!pdf) return Promise.reject(new Error('PDF export needs the pdf option: mountDocxEditor(el, { pdf: exportPdf }), exportPdf from open-docx-editor/pdf.'));
      return pdf(api, { title: name.replace(/\.docx$/i, ''), author });
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
  let autoForms = false;
  const dropView = () => { if (view) view.odx.destroy(); view = null; };
  const fail = (e) => (onError ? onError(e) : console.error(e));
  const toolbar = makeToolbar(bar, () => view?.odx?.hfView || view, { onOpen: (f) => api.open(f).catch(fail), onNew: () => api.newDocument(), onReopen: (bytes) => api.open(bytes, name), onSave: () => (onSave ? onSave(api.save(), name) : api.download()), onExportPdf: pdf ? () => (onSavePdf ? api.exportPdf().then((b) => onSavePdf(b, name.replace(/\.docx$/i, '') + '.pdf')) : api.downloadPdf()).catch(fail) : null });
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
