// A Word-style ribbon (plain DOM, no framework): File, Home, Insert, Layout,
// References, Review and View tabs, laid out like Word for the web. Its state is
// refreshed on every transaction through update().
import { checkAccessibility } from './a11y.js';
import { createRuler } from './ruler.js';
import * as bm from './bookmarks.js';
import { LIST_LIBRARY } from '../ooxml/template.js';
import * as notes from './notes.js';
import * as toc from './toc.js';
import { updateFieldsCommand } from './refs.js';
import * as hf from './hfedit.js';
import * as tb from './txbxedit.js';
import { toggleMarks, marksOn } from './marks.js';
import * as cm from './comments.js';
import { undo, redo } from 'prosemirror-history';
import { addColumnAfter, addColumnBefore, deleteRow, deleteColumn, mergeCells, splitCell, deleteTable, isInTable } from 'prosemirror-tables';
import { TextSelection, AllSelection } from 'prosemirror-state';
import { schema } from './schema.js';
import * as cmd from './commands.js';
import { setTracking, tracking, resolveChange, resolveAll, gotoChange } from './track.js';
import { createFindBar } from './findbar.js';
import { protectionOf } from './protect.js';
import { ZOOM_STEPS } from './zoom.js';
import { createCommandSearch } from './cmdsearch.js';
import { announce } from './live.js';
import { createContextMenu } from './ctxmenu.js';
import { el, iconize, slug, ribbonKit } from './ribbonkit.js';
import * as sp from '../ooxml/sectpr.js';
import { WRAPS } from '../ooxml/wrap.js';
import { setWrapText, currentWrap } from './wraptext.js';
import { sectionsTr, pageSetupPanel } from './pagesetup.js';

const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const FONTS = ['Aptos', 'Arial', 'Calibri', 'Cambria', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana'];
const PAGE_COLORS = ['FFFFFF', 'F2F2F2', 'FFF2CC', 'FBE5D6', 'E2EFDA', 'DDEBF7', 'EDE2F6', 'FCE4EC', 'FFFDE7', 'E0F7FA', 'F1F8E9', 'ECEFF1', '1E1E1E', '262626', '0D2137', '2B1B3D'];
const SHADES = ['FFFFFF', 'F2F2F2', 'D9D9D9', 'BFBFBF', 'FFF2CC', 'FFE699', 'FBE5D6', 'F8CBAD', 'E2EFDA', 'C6E0B4', 'DDEBF7', 'BDD7EE', 'D9E1F2', 'B4C6E7', 'EDE2F6', 'D5C4E8'];
const HIGHLIGHTS = [['yellow', '#ffff00'], ['green', '#00ff00'], ['cyan', '#00ffff'], ['magenta', '#ff00ff'], ['blue', '#0000ff'], ['red', '#ff0000'], ['darkBlue', '#000080'], ['darkCyan', '#008080'], ['darkGreen', '#008000'], ['darkMagenta', '#800080'], ['darkRed', '#800000'], ['darkYellow', '#808000'], ['darkGray', '#808080'], ['lightGray', '#c0c0c0'], ['black', '#000000']];
const SYMBOLS = '©®™§¶†‡•…–—‘’“”«»€£¥¢°±×÷≠≤≥≈∞√∑πµΩαβγδ←→↑↓✓✗☐☑★☆♥';
const LINE_SPACING = [1, 1.15, 1.5, 2, 2.5, 3];
// Page sizes and margins in twips, as Word's Layout menus offer them.
const PAGE_SIZES = [['Letter', '8.5" × 11"', 12240, 15840], ['Legal', '8.5" × 14"', 12240, 20160], ['A4', '8.27" × 11.69"', 11906, 16838], ['A5', '5.83" × 8.27"', 8391, 11906], ['Executive', '7.25" × 10.5"', 10440, 15120]];
const MARGINS = [['Normal', 1440, 1440, 1440, 1440], ['Narrow', 720, 720, 720, 720], ['Moderate', 1440, 1080, 1440, 1080], ['Wide', 1440, 2880, 1440, 2880]];

/**
 * @param container element the ribbon is drawn into
 * @param getView () => the current EditorView
 * @param opts { onOpen(file), onNew(), onSave(), onReopen(bytes), protection } — onReopen reloads the
 *   document after a page setup change (size, margins, orientation); protection: { get(), protect(opts),
 *   unprotect(password) } for Review > Restrict Editing (the editor API's)
 */
export function createToolbar(container, getView, { onOpen, onNew, onSave, onReopen, onExportPdf, protection = null } = {}) {
  const kit = ribbonKit(container, getView);
  const { run, withView, btn, grp, col, line, pop, openMenu, closeMenu, item, menu, pick } = kit;
  const ruler = createRuler(() => { const v = getView(); return v; });
  const setPageColor = (hex) => (state, dispatch, v) => {
    const main = v.odx.main || v;
    main.odx.session.setPageColor(hex);
    main.dom.parentElement.style.setProperty('--odx-page-color', hex ? `#${hex}` : null);
    if (!hex) main.dom.parentElement.style.removeProperty('--odx-page-color');
    main.dispatch(main.state.tr.setDocAttribute('sectPr', main.state.doc.attrs.sectPr)); // marks the document changed
    return true;
  };
  // View > Zoom: the pages drawn larger or smaller (zoom.js), the percentage in a list as Word's zoom box.
  const zoomOfView = (v) => (v.odx.main || v).odx.zoom;
  const zoomSel = pick('Zoom', 'odx-zoomsel', ZOOM_STEPS.map((z) => el('option', { value: String(z) }, `${Math.round(z * 100)}%`)), (value, v) => zoomOfView(v).set(Number(value)));

  // Review > Restrict Editing, as Word's pane: what may be changed and a password; or, once enforced, Stop Protection.
  const EDITS = [['readOnly', 'No changes (read only)'], ['comments', 'Comments'], ['trackedChanges', 'Tracked changes'], ['forms', 'Filling in forms']];
  const typed = (attrs) => {
    const i = el('input', attrs);
    i.addEventListener('keydown', (e) => { if (e.key !== 'Escape' && e.key !== 'Tab') e.stopPropagation(); });
    return i;
  };
  function restrictPanel() {
    const p = protection.get();
    if (!p) return [];
    const msg = el('div', { class: 'odx-cap', role: 'status' });
    const box = (...kids) => [el('div', { class: 'odx-protect' }, ...kids, msg)];
    if (p.enforced && (p.edit || p.formatting)) {
      const what = [p.edit && `Editing: ${EDITS.find((x) => x[0] === p.edit)[1].toLowerCase()}`, p.formatting && 'Formatting: only the allowed styles'].filter(Boolean).join('. ');
      const pw = p.password ? typed({ type: 'password', class: 'odx-num', 'aria-label': 'Password', placeholder: 'Password', autocomplete: 'off' }) : null;
      const stop = async (e) => {
        e.preventDefault();
        if (await protection.unprotect(pw ? pw.value : '')) { closeMenu(); announce('Protection stopped'); getView()?.focus(); } else { msg.textContent = 'The password is incorrect.'; pw?.focus(); }
      };
      pw?.addEventListener('keydown', (e) => { if (e.key === 'Enter') stop(e); });
      setTimeout(() => pw?.focus(), 0);
      return box(el('b', {}, 'This document is protected.'), el('div', {}, `${what}.`), pw, el('button', { type: 'button', class: 'odx-item', onmousedown: stop }, 'Stop Protection'));
    }
    const edit = el('select', { 'aria-label': 'Editing restrictions: allow only this type of editing' }, el('option', { value: '' }, 'No editing restrictions'), ...EDITS.map(([v, t]) => el('option', { value: v }, t)));
    edit.value = p.edit || '';
    const fmt = typed({ type: 'checkbox' });
    fmt.checked = !!p.formatting;
    const pw1 = typed({ type: 'password', class: 'odx-num', 'aria-label': 'New password (optional)', placeholder: 'New password (optional)', autocomplete: 'new-password' });
    const pw2 = typed({ type: 'password', class: 'odx-num', 'aria-label': 'Reenter password to confirm', placeholder: 'Reenter password to confirm', autocomplete: 'new-password' });
    const start = async (e) => {
      e.preventDefault();
      if (!edit.value && !fmt.checked) { msg.textContent = 'Choose a restriction first.'; return edit.focus(); }
      if (pw1.value !== pw2.value) { msg.textContent = 'The passwords don’t match.'; return pw2.focus(); }
      msg.textContent = 'Protecting…';
      await protection.protect({ edit: edit.value || null, formatting: fmt.checked, password: pw1.value });
      closeMenu();
      announce('Protection started');
      getView()?.focus();
    };
    return box(el('b', {}, 'Restrict editing'), el('label', { class: 'odx-cap' }, 'Allow only this type of editing:'), edit,
      el('label', { class: 'odx-check' }, fmt, ' Limit formatting to the allowed styles'),
      el('div', { class: 'odx-cap' }, 'A password (optional) is needed to stop the protection:'), pw1, pw2,
      el('button', { type: 'button', class: 'odx-item', onmousedown: start }, 'Yes, Start Enforcing Protection'));
  }
  // Styles gallery: tiles drawn in each style's own look (the document's style CSS applies inside).
  const styleTiles = el('div', { class: 'odx-stiles' });
  let styleList = [];
  let docScope = '';
  const tile = (st) => el('button', { type: 'button', class: 'odx-stile', title: st.name, 'data-id': st.id || '', onmousedown: run((s, d) => cmd.setStyle(st.id || null)(s, d)) },
    el('p', st.id ? { 'data-style': st.id } : {}, st.name));
  const styleGal = el('div', { class: 'odx-stylegal' }, styleTiles,
    menu('', 'More styles', () => [el('div', { class: `odx-stylepop ${docScope}` }, ...styleList.map(tile))]));
  const styleSel = pick('Styles', 'odx-style', [], (val, v) => cmd.setStyle(val || null)(v.state, v.dispatch));
  const fontSel = pick('Font', 'odx-font', [el('option', { value: '' }, ''), ...FONTS.map((f) => el('option', { value: f, style: `font-family:'${f}'` }, f))],
    (val, v) => cmd.setMark('font', val ? { val } : null)(v.state, v.dispatch));
  const sizeSel = pick('Font size', 'odx-size', [el('option', { value: '' }, ''), ...SIZES.map((s) => el('option', { value: s }, String(s)))],
    (val, v) => cmd.setMark('size', val ? { val: Math.round(Number(val) * 2) } : null)(v.state, v.dispatch));
  const color = el('input', { type: 'color', title: 'Font color', 'aria-label': 'Font color', value: '#c00000', class: 'odx-color', onchange: (e) => { const v = getView(); cmd.setMark('color', { val: e.target.value.slice(1).toUpperCase() })(v.state, v.dispatch); v.focus(); } });
  /** A number field (shown in `unit`, stored in twips) for Layout's indent and spacing boxes. */
  const twipField = (title, unit, onSet) => {
    const f = el('input', { type: 'number', title, 'aria-label': title, class: 'odx-num', step: unit === 'in' ? '0.1' : '6', min: '0', onchange: (e) => {
      const v = getView();
      const n = Number(e.target.value);
      if (!Number.isFinite(n)) return;
      onSet(Math.round(unit === 'in' ? n * 1440 : n * 20))(v.state, v.dispatch);
      v.focus();
    } });
    return f;
  };
  const indLeft = twipField('Indent left (inches)', 'in', (t) => cmd.setIndent({ left: t }));
  const indRight = twipField('Indent right (inches)', 'in', (t) => cmd.setIndent({ right: t }));
  const spBefore = twipField('Spacing before (pt)', 'pt', (t) => cmd.setSpacing({ before: t, beforeAutospacing: null }));
  const spAfter = twipField('Spacing after (pt)', 'pt', (t) => cmd.setSpacing({ after: t, afterAutospacing: null }));

  // --- Files ---
  const fileInput = el('input', { type: 'file', accept: '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document', hidden: true, onchange: (e) => e.target.files[0] && onOpen(e.target.files[0]) });
  const picInput = el('input', { type: 'file', accept: 'image/png,image/jpeg,image/gif,image/bmp', hidden: true, onchange: (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) insertPicture(f); } });
  async function insertPicture(file) {
    const v = getView();
    if (!v) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    const url = URL.createObjectURL(file);
    const img = await new Promise((res) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => res(null); i.src = url; });
    URL.revokeObjectURL(url);
    if (!img) return;
    // Fit the text width, as Word does for a large picture.
    const max = 624;
    const scale = Math.min(1, max / img.naturalWidth);
    const width = Math.round(img.naturalWidth * scale);
    const height = Math.round(img.naturalHeight * scale);
    let { xml, src } = v.odx.session.addImage(bytes, file.type || 'image/png', width, height);
    // In a shared document the picture's address goes to the others with it: a data URL, which any browser can show
    // (a blob URL is only this page's).
    if (v.odx.collab) src = await new Promise((res) => { const r = new FileReader(); r.onload = () => res(r.result); r.readAsDataURL(file); });
    v.dispatch(v.state.tr.replaceSelectionWith(schema.nodes.image.create({ xml, src, width, height })).scrollIntoView());
    v.focus();
  }

  // A new text box (2in by 1in, as Word draws one by default), opened for typing.
  function insertTextBox() {
    const v = getView();
    if (!v || v.odx?.isStory) return;
    const width = 192;
    const height = 96;
    const { xml, src } = v.odx.session.addTextBox(width, height);
    const pos = v.state.selection.from;
    v.dispatch(v.state.tr.replaceSelectionWith(schema.nodes.image.create({ xml, src, width, height, alt: 'Text Box' })).scrollIntoView());
    let at = null;
    v.state.doc.nodesBetween(pos, Math.min(v.state.doc.content.size, pos + 4), (n, p) => { if (at == null && n.type.name === 'image' && n.attrs.xml === xml) at = p; });
    if (at != null) setTimeout(() => tb.editTextBox(v, at), 50); // once the page shows it
  }

  // --- Find and Replace bar ---
  const { findBar, openFind } = createFindBar(getView, withView);

  // --- Link box ---
  const linkBox = el('input', { type: 'url', placeholder: 'https://…', class: 'odx-link-input', hidden: true, onkeydown: (e) => {
    if (e.key !== 'Enter' && e.key !== 'Escape') return;
    e.preventDefault();
    const v = getView();
    if (e.key === 'Enter') cmd.setLink(linkBox.value.trim() || null)(v.state, v.dispatch);
    linkBox.hidden = true;
    v.focus();
  } });
  const editLink = withView((v) => { linkBox.hidden = false; linkBox.value = cmd.markValue(v.state, 'link')?.href || ''; setTimeout(() => linkBox.focus()); return true; });

  // --- Format Painter: copy formatting at the caret, paint it on the next selection ---
  // Click: paint once. Double-click: keep painting until Esc or another click on the button.
  let painter = null;
  let sticky = false;
  let paintTimer = 0;
  const setPainter = (fmt, keep = false) => {
    painter = fmt;
    sticky = !!fmt && keep;
    b.painter.classList.toggle('on', !!painter);
    document.body.classList.toggle('odx-painting', !!painter);
  };
  const formatPainter = withView((v) => { setPainter(painter ? null : cmd.formatAt(v.state)); return true; });
  const stickyPainter = (e) => { e.preventDefault(); const v = getView(); if (v) setPainter(cmd.formatAt(v.state), true); };
  document.addEventListener('mouseup', (e) => {
    const v = getView();
    if (!painter || !v || !v.dom.contains(e.target)) return;
    clearTimeout(paintTimer);
    const fmt = painter;
    paintTimer = setTimeout(() => { // after the click (or double/triple click) has placed the selection
      cmd.applyFormat(fmt)(v.state, v.dispatch);
      if (!sticky) setPainter(null);
    }, 250);
  });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && painter) setPainter(null); });

  // --- Clipboard (cut/copy use the browser's own commands; paste needs clipboard permission) ---
  const clip = (what) => withView((v) => {
    v.focus();
    if (what !== 'paste') { document.execCommand(what); return true; }
    return pasteFrom(false)(v.state, v.dispatch, v); // keeps formatting, like Ctrl+V
  });

  const pasteFrom = (textOnly) => withView((v) => {
    const said = () => { v.odx.notify?.(v); };
    const failed = () => { flash('To paste here, use Ctrl+V (the browser blocked access to the clipboard).'); };
    const read = textOnly || !navigator.clipboard?.read
      ? navigator.clipboard?.readText?.().then((t) => v.pasteText(t))
      : navigator.clipboard.read().then(async (items) => {
        for (const it of items) {
          if (it.types.includes('text/html')) return v.pasteHTML(await (await it.getType('text/html')).text());
          if (it.types.includes('text/plain')) return v.pasteText(await (await it.getType('text/plain')).text());
        }
        return false;
      });
    (read || Promise.reject()).then(said, failed);
    return true;
  });
  const flash = (msg) => { const t = el('div', { class: 'odx-ctxnote', role: 'status' }, msg); container.append(t); setTimeout(() => t.remove(), 3500); };

  // --- Page setup: edits the final section's w:sectPr, then reloads so pages are re-laid out ---
  // (The section the cursor is in, or with `all` every one; then the document is laid out afresh.)
  const applyPageSetup = (v, fn, all = false) => {
    const main = v.odx?.main || v;
    const tr = sectionsTr(main.state, fn, all);
    if (!tr) return;
    main.dispatch(tr);
    if (onReopen) onReopen(main.odx.session.save(main.state.doc));
  };
  const pageSetup = (fn) => withView((v) => { applyPageSetup(v, fn); return true; });
  const setMargins = (t, r, bm, l) => pageSetup((x) => sp.setMargins(x, t, r, bm, l));
  const setOrientation = (landscape) => pageSetup((x) => sp.setOrientation(x, landscape));
  const setColumns = (num) => pageSetup((x) => sp.setColumns(x, num));
  const setSize = (w, h) => pageSetup((x) => sp.setPageSize(x, w, h));

  // --- Word count, shown in a small popup under the button ---
  const wordCount = (owner) => () => {
    const v = getView();
    const c = cmd.wordCount(v.state.doc);
    openMenu(owner(), () => [el('div', { class: 'odx-stats' },
      el('b', {}, 'Word Count'),
      ...[['Pages', v.odx.pages.pages], ['Words', c.words], ['Characters (no spaces)', c.charactersNoSpaces], ['Characters (with spaces)', c.characters], ['Paragraphs', c.paragraphs]]
        .map(([k, n]) => el('div', {}, el('span', {}, k), el('span', {}, String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ','))))) ]);
    return true;
  };

  // --- Navigation pane: the document's headings ---
  const navPane = el('div', { class: 'odx-nav', hidden: true });
  function fillNav() {
    const v = getView();
    if (!v || navPane.hidden) return;
    const s = v.odx.session;
    const items = [];
    v.state.doc.forEach((node, pos) => {
      if (node.type !== schema.nodes.paragraph || !node.textContent.trim()) return;
      const lvl = cmd.outlineLevel(s, node);
      if (lvl == null || lvl > 8) return;
      items.push(el('button', { type: 'button', class: 'odx-navitem', style: `padding-left:${8 + lvl * 14}px`, onmousedown: (e) => {
        e.preventDefault();
        v.dispatch(v.state.tr.setSelection(TextSelection.create(v.state.doc, pos + 1)));
        v.nodeDOM(pos)?.scrollIntoView({ block: 'start' });
        v.focus();
      } }, node.textContent.slice(0, 80)));
    });
    navPane.replaceChildren(el('div', { class: 'odx-navhead' }, 'Navigation'), ...(items.length ? items : [el('div', { class: 'odx-navnone' }, 'No headings. Use the Heading styles to list them here.')]));
  }
  const toggleNav = withView(() => { navPane.hidden = !navPane.hidden; b.nav.classList.toggle('on', !navPane.hidden); fillNav(); return true; });

  const setMarkupView = (mode) => withView((v) => {
    const page = v.dom.closest('.odx-page');
    page?.classList.toggle('odx-all-markup', mode === 'all');
    // (Original, as Word's: the document before the changes, without their markup.)
    page?.classList.toggle('odx-no-markup', mode === 'none' || mode === 'original');
    page?.classList.toggle('odx-original', mode === 'original');
    v.odx?.pages.update();
    markupSel.value = mode;
    return true;
  });
  const markupSel = pick('Markup view', 'odx-markupsel', [el('option', { value: 'simple' }, 'Simple Markup'), el('option', { value: 'all' }, 'All Markup'), el('option', { value: 'none' }, 'No Markup'), el('option', { value: 'original' }, 'Original')],
    (val, v) => setMarkupView(val)(v.state, v.dispatch, v));
  const trackSel = pick('Track Changes', 'odx-tracksel', [el('option', { value: 'off' }, 'Off'), el('option', { value: 'on' }, 'For Everyone')],
    (val, v) => setTracking(val === 'on')(v.state, v.dispatch));

  const noList = (s, d, v) => { const p = cmd.selectedParagraphs(s); if (d) { const tr = s.tr; for (const { node, pos } of p) tr.setNodeMarkup(pos, null, { ...node.attrs, numId: node.attrs.styleId && cmd.paragraphList(v.odx.session, { styleId: node.attrs.styleId }) ? '0' : null, ilvl: 0 }); d(tr); } return true; };
  const gallery = (title, kind, icon = '') => menu(icon, title, () => [
    el('div', { class: 'odx-gallery' }, ...LIST_LIBRARY[kind].map((def) => el('button', { type: 'button', class: 'odx-gal', title: def.label, onmousedown: run((s, d, v) => cmd.setList(v.odx.session, def.levels)(s, d)) }, ...(kind === 'multilevel' ? def.label.split(' ').map((t, i) => el('div', { style: `padding-left:${i * 8}px` }, t)) : [def.label])))),
    item('None', noList),
  ]);
  const listCmd = (kind) => (s, d, v) => cmd.toggleList(v.odx.session, kind)(s, d);
  // Increase/Decrease Indent: a list level in a list, else the left indent by 0.5".
  const indentCmd = (dir) => (s, d, v) => {
    const p = cmd.selectedParagraphs(s)[0];
    if (p && cmd.paragraphList(v.odx.session, p.node.attrs)) return cmd.changeListLevel(v.odx.session, dir)(s, d);
    return cmd.shiftIndent(dir * 720)(s, d);
  };
  const lineSpacing = (s) => {
    const p = cmd.selectedParagraphs(s)[0]?.node.attrs.view || {};
    return p.lineRule && p.lineRule !== 'auto' ? null : (p.line || 240) / 240;
  };

  const b = {
    bold: btn(el('b', {}, 'B'), 'Bold (Ctrl+B)', cmd.toggle('bold')),
    italic: btn(el('i', {}, 'I'), 'Italic (Ctrl+I)', cmd.toggle('italic')),
    underline: btn(el('u', {}, 'U'), 'Underline (Ctrl+U)', cmd.toggle('underline')),
    strike: btn(el('s', {}, 'ab'), 'Strikethrough', cmd.toggle('strike')),
    sub: btn(el('span', {}, 'x', el('sub', {}, '2')), 'Subscript', (s, d) => cmd.setMark('vertAlign', cmd.markValue(s, 'vertAlign') === 'subscript' ? null : { val: 'subscript' })(s, d)),
    sup: btn(el('span', {}, 'x', el('sup', {}, '2')), 'Superscript', (s, d) => cmd.setMark('vertAlign', cmd.markValue(s, 'vertAlign') === 'superscript' ? null : { val: 'superscript' })(s, d)),
    left: btn('⫷', 'Align left (Ctrl+L)', cmd.setAlign('left')),
    center: btn('☰', 'Center (Ctrl+E)', cmd.setAlign('center')),
    right: btn('⫸', 'Align right (Ctrl+R)', cmd.setAlign('right')),
    justify: btn('▤', 'Justify (Ctrl+J)', cmd.setAlign('both')),
    ltr: btn('⁋→', 'Left-to-right text direction', cmd.setDirection(false)),
    rtl: btn('←¶', 'Right-to-left text direction', cmd.setDirection(true)),
    reading: btn('📖', 'Reading view: read without editing (Esc to leave)', withView(() => { setReading(true); return true; }), { label: 'Reading View', big: true }),
    darkPage: btn('◐', 'Dark mode for the page', () => { const on = !container.classList.contains('odx-darkpage'); container.classList.toggle('odx-darkpage', on); b.darkPage.classList.toggle('on', on); getView()?.dom.parentElement.parentElement.classList.toggle('odx-darkpage', on); return true; }, { label: 'Dark Mode', big: true }),
    ruler: btn('📏', 'Ruler', withView((v) => { b.ruler.classList.toggle('on', ruler.toggle(v)); return true; }), { label: 'Ruler', big: true }),
    marks: btn('¶', 'Show/Hide ¶ (Ctrl+Shift+8)', withView((v) => { const m = v.odx?.main || v; toggleMarks(m); b.marks.classList.toggle('on', marksOn(m.state)); return true; }), { label: 'Show/Hide' }),
    painter: (() => { const x = btn('🖌', 'Format Painter: copy the formatting here, then select text to paint it (double-click to paint several places)', formatPainter, { label: 'Format Painter' }); x.addEventListener('dblclick', stickyPainter); return x; })(),
    nav: btn('☷', 'Navigation pane', toggleNav, { label: 'Navigation', big: true }),
  };
  let wcBtn;
  wcBtn = btn('¹²³', 'Word Count', wordCount(() => wcBtn), { label: 'Word Count', big: true });

  const tableTools = grp('Table',
    col(line(btn('⤒', 'Insert row above', cmd.addRow(false)), btn('⤓', 'Insert row below', cmd.addRow(true)), btn('⇤', 'Insert column left', addColumnBefore), btn('⇥', 'Insert column right', addColumnAfter)),
      line(btn('−R', 'Delete row', deleteRow), btn('−C', 'Delete column', deleteColumn), btn('⊞', 'Merge cells', mergeCells), btn('⊟', 'Split cell', splitCell), btn('✕', 'Delete table', deleteTable))));

  const tabs = {
    File: [fileInput],
    Home: [
      grp('Undo', col(btn('↶', 'Undo (Ctrl+Z)', (st, d, v) => (v?.odx?.undo || undo)(st, d, v)), btn('↷', 'Redo (Ctrl+Y)', (st, d, v) => (v?.odx?.redo || redo)(st, d, v)))),
      grp('Clipboard', btn('📋', 'Paste (Ctrl+V)', clip('paste'), { label: 'Paste', big: true }),
        col(btn('✂', 'Cut (Ctrl+X)', clip('cut'), { label: 'Cut' }), btn('⧉', 'Copy (Ctrl+C)', clip('copy'), { label: 'Copy' }), b.painter)),
      grp('Font', col(
        line(fontSel, sizeSel, btn('A⁺', 'Increase font size', (s, d) => cmd.stepFontSize(1)(s, d)), btn('A⁻', 'Decrease font size', (s, d) => cmd.stepFontSize(-1)(s, d)), btn('A̸', 'Clear formatting', cmd.clearFormatting)),
        line(b.bold, b.italic, b.underline, b.strike, b.sub, b.sup,
          menu('🖍', 'Text highlight color', () => [el('div', { class: 'odx-swatches' }, ...HIGHLIGHTS.map(([n, c]) => el('button', { type: 'button', title: n, style: `background:${c}`, onmousedown: run(cmd.setMark('highlight', { val: n })) }))),
            item('No Color', cmd.setMark('highlight', null))]),
          color))),
      grp('Paragraph', col(
        line(btn('•≡', 'Bullets', listCmd('bullet')), gallery('Bullet library', 'bullets'), btn('1≡', 'Numbering', listCmd('decimal')), gallery('Numbering library', 'numbering'), gallery('Multilevel list', 'multilevel', '⋮≡'), btn('⇤', 'Decrease indent', indentCmd(-1)), btn('⇥', 'Increase indent', indentCmd(1))),
        line(b.left, b.center, b.right, b.justify, b.ltr, b.rtl, b.marks,
          menu('▦', 'Shading', () => [el('div', { class: 'odx-swatches' }, ...SHADES.map((c) => el('button', { type: 'button', title: c, style: `background:#${c}`, onmousedown: run(cmd.setShading(c)) }))), item('No color', cmd.setShading(null))]),
          menu('⊞', 'Borders', () => [['Bottom border', 'bottom'], ['Top border', 'top'], ['Left border', 'left'], ['Right border', 'right'], ['No border', 'none'], ['All borders', 'all'], ['Outside borders', 'outside']].map(([n, w]) => item(n, cmd.setBorders(w)))),
          menu('↕', 'Line and paragraph spacing', () => {
            const v = getView();
            const cur = lineSpacing(v.state);
            const p = cmd.selectedParagraphs(v.state)[0]?.node.attrs.view || {};
            return [...LINE_SPACING.map((n) => item(n.toFixed(n % 1 ? 2 : 1).replace(/0$/, '') , cmd.setSpacing({ line: Math.round(n * 240), lineRule: 'auto' }), { on: cur != null && Math.abs(cur - n) < 0.01 })),
              el('hr'),
              item(p.before ? 'Remove space before paragraph' : 'Add space before paragraph', cmd.setSpacing({ before: p.before ? 0 : 240 })),
              item(p.after ? 'Remove space after paragraph' : 'Add space after paragraph', cmd.setSpacing({ after: p.after ? 0 : 240 }))];
          })))),
      grp('Styles', styleGal),
      grp('Editing', col(btn('🔍', 'Find (Ctrl+F)', openFind(false), { label: 'Find' }), btn('⇄', 'Replace (Ctrl+H)', openFind(true), { label: 'Replace' }),
        btn('▣', 'Select all (Ctrl+A)', (s, d) => { d(s.tr.setSelection(new AllSelection(s.doc))); return true; }, { label: 'Select all' }))),
      tableTools,
    ],
    Insert: [
      grp('Pages', btn('⤓', 'Page break (Ctrl+Enter)', cmd.insertBreak('page'), { label: 'Page Break', big: true })),
      grp('Tables', menu('▦', 'Insert table', () => [2, 3, 4, 5, 6].flatMap((r) => [2, 3, 4].map((c) => item(`${c} × ${r} table`, insertTable(r, c))).slice(r === 3 ? 0 : 1, r === 3 ? 3 : 2)), { label: 'Table', big: true })),
      grp('Illustrations', btn('🖼', 'Insert a picture from this device', () => picInput.click(), { label: 'Picture', big: true }), picInput),
      grp('Text', btn('▭', 'Insert a text box', () => (insertTextBox(), true), { label: 'Text Box', big: true })),
      grp('Table of Contents', btn('☰', 'Insert table of contents', withView(toc.insertToc), { label: 'Table of Contents', big: true })),
      grp('Bookmarks', menu('🔖', 'Bookmark', () => {
        const v = getView();
        const input = el('input', { type: 'text', class: 'odx-num odx-bmname', placeholder: 'Bookmark name' });
        const msg = el('div', { class: 'odx-cap' });
        const doAdd = (e) => { e.preventDefault(); const err = bm.addBookmark(getView(), input.value.trim()); if (err) msg.textContent = err; else closeMenu(); };
        const add = el('button', { type: 'button', class: 'odx-item', onmousedown: doAdd }, 'Add');
        input.addEventListener('keydown', (e) => { if (e.key === 'Enter') doAdd(e); e.stopPropagation(); });
        setTimeout(() => input.focus(), 0);
        const list = bm.listBookmarks(v.state.doc).map((x) => el('div', { class: 'odx-bmrow' }, el('span', {}, x.name),
          el('button', { type: 'button', onmousedown: (e) => { e.preventDefault(); closeMenu(); bm.goToBookmark(getView(), x.name); } }, 'Go To'),
          el('button', { type: 'button', onmousedown: (e) => { e.preventDefault(); closeMenu(); bm.deleteBookmark(getView(), x.name); } }, 'Delete')));
        return [el('div', { class: 'odx-bm' }, el('div', { class: 'odx-line' }, input, add), msg, ...list)];
      }, { label: 'Bookmark', big: true })),
      grp('Links', btn('🔗', 'Insert or edit link (Ctrl+K)', editLink, { label: 'Link', big: true }), linkBox),
      grp('Comments', btn('💬', 'New comment (Ctrl+Alt+M)', withView(cm.addComment), { label: 'New Comment', big: true })),
      grp('Header & Footer', btn('▭', 'Edit header and footer (or double-click the top of a page)', withView((v) => hf.editStory(v.odx.main || v, 'header')), { label: 'Header & Footer', big: true }),
        menu('#', 'Page numbers', () => [
          ...['Left', 'Center', 'Right'].map((a) => item(`Top of page, ${a.toLowerCase()}`, withView((v) => hf.insertPageNumbers(v.odx.main || v, 'top', a)))),
          ...['Left', 'Center', 'Right'].map((a) => item(`Bottom of page, ${a.toLowerCase()}`, withView((v) => hf.insertPageNumbers(v.odx.main || v, 'bottom', a)))),
          item('Bottom of page, "Page X of Y"', withView((v) => hf.insertPageNumbers(v.odx.main || v, 'bottom', 'Center', true))),
          item('Current position', (s, d) => { d && d(s.tr.replaceSelectionWith(hf.pageField())); return true; }),
        ], { label: 'Page Numbers', big: true })),
      grp('Symbols', menu('Ω', 'Insert symbol', () => [el('div', { class: 'odx-symbols' }, ...[...SYMBOLS].map((ch) => el('button', { type: 'button', onmousedown: run((s, d) => { d(s.tr.insertText(ch).scrollIntoView()); return true; }) }, ch)))], { label: 'Symbol', big: true })),
    ],
    Layout: [
      grp('Page Setup',
        menu('▯', 'Margins', () => MARGINS.map(([n, t, r, bm, l]) => item(n, setMargins(t, r, bm, l), { hint: `Top ${t / 1440}" Bottom ${bm / 1440}" Left ${l / 1440}" Right ${r / 1440}"` })), { label: 'Margins', big: true }),
        menu('⟲', 'Orientation', () => [item('Portrait', setOrientation(false)), item('Landscape', setOrientation(true))], { label: 'Orientation', big: true }),
        menu('▤', 'Page size', () => PAGE_SIZES.map(([n, hint, w, h]) => item(n, setSize(w, h), { hint })), { label: 'Size', big: true }),
        menu('▥', 'Columns', () => [['One', 1], ['Two', 2], ['Three', 3]].map(([n, k]) => item(n, setColumns(k), { on: !!getView() && sp.columnCount(getView().state.doc.attrs.sectPr) === k })), { label: 'Columns', big: true }),
        menu('⤓', 'Breaks', () => [item('Page break', cmd.insertBreak('page'), { hint: 'Ctrl+Enter' }), item('Column break', cmd.insertBreak('column')), item('Text wrapping break', cmd.insertBreak(null), { hint: 'Shift+Enter' }), el('hr'),
          item('Section break: next page', cmd.insertSectionBreak('nextPage'), { hint: 'New section on the next page' }), item('Section break: continuous', cmd.insertSectionBreak('continuous'), { hint: 'New section on the same page' })], { label: 'Breaks', big: true }),
        menu('⚙', 'Page Setup: margins, paper and orientation, for this section or the whole document', () => { const v = getView(); return v ? pageSetupPanel(v.odx?.main || v, PAGE_SIZES, typed, (fn, all) => applyPageSetup(v, fn, all), closeMenu) : []; }, { label: 'Page Setup', big: true })),
      grp('Arrange', menu('⊡', 'Wrap Text: how the text goes round the selected picture or shape', () => WRAPS.map(([k, n]) => item(n, setWrapText(k), { on: !!getView() && currentWrap(getView().state) === k })), { label: 'Wrap Text', big: true })),
      grp('Page Background',
        menu('🎨', 'Page color', () => [el('div', { class: 'odx-swatches' }, ...PAGE_COLORS.map((c) => el('button', { type: 'button', title: `#${c}`, style: `background:#${c}`, onmousedown: run(setPageColor(c)) }))), item('No Color', setPageColor(null))], { label: 'Page Color', big: true }),
        menu('▢', 'Page borders', () => [['None', 'none'], ['Box', 'box'], ['Double', 'double'], ['Thick', 'thick'], ['Shadow', 'shadow']].map(([n, k]) => item(n, cmd.setPageBorders(k))), { label: 'Page Borders', big: true })),
      grp('Paragraph',
        col(el('span', { class: 'odx-cap' }, 'Indent'), line(el('label', {}, 'Left ', indLeft, '"')), line(el('label', {}, 'Right ', indRight, '"'))),
        col(el('span', { class: 'odx-cap' }, 'Spacing'), line(el('label', {}, 'Before ', spBefore, ' pt')), line(el('label', {}, 'After ', spAfter, ' pt')))),
    ],
    References: [
      grp('Table of Contents', menu('☰', 'Table of contents', () => [
        item('Automatic table', withView(toc.insertToc), { hint: 'Headings 1-3, with page numbers' }),
        item('Update table', withView(toc.updateToc)),
      ], { label: 'Table of Contents', big: true }),
      btn('↻', 'Update fields: page numbers, cross-references and the table of contents, from the pages as they are now (F9)', withView((v) => updateFieldsCommand(v, { all: true })), { label: 'Update Fields', big: true })),
      grp('Footnotes', btn('ᵃᵇ¹', 'Insert footnote (Ctrl+Alt+F)', withView((v) => notes.insertNote(v.odx.main || v, 'footnote')), { label: 'Insert Footnote', big: true }),
        btn('ⁱ', 'Insert endnote (Ctrl+Alt+D)', withView((v) => notes.insertNote(v.odx.main || v, 'endnote')), { label: 'Insert Endnote', big: true }),
        menu('↧', 'Next footnote', () => [item('Next footnote', withView((v) => notes.nextNote(v.odx.main || v, 'footnote'))), item('Next endnote', withView((v) => notes.nextNote(v.odx.main || v, 'endnote')))], { label: 'Next Footnote', big: true })),
    ],
    Review: [
      grp('Proofing', wcBtn, menu('♿', 'Check accessibility', () => {
        const v = getView();
        const issues = checkAccessibility(v.odx?.main || v);
        if (!issues.length) return [el('div', { class: 'odx-stats' }, '✓ No accessibility issues found.')];
        return [el('div', { class: 'odx-a11y' }, el('b', {}, `${issues.length} issue${issues.length === 1 ? '' : 's'}`),
          ...issues.map((x) => el('button', { type: 'button', class: `odx-item odx-a11y-${x.kind}`, onmousedown: (e) => {
            e.preventDefault(); closeMenu();
            const m = getView();
            if (x.pos != null) { m.dispatch(m.state.tr.setSelection(TextSelection.near(m.state.doc.resolve(Math.min(x.pos + 1, m.state.doc.content.size)))).scrollIntoView()); m.focus(); }
          } }, el('span', {}, x.title), el('small', {}, x.detail))))];
      }, { label: 'Check Accessibility', big: true })),
      grp('Comments', btn('💬', 'New comment (Ctrl+Alt+M)', withView(cm.addComment), { label: 'New Comment', big: true }),
        btn('🗑', 'Delete comment', withView((v) => v.odx.session.activeComment != null && cm.deleteComment(v, v.odx.session.activeComment)), { label: 'Delete', big: true }),
        btn('⇠', 'Previous comment', withView((v) => cm.gotoComment(v, -1)), { label: 'Previous', big: true }),
        btn('⇢', 'Next comment', withView((v) => cm.gotoComment(v, 1)), { label: 'Next', big: true })),
      grp('Markup', col(el('span', { class: 'odx-cap' }, 'Markup view:'), markupSel)),
      ...(protection ? [grp('Protect', menu('🔒', 'Restrict editing', restrictPanel, { label: 'Restrict Editing', big: true }))] : []),
      grp('Tracking', col(el('span', { class: 'odx-cap' }, 'Track Changes:'), trackSel),
        menu('✓', 'Accept', () => [item('Accept this change', resolveChange(true)), item('Accept all changes', resolveAll(true))], { label: 'Accept', big: true }),
        menu('✗', 'Reject', () => [item('Reject this change', resolveChange(false)), item('Reject all changes', resolveAll(false))], { label: 'Reject', big: true }),
        btn('◀', 'Previous change', gotoChange(-1), { label: 'Previous', big: true }),
        btn('▶', 'Next change', gotoChange(1), { label: 'Next', big: true })),
    ],
    View: [
      grp('Views', b.reading, b.darkPage),
      // (Compact, to fit a narrow window's ribbon: out, the percentage and in on one line; the fits beside them.)
      grp('Zoom', col(el('span', { class: 'odx-cap' }, 'Zoom:'), line(btn('−', 'Zoom out', withView((v) => zoomOfView(v).step(-1))), zoomSel, btn('+', 'Zoom in', withView((v) => zoomOfView(v).step(1))))),
        col(btn('▣', 'Zoom to 100%', withView((v) => zoomOfView(v).set(1)), { label: '100%' }),
          btn('▯', 'One page: fit a whole page in the window', withView((v) => zoomOfView(v).set('onePage')), { label: 'One Page' }),
          btn('↔', 'Page width: fit the page to the window\'s width', withView((v) => zoomOfView(v).set('pageWidth')), { label: 'Page Width' }))),
      grp('Show', b.ruler, b.nav),
      grp('Print', btn('🖶', 'Print, or Save as PDF from the print dialog', () => window.print(), { label: 'Print', big: true })),
      grp('Header & Footer', btn('▭', 'Edit header and footer', withView((v) => hf.editStory(v.odx.main || v, 'header')), { label: 'Header & Footer', big: true })),
    ],
  };

  // Tab strip and panels.
  let current = 'Home';
  const panels = {};
  const tabBtns = {};
  const strip = el('div', { class: 'odx-tabs' });
  const tablist = el('div', { class: 'odx-tablist', role: 'tablist', 'aria-label': 'Ribbon tabs' });
  strip.append(tablist);
  let panelSeq = 0;
  const ribbonId = `odx-rb${Math.random().toString(36).slice(2, 7)}`;
  const body = el('div', { class: 'odx-ribbon' });
  for (const [name, groups] of Object.entries(tabs)) {
    const id = `${ribbonId}-${panelSeq++}`;
    panels[name] = el('div', { class: 'odx-panel', role: 'tabpanel', id: `${id}-p`, 'aria-labelledby': `${id}-t`, hidden: name !== current }, ...groups);
    tabBtns[name] = el('button', { type: 'button', role: 'tab', id: `${id}-t`, 'aria-controls': `${id}-p`, 'aria-selected': String(name === current), tabindex: name === current ? '0' : '-1', class: `odx-rtab${name === current ? ' on' : ''}`, onmousedown: (e) => { e.preventDefault(); if (name === 'File') openFileMenu(tabBtns.File); else showTab(name); } }, name);
    tabBtns[name].addEventListener('keydown', (e) => {
      const names = Object.keys(tabBtns).filter((k) => !tabBtns[k].hidden);
      const i = names.indexOf(name);
      const to = { ArrowRight: names[(i + 1) % names.length], ArrowLeft: names[(i - 1 + names.length) % names.length], Home: names[0], End: names.at(-1) }[e.key];
      if (to) { e.preventDefault(); if (to !== 'File') showTab(to); tabBtns[to].focus(); }
    });
    tablist.append(tabBtns[name]);
    body.append(panels[name]);
  }
  // Groups that don't fit on the row are hidden behind "More" (as Word does); it opens the rest as further rows.
  const moreBtn = el('button', { type: 'button', class: 'odx-more', title: 'More commands', 'aria-label': 'More commands', 'aria-expanded': 'false', hidden: true,
    onmousedown: (e) => { e.preventDefault(); const on = !body.classList.contains('odx-expand'); body.classList.toggle('odx-expand', on); moreBtn.setAttribute('aria-expanded', String(on)); fitRibbon(); } }, '⋯');
  body.append(moreBtn);
  function fitRibbon() {
    const panel = panels[current];
    const groups = Object.values(panels).flatMap((p) => [...p.children]);
    groups.forEach((g) => g.classList.remove('odx-cut'));
    if (!panel || panel.hidden) return;
    const room = panel.clientWidth - 36;
    const cut = [...panel.children].filter((g) => !g.hidden && g.offsetLeft + g.offsetWidth > room);
    const over = cut.length > 0 || body.classList.contains('odx-expand') && panel.scrollHeight > panel.firstElementChild?.offsetHeight + 8;
    moreBtn.hidden = !over;
    if (!over) { body.classList.remove('odx-expand'); moreBtn.setAttribute('aria-expanded', 'false'); }
    if (!body.classList.contains('odx-expand')) cut.forEach((g) => g.classList.add('odx-cut'));
  }
  if (typeof ResizeObserver === 'function') new ResizeObserver(() => fitRibbon()).observe(body);
  // File is a drop-down menu in Word for the web, not a ribbon tab.
  const fileItem = (icon, label, fn, hint = '') => el('button', { type: 'button', class: 'odx-item odx-fileitem', title: hint || null, onmousedown: (e) => { e.preventDefault(); closeMenu(); fn(); } },
    el('span', { class: 'odx-ctxico', 'aria-hidden': 'true' }, icon), el('span', {}, label), hint.startsWith('Ctrl') && el('small', {}, hint));
  function openFileMenu(owner) {
    openMenu(owner, () => [el('div', { class: 'odx-filemenu', role: 'none' },
      // (Off with features { newDocument: false, openFile: false }: an editor tied to one file, as a WOPI host's is.)
      features.newDocument !== false && fileItem('⊕', 'New', () => onNew()),
      features.openFile !== false && fileItem('📂', 'Open', () => fileInput.click()),
      el('hr'),
      fileItem('⧉', 'Create a Copy', () => onSave(), 'Download a copy as .docx'),
      onExportPdf ? fileItem('↦', 'Export as PDF', () => onExportPdf(), 'Download a PDF of the pages') : fileItem('↦', 'Export', () => window.print(), 'Save as PDF from the print dialog'),
      fileItem('🖶', 'Print', () => window.print(), 'Ctrl+P'),
      el('hr'),
      fileItem('ⓘ', 'Info', () => { showTab('Review'); setTimeout(() => wcBtn.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }))); }, 'Word count and document details'))]);
  }
  function showTab(name) {
    current = name;
    closeMenu();
    for (const k of Object.keys(panels)) {
      panels[k].hidden = k !== name;
      tabBtns[k].classList.toggle('on', k === name);
      tabBtns[k].setAttribute('aria-selected', String(k === name));
      tabBtns[k].tabIndex = k === name ? 0 : -1;
    }
    fitRibbon();
  }

  // Reading view: the ribbon folds away and the document can't be edited until Esc or "Close".
  const exitRead = el('button', { type: 'button', class: 'odx-exitread', hidden: true, onmousedown: (e) => { e.preventDefault(); setReading(false); } }, 'Close Reading View');
  function setReading(on) {
    const v = getView();
    const main = v?.odx?.main || v;
    container.classList.toggle('odx-reading', on);
    exitRead.hidden = !on;
    main?.setProps({ editable: () => !on });
    main?.dom.parentElement.parentElement.classList.toggle('odx-reading-doc', on);
  }
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && container.classList.contains('odx-reading')) setReading(false); });

  // Simplified ribbon (one line, as Word for the web's default) or the classic one; remembered per browser.
  let simple = false;
  try { simple = localStorage.getItem('odx-ribbon') === 'simple'; } catch { /* storage blocked */ }
  const layoutBtn = el('button', { type: 'button', class: 'odx-ribbon-mode', onmousedown: (e) => { e.preventDefault(); setSimple(!simple); } });
  function setSimple(on) {
    simple = on;
    container.classList.toggle('odx-simple', on);
    layoutBtn.textContent = on ? '⌄' : '⌃';
    layoutBtn.title = on ? 'Classic ribbon' : 'Simplified ribbon';
    layoutBtn.setAttribute('aria-label', layoutBtn.title);
    try { localStorage.setItem('odx-ribbon', on ? 'simple' : 'classic'); } catch { /* storage blocked */ }
    fitRibbon();
  }
  // Command search (Alt+Q): every ribbon command by name (cmdsearch.js).
  const cmdSearch = createCommandSearch({ panels, showTab, getView, container,
    reveal: (c) => { if (c.closest('.odx-cut')) { body.classList.add('odx-expand'); moreBtn.setAttribute('aria-expanded', 'true'); fitRibbon(); } } });
  strip.append(cmdSearch.box, el('span', { class: 'odx-tabs-fill' }), exitRead, layoutBtn);
  setSimple(simple);

  container.classList.add('odx-toolbar');

  // --- Right-click: the context menu (ctxmenu.js) ---
  const ctx = createContextMenu(container, getView, kit, { fonts: FONTS, sizes: SIZES, fontSel, sizeSel, clip, pasteFrom, formatPainter, bullets: listCmd('bullet'),
    showTab, panel: (name) => panels[name], editLink, comments: () => features.comments !== false });
  container.replaceChildren(strip, body, findBar, pop, ctx);
  iconize(container);
  // Toggle buttons tell screen readers whether they're on (4.1.2), kept in step with their 'on' class.
  for (const k of ['bold', 'italic', 'underline', 'strike', 'sub', 'sup', 'left', 'center', 'right', 'justify', 'painter', 'nav', 'darkPage', 'ruler', 'marks']) {
    const x = b[k];
    if (!x) continue;
    const sync = () => x.setAttribute('aria-pressed', String(x.classList.contains('on')));
    sync();
    new MutationObserver(sync).observe(x, { attributes: true, attributeFilter: ['class'] });
  }
  // F6 moves between the document and the ribbon, as in Word; Esc in the ribbon goes back to the document.
  document.addEventListener('keydown', (e) => {
    if (e.key !== 'F6' || !container.isConnected || container.offsetParent === null) return;
    const v = getView();
    e.preventDefault();
    if (container.contains(document.activeElement)) v?.focus();
    else tabBtns[current]?.focus();
  });
  container.addEventListener('keydown', (e) => { if (e.key === 'Escape' && pop.hidden && !findBar.contains(e.target)) { e.preventDefault(); getView()?.focus(); } });

  // The navigation pane sits beside the pages once a document is mounted.
  const placeNav = () => {
    const canvas = getView()?.dom.closest('.odx-canvas');
    if (canvas && navPane.parentElement !== canvas.parentElement) canvas.before(navPane);
  };
  let navTimer = 0;

  /**
   * Turn ribbon features off or on: a tab by name ({ review: false }) or a group by its
   * label, on every tab ({ comments: false, tracking: false, 'header-and-footer': false }).
   */
  let features = {};
  function applyFeatures() {
    let first = null;
    for (const name of Object.keys(panels)) {
      const off = features[slug(name)] === false;
      tabBtns[name].hidden = off;
      if (!off && !first) first = name;
      panels[name].querySelectorAll('.odx-group').forEach((g) => { g.hidden = features[g.dataset.feature] === false; });
    }
    if (features[slug(current)] === false && first) showTab(first);
    fitRibbon();
  }

  return {
    setFeatures(f) { features = { ...features, ...f }; applyFeatures(); },
    get features() { return { ...features }; },
    /** Refresh the styles list after a document is loaded. */
    setDocument(session) {
      if (ruler.shown) { ruler.toggle(getView()); ruler.toggle(getView()); }
      const def = session.styles.defaultParagraph;
      const cfg = getView() && protectionOf(getView().state);
      const allowed = (s) => !(cfg?.formatting && (cfg.locked.has(s.id) || cfg.locked.has(String(s.name).toLowerCase())));
      styleList = [{ id: null, name: 'Normal' }, ...session.styleMenu().filter((x) => x.id !== def && allowed(x))];
      const quick = ['Normal', 'No Spacing', 'Heading 1', 'Heading 2', 'Heading 3', 'Title', 'Subtitle'];
      styleTiles.replaceChildren(...quick.map((n) => styleList.find((x) => x.name.toLowerCase() === n.toLowerCase())).filter(Boolean).slice(0, 4).map(tile));
      const scope = [...(getView()?.dom.parentElement.classList || [])].find((c) => c.startsWith('odx-doc-'));
      for (const c of [...styleGal.classList]) if (c.startsWith('odx-doc-')) styleGal.classList.remove(c);
      if (scope) styleGal.classList.add(scope);
      docScope = scope || '';
      styleSel.replaceChildren(el('option', { value: '' }, 'Normal'), ...session.styleMenu().filter((s) => s.id !== session.styles.defaultParagraph && allowed(s)).map((s) => el('option', { value: s.id }, s.name)));
      placeNav();
      fillNav();
    },
    update(view) {
      const s = view.state;
      ruler.update(view);
      for (const k of ['bold', 'italic', 'underline', 'strike']) b[k].classList.toggle('on', cmd.markActive(s, k));
      const va = cmd.markValue(s, 'vertAlign');
      b.sup.classList.toggle('on', va === 'superscript');
      b.sub.classList.toggle('on', va === 'subscript');
      const p = cmd.selectedParagraphs(s)[0]?.node.attrs;
      const styleIdNow = p && p.styleId && p.styleId !== view.odx.session.styles.defaultParagraph ? p.styleId : '';
      if (![...styleSel.options].some((o) => o.value === styleIdNow)) styleSel.append(el('option', { value: styleIdNow }, styleIdNow));
      styleSel.value = styleIdNow;
      const align = (p && p.align) || 'left';
      b.left.classList.toggle('on', align === 'left' || align === 'start');
      b.center.classList.toggle('on', align === 'center');
      b.right.classList.toggle('on', align === 'right' || align === 'end');
      b.justify.classList.toggle('on', align === 'both');
      // With no direct formatting, show what the text actually uses (its style's font and size), as Word does.
      const shown = (() => { try { const { node, offset } = view.domAtPos(s.selection.from); const e = (node.nodeType === 1 ? node.childNodes[offset] || node : node.parentElement); return getComputedStyle(e.nodeType === 1 ? e : e.parentElement); } catch { return null; } })();
      const size = cmd.markValue(s, 'size') || (shown ? Math.round(parseFloat(shown.fontSize) * 0.75 * 2) : null);
      if (size && ![...sizeSel.options].some((o) => o.value === String(size / 2))) sizeSel.append(el('option', { value: size / 2 }, String(size / 2)));
      sizeSel.value = size ? String(size / 2) : '';
      const font = cmd.markValue(s, 'font') || (shown ? shown.fontFamily.split(',')[0].replace(/["']/g, '').trim() : null);
      if (font && ![...fontSel.options].some((o) => o.value === font)) fontSel.append(el('option', { value: font }, font));
      fontSel.value = font || '';
      const sid = p ? p.styleId || '' : null;
      styleGal.querySelectorAll('.odx-stile').forEach((t) => t.classList.toggle('on', t.dataset.id === sid));
      const pv = (p && p.view) || {};
      const inch = (t) => (t ? +(t / 1440).toFixed(2) : 0);
      if (document.activeElement !== indLeft) indLeft.value = inch(pv.indLeft);
      if (document.activeElement !== indRight) indRight.value = inch(pv.indRight);
      if (document.activeElement !== spBefore) spBefore.value = (pv.before || 0) / 20;
      if (document.activeElement !== spAfter) spAfter.value = (pv.after || 0) / 20;
      tableTools.hidden = !isInTable(s);
      trackSel.value = tracking(s).on ? 'on' : 'off';
      // The zoom now, as a percentage (one that isn't a step, such as Page Width's, added to the list).
      const z = (view.odx?.main || view).odx?.zoom?.value ?? 1;
      if (![...zoomSel.options].some((o) => Number(o.value) === z)) zoomSel.append(el('option', { value: String(z), 'data-fit': '1' }, `${Math.round(z * 100)}%`));
      zoomSel.querySelectorAll('[data-fit]').forEach((o) => { if (Number(o.value) !== z) o.remove(); });
      zoomSel.value = String(z);
      placeNav();
      clearTimeout(navTimer);
      navTimer = setTimeout(fillNav, 400);
    },
  };
}


function insertTable(rows, cols) {
  return (state, dispatch, view) => {
    const { table, table_row, table_cell, paragraph } = schema.nodes;
    // As wide as the text where it goes, as Word's Insert Table: the page's text width (or its column's, or its cell's)
    // less nothing for the paragraph's indents; 6.5" (the Letter page's at 1" margins) when that can't be measured.
    let text = 624;
    try {
      const at = view && view.domAtPos(state.selection.from).node;
      const p = at && (at.nodeType === 1 ? at : at.parentElement)?.closest('p');
      if (p) {
        const cs = getComputedStyle(p);
        const width = p.offsetWidth + (parseFloat(cs.marginLeft) || 0) + (parseFloat(cs.marginRight) || 0);
        if (width > 48) text = width;
      }
    } catch { /* the default */ }
    const w = Math.floor(text / cols);
    const cell = () => table_cell.create({ colwidth: [w] }, paragraph.create());
    const t = table.create(null, Array.from({ length: rows }, () => table_row.create(null, Array.from({ length: cols }, cell))));
    if (dispatch) {
      const tr = state.tr.replaceSelectionWith(t);
      // put the caret in the first cell of the new table
      let start = null;
      tr.doc.nodesBetween(tr.mapping.map(state.selection.from, -1), tr.doc.content.size, (n, pos) => {
        if (start == null && n.type === table_cell) start = pos + 2;
        return start == null;
      });
      if (start != null) tr.setSelection(TextSelection.create(tr.doc, start));
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}
