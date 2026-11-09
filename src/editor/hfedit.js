// Editing headers and footers, as in Word: double-click the top or bottom
// margin of a page (or use Insert > Header & Footer) and the header or footer
// opens in its own small editor over the page. Changes are saved to its part.
import { textSpansPlugin } from './textspans.js';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo, redo } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import { baseKeymap } from 'prosemirror-commands';
import { schema } from './schema.js';
import * as cmd from './commands.js';

const FIELD = (instr, shown) => schema.nodes.opaque_inline.create({
  kind: 'fldSimple',
  xml: `<w:fldSimple w:instr=" ${instr} "><w:r><w:t>${shown}</w:t></w:r></w:fldSimple>`,
  text: shown,
});
export const pageField = () => FIELD('PAGE', '1');
export const numPagesField = () => FIELD('NUMPAGES', '1');

/**
 * Section `index`'s w:sectPr (sections in document order; the last one is the document's own) and a function
 * that replaces it.
 */
function section(main, index) {
  const ends = [];
  main.state.doc.forEach((node, pos) => {
    const e = node.type === schema.nodes.paragraph && (node.attrs.extra || []).find((x) => x[0] === 'sectPr');
    if (e) ends.push({ node, pos, xml: e[1] });
  });
  const end = index < ends.length ? ends[index] : null;
  return {
    xml: end ? end.xml : main.state.doc.attrs.sectPr || '',
    set(xml) {
      const tr = end
        ? main.state.tr.setNodeMarkup(end.pos, null, { ...end.node.attrs, extra: end.node.attrs.extra.map((x) => (x[0] === 'sectPr' ? ['sectPr', xml] : x)) })
        : main.state.tr.setDocAttribute('sectPr', xml);
      main.dispatch(tr.setMeta('odx-untracked', true));
    },
  };
}

/** Turn "Different First Page" (w:titlePg) on or off for section `index`. */
export function setDifferentFirstPage(main, index, on) {
  const s = section(main, index);
  let xml = (s.xml || '<w:sectPr></w:sectPr>').replace(/<w:titlePg\b[^>]*\/>/g, '');
  if (on) {
    // Its place in w:sectPr: before these, else at the end.
    const after = /<w:(?:textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)\b/.exec(xml);
    xml = after ? xml.slice(0, after.index) + '<w:titlePg/>' + xml.slice(after.index) : xml.replace(/<\/w:sectPr>$/, '<w:titlePg/></w:sectPr>');
  }
  s.set(xml);
}

const REF = (kind, type) => new RegExp(`<w:${kind}Reference\\b[^>]*w:type="${type}"[^>]*>`);

/** True when section `index` has no `kind`/`type` header or footer of its own and so shows the previous section's. */
export function linkedToPrevious(main, kind, type, index) {
  return index > 0 && !REF(kind, type).test(section(main, index).xml);
}

/**
 * Word's "Link to Previous". Off: the section gets its own copy of what it showed. On: its own one is dropped
 * and it shows the previous section's again (the part stays in the package, unreferenced, as Word leaves it).
 */
export function setLinkToPrevious(main, kind, type, index, hf, on) {
  const s = section(main, index);
  if (on) { s.set(s.xml.replace(REF(kind, type), '')); return; }
  if (!linkedToPrevious(main, kind, type, index)) return;
  const { session } = main.odx;
  const inherited = hf && hf[kind][type];
  const rId = session.addStory(kind);
  if (inherited) session.setStory(rId, inherited);
  const ref = `<w:${kind}Reference w:type="${type}" r:id="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`;
  s.set(s.xml ? s.xml.replace(/^(<w:sectPr\b[^>]*?)(\/?)>/, (m, open, selfClose) => `${open}>${ref}${selfClose ? '</w:sectPr>' : ''}`) : `<w:sectPr>${ref}</w:sectPr>`);
}

/** The rId of `kind` ('header'|'footer') of `type` for section `index`, creating one when there is none. */
function ensureStory(main, kind, type, hf, index) {
  const { session } = main.odx;
  const sect = section(main, index);
  const sectPr = sect.xml;
  const re = new RegExp(`<w:${kind}Reference\\b[^>]*w:type="${type}"[^>]*>`);
  const found = re.exec(sectPr);
  if (found) return /r:id="([^"]+)"/.exec(found[0])[1];
  // Inherited from an earlier section: edit that one (as Word does unless "Link to Previous" is off).
  if (hf && hf[kind][type]) {
    for (const [rId, d] of session._stories || []) if (d === hf[kind][type]) return rId;
  }
  const rId = session.addStory(kind);
  const ref = `<w:${kind}Reference w:type="${type}" r:id="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`;
  // References come first in w:sectPr, headers before footers.
  const next = sectPr
    ? sectPr.replace(/^(<w:sectPr\b[^>]*?)(\/?)>/, (m, open, selfClose) => `${open}>${ref}${selfClose ? '</w:sectPr>' : ''}`)
    : `<w:sectPr>${ref}</w:sectPr>`;
  sect.set(next);
  session.story(rId);
  return rId;
}

let open = null; // { main, hv, el, rId, kind }

export const editingStory = () => open;

/** Save and close the open header/footer editor. */
export function closeStory() {
  if (!open) return;
  const { main, hv, el, rId } = open;
  open = null;
  main.odx.session.setStory(rId, hv.state.doc);
  main.odx.stories?.story(rId, hv.state.doc); // (shared: src/collab.js)
  main.odx.hfView = null;
  hv.destroy();
  el.remove();
  main.dom.classList.remove('odx-dim');
  main.odx.pages.restory();
  main.odx.notify?.(main);
  main.focus();
}

/** Open the header or footer of page `p` (0-based; default: the page with the cursor). */
export function editStory(main, kind, p = null) {
  closeStory();
  const { session, pages } = main.odx;
  if (p == null) {
    const c = main.coordsAtPos(main.state.selection.from);
    p = pages.pageAt(c.top - main.dom.parentElement.getBoundingClientRect().top).p;
  }
  const at = pages.pageAt(pages.pageTop(p) + 1);
  const linked = linkedToPrevious(main, kind, at.type, at.sect);
  const rId = ensureStory(main, kind, at.type, at.hf, at.sect);
  const doc = session.story(rId) || schema.nodes.doc.create(null, schema.nodes.paragraph.create());
  const sheet = main.dom.parentElement;
  const { page } = at;
  const hfDist = at.hf ? (kind === 'header' ? at.hf.headerDist : at.hf.footerDist) : 48;

  const el = document.createElement('div');
  el.className = `odx-hfedit odx-hfedit-${kind}`;
  el.style.left = `${page.left}px`;
  el.style.right = `${page.right}px`;
  if (kind === 'header') el.style.top = `${at.top + hfDist}px`;
  else el.style.bottom = `${sheet.scrollHeight - (at.top + page.height - hfDist)}px`;
  const tag = document.createElement('div');
  tag.className = 'odx-hfedit-tag';
  const label = { first: 'First Page ', even: 'Even Page ', default: '' }[at.type] + (kind === 'header' ? 'Header' : 'Footer') + (at.sect > 0 ? ` - Section ${at.sect + 1}` : '') + (linked ? ' (Same as Previous)' : '');
  const close = document.createElement('button');
  close.type = 'button';
  close.textContent = 'Close Header and Footer';
  close.onmousedown = (e) => { e.preventDefault(); closeStory(); };
  close.onclick = (e) => { if (e.detail === 0) closeStory(); };
  // Word's Header & Footer options, for this section (first page) and the whole document (odd and even).
  const option = (text, checked, apply) => {
    const lab = document.createElement('label');
    lab.className = 'odx-hfedit-opt';
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.checked = checked;
    box.onchange = () => {
      closeStory();
      apply(box.checked);
      main.odx.pages.restory();
      // Reopen on the same page once the pages show the new arrangement.
      setTimeout(() => editStory(main, kind, p), 300);
    };
    lab.append(box, ` ${text}`);
    return lab;
  };
  const firstOn = /<w:titlePg(?![^>]*w:val="(?:0|false|off)")/.test(section(main, at.sect).xml);
  tag.append(label,
    option('Different First Page', firstOn, (on) => setDifferentFirstPage(main, at.sect, on)),
    option('Different Odd & Even Pages', session.evenAndOddHeaders, (on) => { session.evenAndOddHeaders = on; }));
  if (at.sect > 0) tag.append(option('Link to Previous', linked, (on) => setLinkToPrevious(main, kind, at.type, at.sect, at.hf, on)));
  tag.append(close);
  el.append(tag);
  sheet.append(el);

  const keys = keymap({
    'Mod-z': undo, 'Mod-y': redo, 'Mod-Shift-z': redo,
    'Mod-b': cmd.toggle('bold'), 'Mod-i': cmd.toggle('italic'), 'Mod-u': cmd.toggle('underline'),
    'Mod-e': cmd.setAlign('center'), 'Mod-l': cmd.setAlign('left'), 'Mod-r': cmd.setAlign('right'),
    Tab: cmd.insertTab,
    Escape: () => { closeStory(); return true; },
  });
  const hv = new EditorView(el, {
    state: EditorState.create({ doc, plugins: [history(), keys, keymap(baseKeymap), textSpansPlugin] }),
    attributes: { class: 'odx-hf odx-hfedit-content', spellcheck: 'true' },
    dispatchTransaction(tr) {
      hv.updateState(hv.state.apply(tr));
      main.odx.notify?.(hv);
    },
  });
  hv.odx = { ...main.odx, isStory: true };
  main.odx.hfView = hv;
  main.dom.classList.add('odx-dim');
  open = { main, hv, el, rId, kind };
  hv.focus();
  return hv;
}

/** Double-clicking a page's top or bottom margin edits its header or footer; double-clicking the body closes it. */
export function hfDoubleClick(main) {
  const sheet = main.dom.parentElement;
  sheet.addEventListener('dblclick', (e) => {
    if (main.odx.readOnly) return;
    if (open && open.el.contains(e.target)) return;
    if (e.target.closest?.('.odx-txbxedit') || /txbxContent/.test(main.state.selection.node?.attrs.xml || '')) return; // a text box (see txbxedit.js)
    const y = e.clientY - sheet.getBoundingClientRect().top;
    const at = main.odx.pages.pageAt(y);
    const inBody = at.y > at.m.top && at.y < at.page.height - at.m.bottom;
    if (inBody) { if (open) closeStory(); return; }
    if (at.y > at.page.height) return; // the gap between pages
    e.preventDefault();
    editStory(main, at.y <= at.m.top ? 'header' : 'footer', at.p);
  });
}

const PN_ALIGN = { Left: 'left', Center: 'center', Right: 'right' };

/** Page numbers: replace the header or footer with a number aligned left, center or right. */
export function insertPageNumbers(main, where, align, ofTotal = false) {
  const hv = editStory(main, where === 'top' ? 'header' : 'footer');
  const content = ofTotal
    ? [schema.text('Page '), pageField(), schema.text(' of '), numPagesField()]
    : [pageField()];
  const style = where === 'top' ? 'Header' : 'Footer';
  const p = schema.nodes.paragraph.create({ styleId: main.odx.session.styles.styles.has(style) ? style : null, align: PN_ALIGN[align] }, content);
  const tr = hv.state.tr.replaceWith(0, hv.state.doc.content.size, p);
  hv.dispatch(tr);
  closeStory();
}
