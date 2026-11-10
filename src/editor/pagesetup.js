// Page Setup (Layout > Page Setup, Word's dialog): margins, gutter, header and footer distances, orientation and paper
// size, applied to the section the cursor is in or to the whole document. Layout's Margins, Orientation, Size and Columns
// menus change the section the cursor is in too, as Word's do.
import { el } from './ribbonkit.js';
import { sectionEnd } from './sections.js';
import * as sp from '../ooxml/sectpr.js';

/** Every section's end: [{ pos, node, xml }] for the paragraphs that end one, then { pos: null, xml } for the last. */
export function sectionEnds(doc) {
  const out = [];
  doc.forEach((node, pos) => { const xml = sectionEnd(node); if (xml != null) out.push({ pos, node, xml }); });
  out.push({ pos: null, xml: doc.attrs.sectPr || sp.DEFAULT_SECTPR });
  return out;
}

/** The section holding document position `at` (its end, from sectionEnds). */
export function sectionAt(doc, at) {
  const ends = sectionEnds(doc);
  return ends.find((e) => e.pos == null || e.pos + e.node.nodeSize > at);
}

/**
 * A transaction applying `fn` (w:sectPr XML in, XML out) to the section at the cursor, or with `all`, to every section;
 * null when nothing changes.
 */
export function sectionsTr(state, fn, all = false) {
  const tr = state.tr;
  const targets = all ? sectionEnds(state.doc) : [sectionAt(state.doc, state.selection.from)];
  for (const s of targets) {
    const next = fn(s.xml);
    if (next === s.xml) continue;
    if (s.pos == null) tr.setDocAttribute('sectPr', next);
    else tr.setNodeMarkup(s.pos, null, { ...s.node.attrs, extra: s.node.attrs.extra.map((e) => (e[0] === 'sectPr' ? ['sectPr', next] : e)) });
  }
  return tr.steps.length ? tr : null;
}

const IN = 1440; // twips to the inch
const num = (xml, tag, a, d) => Number(new RegExp(`<w:${tag}\\b[^>]*\\sw:${a}="(-?\\d+)"`).exec(xml)?.[1] ?? d);

/**
 * The dialog's content (for a ribbon menu): the section's values in inches, Apply to, OK. `apply(fn, all)` makes the
 * change (toolbar.js: then the document is laid out afresh); `sizes` are Layout's paper sizes [[name, hint, w, h]].
 */
export function pageSetupPanel(view, sizes, typed, apply, close) {
  const many = sectionEnds(view.state.doc).length > 1;
  const xml = sectionAt(view.state.doc, view.state.selection.from).xml;
  const field = (label, twips) => {
    const i = typed({ type: 'number', step: '0.01', min: '0', class: 'odx-num', 'aria-label': `${label} (inches)`, value: String(Math.round((twips / IN) * 100) / 100) });
    return [el('label', { class: 'odx-psrow' }, el('span', {}, label), i, '"'), i];
  };
  const [topRow, top] = field('Top', Math.abs(num(xml, 'pgMar', 'top', 1440)));
  const [bottomRow, bottom] = field('Bottom', Math.abs(num(xml, 'pgMar', 'bottom', 1440)));
  const [leftRow, left] = field('Left', num(xml, 'pgMar', 'left', 1440));
  const [rightRow, right] = field('Right', num(xml, 'pgMar', 'right', 1440));
  const [gutterRow, gutter] = field('Gutter', num(xml, 'pgMar', 'gutter', 0));
  const [headerRow, header] = field('Header', num(xml, 'pgMar', 'header', 720));
  const [footerRow, footer] = field('Footer', num(xml, 'pgMar', 'footer', 720));
  const { w, h } = sp.pageSize(xml);
  const orient = el('select', { 'aria-label': 'Orientation' }, el('option', { value: 'portrait' }, 'Portrait'), el('option', { value: 'landscape' }, 'Landscape'));
  orient.value = w > h ? 'landscape' : 'portrait';
  // Paper: a size of the list (portrait), or Custom with its width and height.
  const [pw, ph] = [Math.min(w, h), Math.max(w, h)];
  const paper = el('select', { 'aria-label': 'Paper size' }, ...sizes.map(([n, hint]) => el('option', { value: n }, `${n} (${hint})`)), el('option', { value: 'custom' }, 'Custom size'));
  paper.value = sizes.find((s) => Math.abs(s[2] - pw) < 20 && Math.abs(s[3] - ph) < 20)?.[0] || 'custom';
  const [widthRow, width] = field('Width', pw);
  const [heightRow, height] = field('Height', ph);
  const custom = () => { const on = paper.value === 'custom'; widthRow.hidden = heightRow.hidden = !on; };
  paper.addEventListener('change', custom);
  custom();
  const scope = el('select', { 'aria-label': 'Apply to' }, el('option', { value: 'section' }, 'This section'), el('option', { value: 'all' }, 'Whole document'));
  scope.value = many ? 'section' : 'all';
  const msg = el('div', { class: 'odx-cap', role: 'status' });
  const tw = (i) => Math.round(Number(i.value) * IN);
  const ok = (e) => {
    e.preventDefault();
    const vals = [top, bottom, left, right, gutter, header, footer, ...(paper.value === 'custom' ? [width, height] : [])];
    if (vals.some((i) => i.value === '' || !(Number(i.value) >= 0))) { msg.textContent = 'Enter each measurement in inches.'; return; }
    if (tw(left) + tw(right) + tw(gutter) >= (paper.value === 'custom' ? tw(width) : sizes.find((s) => s[0] === paper.value)[2]) - 720) { msg.textContent = 'The margins leave no room for the text.'; return; }
    const size = paper.value === 'custom' ? [tw(width), tw(height)] : sizes.find((s) => s[0] === paper.value).slice(2);
    const fn = (x) => {
      let out = sp.setAttrs(x, 'pgMar', { top: tw(top), bottom: tw(bottom), left: tw(left), right: tw(right), gutter: tw(gutter), header: tw(header), footer: tw(footer) });
      out = sp.setAttrs(out, 'pgSz', { w: Math.min(...size), h: Math.max(...size), orient: null });
      return sp.setOrientation(out, orient.value === 'landscape');
    };
    close();
    apply(fn, scope.value === 'all');
  };
  return [el('div', { class: 'odx-protect odx-pagesetup', role: 'group', 'aria-label': 'Page Setup' },
    el('b', {}, 'Margins'), topRow, bottomRow, leftRow, rightRow, gutterRow,
    el('label', { class: 'odx-psrow' }, el('span', {}, 'Orientation'), orient),
    el('b', {}, 'Paper'), el('label', { class: 'odx-psrow' }, el('span', {}, 'Size'), paper), widthRow, heightRow,
    el('b', {}, 'From edge'), headerRow, footerRow,
    el('label', { class: 'odx-psrow' }, el('span', {}, 'Apply to'), scope),
    el('button', { type: 'button', class: 'odx-item', onmousedown: ok }, 'OK'), msg)];
}
