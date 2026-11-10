// Word-like editing commands over the schema.
import { toggleMark } from 'prosemirror-commands';
import { TextSelection } from 'prosemirror-state';
import { addRowAfter, addRowBefore, isInTable, selectedRect, TableMap } from 'prosemirror-tables';
import { schema } from './schema.js';
import { resolveStyle } from '../ooxml/styles.js';
import { pPrXml } from '../ooxml/writer.js';
import { paragraphAttrsFromPPr } from '../ooxml/reader.js';
import { unescapeXml } from '../ooxml/xml.js';

const { paragraph } = schema.nodes;

/** Paragraphs touched by the selection, with positions. */
export function selectedParagraphs(state) {
  const out = [];
  const { from, to } = state.selection;
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type === paragraph) {
      out.push({ node, pos });
      return false;
    }
  });
  return out;
}

function updateParagraphs(fn) {
  return (state, dispatch) => {
    const paras = selectedParagraphs(state);
    if (!paras.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const { node, pos } of paras) tr.setNodeMarkup(pos, null, { ...node.attrs, ...fn(node.attrs, node) });
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

export const setStyle = (styleId) => updateParagraphs(() => ({ styleId }));
export const setAlign = (align) => updateParagraphs((a) => ({ align: a.align === align && align !== 'left' ? null : align }));

/** Effective list of a paragraph: direct numPr, else its style's (numId "0" switches lists off). */
export function paragraphList(session, attrs) {
  let { numId, ilvl } = attrs;
  if (numId == null && attrs.styleId && session) {
    const st = resolveStyle(session.styles.styles, attrs.styleId).pPr;
    if (st.numId != null) ({ numId, ilvl = 0 } = { numId: st.numId, ilvl: st.ilvl });
  }
  return numId == null || numId === '0' ? null : { numId, ilvl: ilvl || 0 };
}

export function listKind(session, attrs) {
  const l = paragraphList(session, attrs);
  if (!l) return null;
  const lvl = session.listLevel(l.numId, l.ilvl);
  return lvl ? (lvl.fmt === 'bullet' ? 'bullet' : 'decimal') : null;
}

export const toggleList = (session, kind) => (state, dispatch) => {
  const paras = selectedParagraphs(state);
  if (!paras.length) return false;
  if (!dispatch) return true;
  const allOn = paras.every(({ node }) => listKind(session, node.attrs) === kind);
  const numId = allOn ? null : session.listId(kind);
  const tr = state.tr;
  for (const { node, pos } of paras) tr.setNodeMarkup(pos, null, { ...node.attrs, numId: allOn && node.attrs.styleId && paragraphList(session, { styleId: node.attrs.styleId }) ? '0' : numId, ilvl: allOn ? 0 : node.attrs.ilvl });
  dispatch(tr.scrollIntoView());
  return true;
};

export const changeListLevel = (session, delta) => (state, dispatch) => {
  const paras = selectedParagraphs(state).filter(({ node }) => paragraphList(session, node.attrs));
  if (!paras.length) return false;
  if (dispatch) {
    const tr = state.tr;
    for (const { node, pos } of paras) {
      const l = paragraphList(session, node.attrs);
      tr.setNodeMarkup(pos, null, { ...node.attrs, numId: l.numId, ilvl: Math.max(0, Math.min(8, l.ilvl + delta)) });
    }
    dispatch(tr);
  }
  return true;
};

/** Enter: Word semantics: lists continue, an empty list item ends the list, headings are followed by their "next" style. */
export const splitParagraph = (session) => (state, dispatch) => {
  const { $from, $to } = state.selection;
  if ($from.parent.type !== paragraph || !$from.sameParent($to)) return false;
  const p = $from.parent;
  if (p.content.size === 0 && paragraphList(session, p.attrs)) {
    if (dispatch) dispatch(state.tr.setNodeMarkup($from.before(), null, { ...p.attrs, numId: p.attrs.styleId && paragraphList(session, { styleId: p.attrs.styleId }) ? '0' : null, ilvl: 0 }));
    return true;
  }
  if (!dispatch) return true;
  const atEnd = $to.parentOffset === p.content.size;
  const st = p.attrs.styleId && session.styles.styles.get(p.attrs.styleId);
  const next = {
    ...p.attrs,
    styleId: atEnd && st && st.next && st.next !== st.id ? st.next : p.attrs.styleId,
    extra: p.attrs.extra.filter(([n]) => n !== 'sectPr'),
  };
  const tr = state.tr.deleteSelection();
  tr.split(tr.mapping.map($from.pos), 1, [{ type: paragraph, attrs: next }]);
  // carry the run formatting of the caret into the new paragraph
  const marks = state.storedMarks || ($from.parentOffset ? $from.marks() : null);
  if (marks) tr.ensureMarks(marks.filter((m) => m.type.name !== 'link'));
  dispatch(tr.scrollIntoView());
  return true;
};

/** Backspace at the start of a list item removes the bullet first, as Word does. */
export const outdentOrUnlist = (session) => (state, dispatch) => {
  const { $from, empty } = state.selection;
  if (!empty || $from.parentOffset !== 0 || $from.parent.type !== paragraph) return false;
  const l = paragraphList(session, $from.parent.attrs);
  if (!l) return false;
  if (dispatch) {
    const attrs = $from.parent.attrs;
    dispatch(state.tr.setNodeMarkup($from.before(), null, l.ilvl > 0 ? { ...attrs, numId: l.numId, ilvl: l.ilvl - 1 } : { ...attrs, numId: attrs.styleId && paragraphList(session, { styleId: attrs.styleId }) ? '0' : null, ilvl: 0 }));
  }
  return true;
};

export const insertTab = (state, dispatch) => {
  if (dispatch) dispatch(state.tr.replaceSelectionWith(schema.nodes.tab.create()).scrollIntoView());
  return true;
};

export const insertBreak = (type = null) => (state, dispatch) => {
  if (dispatch) dispatch(state.tr.replaceSelectionWith(schema.nodes.hard_break.create({ type })).scrollIntoView());
  return true;
};

export const toggle = (name) => toggleMark(schema.marks[name]);

/** Set (or with null, clear) a valued mark such as size, color or font over the selection. */
export const setMark = (name, attrs) => (state, dispatch) => {
  const type = schema.marks[name];
  const { from, to, empty } = state.selection;
  if (!dispatch) return true;
  if (empty) {
    const marks = (state.storedMarks || state.selection.$from.marks()).filter((m) => m.type !== type);
    dispatch(state.tr.setStoredMarks(attrs ? [...marks, type.create(attrs)] : marks));
  } else {
    const tr = state.tr.removeMark(from, to, type);
    if (attrs) tr.addMark(from, to, type.create(attrs));
    dispatch(tr.scrollIntoView());
  }
  return true;
};

export function setLink(href) {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    const type = schema.marks.link;
    if (!dispatch) return true;
    const tr = state.tr;
    if (empty && href) {
      tr.insert(from, schema.text(href, [type.create({ href })]));
      tr.setSelection(TextSelection.create(tr.doc, from + href.length));
    } else {
      tr.removeMark(from, to, type);
      if (href) tr.addMark(from, to, type.create({ href }));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export function markActive(state, name) {
  const type = schema.marks[name];
  const { from, $from, to, empty } = state.selection;
  const isOn = (m) => m && m.type === type && m.attrs.on !== false;
  if (empty) return (state.storedMarks || $from.marks()).some(isOn);
  let found = false;
  state.doc.nodesBetween(from, to, (n) => {
    if (n.isText && n.marks.some(isOn)) found = true;
  });
  return found;
}

export function markValue(state, name) {
  const type = schema.marks[name];
  const marks = state.storedMarks || state.selection.$from.marks();
  const m = marks.find((x) => x.type === type);
  return m ? m.attrs.val ?? m.attrs : null;
}

// --- Paragraph spacing and indents (stored in the paragraph's raw w:spacing / w:ind) ---

/** `attrs` with w:<name> attributes changed (null removes one); the element goes when it has none left. */
function withPPrChild(attrs, name, changes) {
  const extra = attrs.extra || [];
  const cur = extra.find((e) => e[0] === name);
  const vals = {};
  if (cur) for (const m of cur[1].matchAll(/\sw:(\w+)="([^"]*)"/g)) vals[m[1]] = m[2];
  for (const [k, v] of Object.entries(changes)) {
    if (v == null) delete vals[k];
    else vals[k] = String(v);
  }
  const xml = Object.keys(vals).length ? `<w:${name}${Object.entries(vals).map(([k, v]) => ` w:${k}="${v}"`).join('')}/>` : null;
  const rest = extra.filter((e) => e[0] !== name);
  return { ...attrs, extra: xml ? [...rest, [name, xml]] : rest };
}

/** Spacing in twips (before/after) and line (240 = single, lineRule auto). null clears a value. */
export const setSpacing = (changes) => updateParagraphs((a) => {
  const next = withPPrChild(a, 'spacing', changes);
  const view = { ...a.view };
  for (const k of ['before', 'after', 'line', 'lineRule']) {
    if (!(k in changes)) continue;
    if (changes[k] == null) delete view[k];
    else view[k] = k === 'lineRule' ? changes[k] : Number(changes[k]);
  }
  return { extra: next.extra, view };
});

/** Indents in twips: left, right, firstLine (negative = hanging). null clears a value. */
export const setIndent = (changes) => updateParagraphs((a) => {
  const xmlChanges = {};
  const view = { ...a.view };
  if ('left' in changes) { xmlChanges.left = changes.left; xmlChanges.start = null; changes.left == null ? delete view.indLeft : (view.indLeft = changes.left); }
  if ('right' in changes) { xmlChanges.right = changes.right; xmlChanges.end = null; changes.right == null ? delete view.indRight : (view.indRight = changes.right); }
  if ('firstLine' in changes) {
    const f = changes.firstLine;
    xmlChanges.firstLine = f != null && f >= 0 ? f : null;
    xmlChanges.hanging = f != null && f < 0 ? -f : null;
    f == null ? delete view.indFirst : (view.indFirst = f);
  }
  return { extra: withPPrChild(a, 'ind', xmlChanges).extra, view };
});

/** Move the left indent by `delta` twips (Increase / Decrease Indent), never below 0. */
export const shiftIndent = (delta) => (state, dispatch) => {
  const p = selectedParagraphs(state)[0];
  const cur = (p && p.node.attrs.view?.indLeft) || 0;
  return setIndent({ left: Math.max(0, cur + delta) })(state, dispatch);
};

// --- Character formatting helpers ---

const FONT_STEPS = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
/** Grow or shrink the font one step (half-points in the mark). `base` is the size shown when no mark is set. */
export const stepFontSize = (dir, base = 11) => (state, dispatch) => {
  const cur = (markValue(state, 'size') || base * 2) / 2;
  const next = dir > 0 ? FONT_STEPS.find((s) => s > cur) ?? cur + 10 : [...FONT_STEPS].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1);
  return setMark('size', { val: Math.round(next * 2) })(state, dispatch);
};

/** Remove direct character formatting (keeps links, tracked changes and character styles). */
export const clearFormatting = (state, dispatch) => {
  const { from, to } = state.selection;
  if (dispatch) {
    const tr = state.tr;
    for (const name of ['font', 'bold', 'italic', 'strike', 'color', 'size', 'highlight', 'underline', 'vertAlign', 'rextra']) if (schema.marks[name]) tr.removeMark(from, to, schema.marks[name]);
    dispatch(tr);
  }
  return true;
};

/** Character formatting marks at the selection, for Format Painter. */
export function formatAt(state) {
  const keep = (m) => !['link', 'ins', 'del'].includes(m.type.name);
  const { $from } = state.selection;
  const marks = (state.storedMarks || (state.selection.empty ? $from.marks() : (state.doc.nodeAt($from.pos)?.marks || $from.marks()))).filter(keep);
  const p = $from.parent.type === paragraph ? $from.parent.attrs : null;
  return { marks, para: p };
}

// Paragraph formatting Format Painter copies (not the content control, section or tracked-change records).
const PAINT_SKIP_EXTRA = new Set(['sectPr', 'pPrChange', 'rPr']);

/** Replace the selection's character formatting with `marks` (Format Painter's paste). */
export const applyFormat = (fmt) => (state, dispatch) => {
  const marks = Array.isArray(fmt) ? fmt : fmt.marks;
  let { from, to, $from, $to } = state.selection;
  if (from === to) {
    // No selection: paint the word at the caret, as Word does.
    const text = $from.parent.textBetween(0, $from.parent.content.size, '\ufffc', '\ufffc');
    let a = $from.parentOffset, b = a;
    while (a > 0 && /\w/.test(text[a - 1])) a--;
    while (b < text.length && /\w/.test(text[b])) b++;
    from = $from.start() + a;
    to = $from.start() + b;
  }
  if (dispatch) {
    const tr = state.tr;
    if (to > from) {
      for (const name of ['cstyle', 'font', 'bold', 'italic', 'strike', 'color', 'size', 'highlight', 'underline', 'vertAlign', 'rextra']) if (schema.marks[name]) tr.removeMark(from, to, schema.marks[name]);
      for (const m of marks) tr.addMark(from, to, m);
    }
    // Paragraph formatting too when whole paragraphs are painted (the selection reaches a paragraph's end).
    const src = fmt.para;
    if (src && !state.selection.empty && $to.parentOffset === $to.parent.content.size) {
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (node.type !== paragraph) return true;
        const extra = [...(src.extra || []).filter((e) => !PAINT_SKIP_EXTRA.has(e[0])), ...(node.attrs.extra || []).filter((e) => PAINT_SKIP_EXTRA.has(e[0]))];
        tr.setNodeMarkup(pos, null, { ...src, sdt: node.attrs.sdt, deco: node.attrs.deco, extra });
        return false;
      });
    }
    dispatch(tr);
  }
  return true;
};

/** Words, characters and paragraphs in the document (deleted text not counted). */
export function wordCount(doc) {
  let text = '';
  let paragraphs = 0;
  doc.descendants((n) => {
    if (n.type === paragraph) {
      if (n.textContent.trim()) paragraphs++;
      text += '\n';
    }
    if (n.isText && !n.marks.some((m) => m.type.name === 'del')) text += n.text;
    // Text boxes' text too, as Word counts it: the shape's copy only (its VML fallback holds the same text again).
    if (n.type.name === 'image' && /<w:txbxContent\b/.test(n.attrs.xml || '')) {
      const box = /<w:txbxContent\b[^>]*>([\s\S]*?)<\/w:txbxContent>/.exec(n.attrs.xml)[1];
      for (const p of box.split(/<\/w:p>/)) {
        const t = [...p.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('');
        if (t.trim()) { paragraphs++; text += '\n' + unescapeXml(t) + '\n'; }
      }
    }
    return true;
  });
  const words = (text.match(/\S+/g) || []).length;
  return { words, characters: text.replace(/\n/g, '').length, charactersNoSpaces: text.replace(/\s/g, '').length, paragraphs };
}

/** A paragraph's outline level (0 = Heading 1), from its formatting, its style, or a built-in "heading N" style name; null for body text. */
export function outlineLevel(session, node) {
  const lvl = node.attrs.view?.outlineLvl ?? session.styleProp(node.attrs.styleId, 'outlineLvl');
  if (lvl != null) return lvl > 8 ? null : lvl;
  for (let id = node.attrs.styleId, n = 0; id && n < 10; n++) {
    const st = session.styles.styles.get(id);
    // (A built-in heading used without being defined is one too: Heading1 to Heading9.)
    const m = st ? /^heading\s*(\d)$/i.exec(st.name || '') : /^heading([1-9])$/i.exec(id);
    if (m) return Number(m[1]) - 1;
    id = st && st.basedOn;
  }
  return null;
}

/** Apply a list from the galleries (levels as in LIST_LIBRARY) to the selected paragraphs, keeping their levels. */
export const setList = (session, levels) => (state, dispatch) => {
  const paras = selectedParagraphs(state);
  if (!paras.length) return false;
  if (dispatch) {
    const numId = session.listFor(levels);
    const tr = state.tr;
    for (const { node, pos } of paras) tr.setNodeMarkup(pos, null, { ...node.attrs, numId, ilvl: paragraphList(session, node.attrs)?.ilvl ?? node.attrs.ilvl ?? 0 });
    dispatch(tr.scrollIntoView());
  }
  return true;
};

/** attrs with `view` re-derived from the pPr they now write (after changing raw pPr children). */
function reView(attrs) {
  const xml = pPrXml(attrs) || '<w:pPr/>';
  return { ...attrs, view: paragraphAttrsFromPPr(xml).view || {} };
}

const BORDER_SIDES = ['top', 'left', 'bottom', 'right', 'between'];
const bdr = (side) => `<w:${side} w:val="single" w:sz="4" w:space="1" w:color="auto"/>`;

/** Paragraph borders, as Word's Borders menu: 'bottom' | 'top' | 'left' | 'right' | 'none' | 'all' | 'outside' (toggles a single side). */
export const setBorders = (which) => (state, dispatch) => {
  const paras = selectedParagraphs(state);
  if (!paras.length) return false;
  if (!dispatch) return true;
  const first = paras[0].node.attrs;
  const cur = new Set(BORDER_SIDES.filter((sd) => new RegExp(`<w:${sd}\\b[^>]*w:val="(?!none|nil)`).test((first.extra || []).find((e) => e[0] === 'pBdr')?.[1] || '')));
  let sides;
  if (which === 'none') sides = new Set();
  else if (which === 'all') sides = new Set(BORDER_SIDES);
  else if (which === 'outside') sides = new Set(['top', 'left', 'bottom', 'right']);
  else { sides = new Set(cur); sides.has(which) ? sides.delete(which) : sides.add(which); }
  const xml = sides.size ? `<w:pBdr>${BORDER_SIDES.filter((sd) => sides.has(sd)).map(bdr).join('')}</w:pBdr>` : null;
  const tr = state.tr;
  for (const { node, pos } of paras) {
    const rest = (node.attrs.extra || []).filter((e) => e[0] !== 'pBdr');
    tr.setNodeMarkup(pos, null, reView({ ...node.attrs, extra: xml ? [...rest, ['pBdr', xml]] : rest }));
  }
  dispatch(tr.scrollIntoView());
  return true;
};

/** Paragraph shading (hex fill like 'FFFF00'; null removes it). */
export const setShading = (fill) => updateParagraphs((a) => {
  const rest = (a.extra || []).filter((e) => e[0] !== 'shd');
  return reView({ ...a, extra: fill ? [...rest, ['shd', `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`]] : rest });
});

/** w:sectPr with its w:type set ('nextPage' | 'continuous' | 'evenPage' | 'oddPage'). */
function withSectType(xml, kind) {
  let s = (xml || '<w:sectPr/>').replace(/<w:type\b[^>]*\/>/g, '');
  if (/^<w:sectPr\b[^>]*\/>$/.test(s)) s = s.replace(/\/>$/, '></w:sectPr>');
  const t = `<w:type w:val="${kind}"/>`;
  const re = /<w:(?:headerReference|footerReference)\b[^>]*\/>|<w:(?:footnotePr|endnotePr)\b[\s\S]*?<\/w:(?:footnotePr|endnotePr)>/g;
  let at = -1;
  for (const m of s.matchAll(re)) at = m.index + m[0].length;
  if (at < 0) at = s.indexOf('>') + 1;
  return s.slice(0, at) + t + s.slice(at);
}

/** Insert a section break at the cursor: the text after it starts a new section ('nextPage' or 'continuous'). */
export const insertSectionBreak = (kind) => (state, dispatch) => {
  const { $from } = state.selection;
  if ($from.depth !== 1 || $from.parent.type !== paragraph) return false;
  if (!dispatch) return true;
  const doc = state.doc;
  const idx = $from.index(0);
  const sectOf = (n) => n.type === paragraph && (n.attrs.extra || []).find((e) => e[0] === 'sectPr')?.[1];
  let followIdx = -1;
  let follow = doc.attrs.sectPr || '<w:sectPr/>';
  for (let i = idx; i < doc.childCount; i++) {
    const x = sectOf(doc.child(i));
    if (x) { followIdx = i; follow = x; break; }
  }
  const tr = state.tr.deleteSelection();
  const pos = tr.selection.from;
  tr.split(pos);
  // The paragraph before the break ends the new section, with the same page setup as the section it came from.
  const before = tr.doc.resolve(pos).before(1);
  const pb = tr.doc.nodeAt(before);
  tr.setNodeMarkup(before, null, { ...pb.attrs, extra: [...(pb.attrs.extra || []).filter((e) => e[0] !== 'sectPr'), ['sectPr', follow]] });
  // The section after the break starts as asked.
  if (followIdx === idx) {
    const after = tr.mapping.map(pos) ;
    const $a = tr.doc.resolve(after);
    const pa = $a.parent;
    tr.setNodeMarkup($a.before(1), null, { ...pa.attrs, extra: [...pa.attrs.extra.filter((e) => e[0] !== 'sectPr'), ['sectPr', withSectType(follow, kind)]] });
  } else if (followIdx > idx) {
    let p = 0;
    for (let i = 0; i < followIdx + 1; i++) p += tr.doc.child(i).nodeSize;
    const n = tr.doc.child(followIdx + 1);
    tr.setNodeMarkup(p, null, { ...n.attrs, extra: n.attrs.extra.map((e) => (e[0] === 'sectPr' ? ['sectPr', withSectType(e[1], kind)] : e)) });
  } else tr.setDocAttribute('sectPr', withSectType(follow, kind));
  dispatch(tr.scrollIntoView());
  return true;
};

/** Page borders on the last section: 'none' | 'box' | 'double' | 'shadow' | 'thick'. */
export const setPageBorders = (kind) => (state, dispatch) => {
  if (!dispatch) return true;
  let sect = (state.doc.attrs.sectPr || '<w:sectPr></w:sectPr>').replace(/<w:pgBorders\b[\s\S]*?<\/w:pgBorders>|<w:pgBorders\b[^>]*\/>/, '');
  if (/^<w:sectPr\b[^>]*\/>$/.test(sect)) sect = sect.replace(/\/>$/, '></w:sectPr>');
  if (kind !== 'none') {
    const [val, sz, extra] = { box: ['single', 4, ''], double: ['double', 4, ''], thick: ['single', 18, ''], shadow: ['single', 4, ' w:shadow="1"'] }[kind];
    const side = (s) => `<w:${s} w:val="${val}" w:sz="${sz}" w:space="24" w:color="auto"${extra}/>`;
    const xml = `<w:pgBorders w:offsetFrom="page">${['top', 'left', 'bottom', 'right'].map(side).join('')}</w:pgBorders>`;
    const m = /<w:paperSrc\b[^>]*\/>|<w:pgMar\b[^>]*\/>|<w:pgSz\b[^>]*\/>/.exec(sect);
    const at = m ? m.index + m[0].length : sect.indexOf('>') + 1;
    sect = sect.slice(0, at) + xml + sect.slice(at);
  }
  dispatch(state.tr.setDocAttribute('sectPr', sect));
  return true;
};

/**
 * Insert Rows Above / Below, as Word does it: the new row takes the formatting of the row it is made from (its height
 * and other row properties, each cell's shading, borders and vertical alignment, and its first paragraph's style,
 * alignment and paragraph-mark formatting), not a plain row.
 */
export const addRow = (below) => (state, dispatch) => {
  if (!isInTable(state)) return false;
  const rect = selectedRect(state);
  const at = below ? rect.bottom : rect.top; // the new row's index
  const run = below ? addRowAfter : addRowBefore;
  return run(state, dispatch && ((tr) => {
    const tStart = rect.tableStart;
    const table = tr.doc.nodeAt(tStart - 1);
    const map = TableMap.get(table);
    const from = below ? at - 1 : at + 1; // the row it was made from, in the new table
    if (from < 0 || from >= map.height) { dispatch(tr); return; }
    const rowPos = (r) => { let p = 0; for (let i = 0; i < r; i++) p += table.child(i).nodeSize; return p; };
    const src = table.child(from), made = table.child(at);
    const noRev = (xml) => (xml ? xml.replace(/<w:(ins|del)\b[^>]*\/>/g, '').replace(/^<w:trPr\b[^>]*>\s*<\/w:trPr>$/, '') || null : xml);
    if (src.attrs.trPr && !made.attrs.trPr) tr.setNodeMarkup(tStart + rowPos(at), null, { ...made.attrs, trPr: noRev(src.attrs.trPr) });
    const DROP = ['cellIns', 'cellDel', 'cellMerge', 'tcPrChange'];
    for (let c = 0; c < map.width; c++) {
      const pos = map.map[at * map.width + c];
      if (map.findCell(pos).top !== at || map.findCell(pos).left !== c) continue; // spans into the row from above
      const cell = table.nodeAt(pos);
      const like = table.nodeAt(map.map[from * map.width + c]);
      if (!cell || !like) continue;
      tr.setNodeMarkup(tStart + pos, null, { ...cell.attrs, tcPr: (like.attrs.tcPr || []).filter((t) => !DROP.includes(t[0])), colwidth: cell.attrs.colwidth || like.attrs.colwidth });
      const p = like.firstChild;
      if (p && p.type === schema.nodes.paragraph && cell.firstChild?.type === schema.nodes.paragraph) {
        tr.setNodeMarkup(tStart + pos + 1, null, { ...p.attrs, pattrs: null, sdt: null, extra: (p.attrs.extra || []).filter((e) => e[0] !== 'sectPr' && e[0] !== 'pPrChange') });
      }
    }
    dispatch(tr);
  }));
};
