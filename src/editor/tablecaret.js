// The arrow keys at a table's edges, as in Word: Left and Right go on from a cell's end into the next cell, and from the
// paragraph under or over a table into its nearest cell, in right-to-left text as in left-to-right; Up and Down go to
// the cell above or below, or out of the table, keeping the caret's column (Word's tables measured with real keys in
// Word, rtl-caret.docx: Down from the second cell to the second row's caret under it, not to that cell's start).
// The table editing's own arrows, and the browser's, are kept everywhere else.
//
// (Why: the table editing's Left and Right find a cell's edge on screen but then step in the document's order, so in a
// right-to-left cell Left at the cell's end stepped one character back, the browser's Left stepped it forward again,
// and the caret never left the cell; the browser's own caret, moving back from the paragraph under a table, went over
// the whole table; and its Up and Down put the caret at the start of the next row's cell.)
import { Selection, TextSelection } from 'prosemirror-state';
import { TableMap } from 'prosemirror-tables';

const isCell = (n) => n.type.spec.tableRole === 'cell' || n.type.spec.tableRole === 'header_cell';
// The depth of the cell the caret's textblock is in, when it is the cell's first (dir < 0) or last (dir > 0) block.
const edgeCell = ($pos, dir) => {
  for (let d = $pos.depth - 1; d > 0; d--) {
    if (isCell($pos.node(d))) return d;
    if ((dir < 0 ? $pos.index(d) : $pos.indexAfter(d)) !== (dir < 0 ? 0 : $pos.node(d).childCount)) return null;
  }
  return null;
};
const blockDir = (view, $pos) => {
  const dom = view.nodeDOM($pos.before());
  return dom?.nodeType === 1 && getComputedStyle(dom).direction === 'rtl' ? 'rtl' : 'ltr';
};

// The next character boundary (a grapheme's: an Arabic letter with its marks is one) from the caret in the textblock.
const graphemeStep = ($head, dir) => {
  const text = $head.parent.textBetween(0, $head.parent.content.size, undefined, '￼'); // (a leaf one character, as its size)
  const at = $head.parentOffset;
  let to = dir > 0 ? text.length : 0;
  if (typeof Intl !== 'undefined' && Intl.Segmenter) {
    for (const { index } of new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text)) {
      if (dir > 0 && index > at) { to = index; break; }
      if (dir < 0) { if (index < at) to = index; else break; }
    }
  } else to = Math.max(0, Math.min(text.length, at + dir));
  return $head.start() + to;
};

/** Left (dir -1) or Right (1) from a textblock's edge into, out of, or across a table's cells. */
export const tableArrow = (key) => (state, dispatch, view) => {
  if (!view) return false;
  const sel = state.selection;
  if (!(sel instanceof TextSelection) || !sel.empty || !sel.$head.parent.isTextblock) return false;
  const $head = sel.$head;
  const dir = (key < 0) === (blockDir(view, $head) === 'rtl') ? 1 : -1; // (in the document's order)
  if (!view.endOfTextblock(dir > 0 ? 'forward' : 'backward')) {
    // (Inside a cell where its text's edge on screen is not its end in the document's order, as left-to-right words in
    // a right-to-left cell: the table editing would take it for the cell's edge. One character on, as the browser's.)
    if (edgeCell($head, -1) == null && edgeCell($head, 1) == null) return false;
    if (!view.endOfTextblock(key < 0 ? 'left' : 'right')) return false;
    if (dispatch) dispatch(state.tr.setSelection(TextSelection.create(state.doc, graphemeStep($head, dir))).scrollIntoView());
    return true;
  }
  const inCell = edgeCell($head, dir) != null;
  const $side = state.doc.resolve(dir > 0 ? $head.after() : $head.before());
  // (Out of a paragraph, only when the next block that way is a table.)
  if (!inCell && (dir > 0 ? $side.nodeAfter : $side.nodeBefore)?.type.spec.tableRole !== 'table') return false;
  const to = Selection.findFrom($side, dir, true);
  if (!to) return inCell; // (the document's first or last cell: the caret stays)
  if (dispatch) dispatch(state.tr.setSelection(to).scrollIntoView());
  return true;
};

// The caret's column, kept over a run of Up and Down: { x, head } with head where the last move put it.
let goal = null;

// The caret's rectangle at pos as the browser draws it: a collapsed range in the text there. (The view's own measure
// takes the wrong side at the edges of right-to-left text: a cell's start in a right-to-left cell measured at its left.)
const caretRect = (view, pos) => {
  let { node, offset } = view.domAtPos(pos);
  while (node.nodeType === 1 && node.childNodes.length) {
    if (offset < node.childNodes.length) { node = node.childNodes[offset]; offset = 0; } else { node = node.childNodes[offset - 1]; offset = node.nodeType === 3 ? node.length : node.childNodes.length; }
  }
  if (node.nodeType === 3) {
    const r = document.createRange();
    r.setStart(node, offset); r.collapse(true);
    const rect = r.getClientRects()[0];
    if (rect) return rect;
  }
  return view.coordsAtPos(pos);
};
// The caret's x where the browser shows it (at a line's end or a change of direction, the side it is drawn on).
const caretAt = (view) => {
  const s = view.root.getSelection?.();
  if (s?.rangeCount && s.isCollapsed && view.dom.contains(s.anchorNode) && s.anchorNode.nodeType === 3) {
    const rect = s.getRangeAt(0).getClientRects()[0];
    if (rect) return rect.left;
  }
  return caretRect(view, view.state.selection.head).left;
};
// The place on a textblock's first (dir > 0) or last line nearest x.
const nearestOnLine = (view, start, end, dir, x) => {
  let best = dir > 0 ? start : end;
  if (end - start > 4000) return best;
  const dom = view.nodeDOM(start - 1);
  const lines = dom?.nodeType === 1 ? (() => { const r = document.createRange(); r.selectNodeContents(dom); return [...r.getClientRects()].filter((q) => q.height > 0); })() : [];
  const line = dir > 0 ? lines.reduce((a, q) => (!a || q.top < a.top ? q : a), null) : lines.reduce((a, q) => (!a || q.bottom > a.bottom ? q : a), null);
  if (!line) return best;
  const mid = (line.top + line.bottom) / 2;
  let gap = Infinity;
  for (let p = start; p <= end; p++) {
    let r;
    try { r = caretRect(view, p); } catch { continue; }
    if (r.top > mid || r.bottom < mid) continue;
    const g = Math.abs(r.left - x);
    if (g < gap) { gap = g; best = p; }
  }
  return best;
};
// The textblock at the top (dir > 0) or bottom of a cell or other node at pos.
const blockIn = (doc, pos, dir) => {
  const node = doc.nodeAt(pos);
  const sel = dir > 0 ? Selection.findFrom(doc.resolve(pos + 1), 1, true) : Selection.findFrom(doc.resolve(pos + node.nodeSize - 1), -1, true);
  return sel && sel.$head.parent.isTextblock ? sel.$head : null;
};
// The cell of a table's first (dir > 0) or last row under x on screen.
const cellUnder = (view, tablePos, dir, x) => {
  const table = view.state.doc.nodeAt(tablePos);
  const map = TableMap.get(table);
  const row = dir > 0 ? 0 : map.height - 1;
  let best = null, gap = Infinity;
  for (let c = 0; c < map.width; c++) {
    const pos = tablePos + 1 + map.map[row * map.width + c];
    const r = view.nodeDOM(pos)?.getBoundingClientRect?.();
    if (!r) continue;
    const g = x < r.left ? r.left - x : x > r.right ? x - r.right : 0;
    if (g < gap) { gap = g; best = pos; }
  }
  return best;
};

/** Up (dir -1) or Down (1) from a cell's first or last line, or into a table from the line over or under it. */
export const tableVert = (dir) => (state, dispatch, view) => {
  if (!view) return false;
  const sel = state.selection;
  if (!(sel instanceof TextSelection) || !sel.empty || !sel.$head.parent.isTextblock) return false;
  const $head = sel.$head;
  if (!view.endOfTextblock(dir > 0 ? 'down' : 'up')) return false;
  const d = edgeCell($head, dir);
  let target = null; // the textblock's $pos to go to
  let x;
  try { x = goal && goal.head === $head.pos ? goal.x : caretAt(view); } catch { return false; }
  if (d != null) {
    const cellPos = $head.before(d);
    const table = $head.node(d - 2), tableStart = $head.start(d - 2);
    const map = TableMap.get(table);
    const next = map.nextCell(cellPos - tableStart, 'vert', dir);
    if (next != null) {
      // (The cell under x in that row, for cells that span columns.)
      const rect = map.findCell(next);
      const row = dir > 0 ? rect.top : rect.bottom - 1;
      let best = tableStart + next, gap = Infinity;
      for (let c = 0; c < map.width; c++) {
        const pos = tableStart + map.map[row * map.width + c];
        const r = view.nodeDOM(pos)?.getBoundingClientRect?.();
        if (!r) continue;
        const g = x < r.left ? r.left - x : x > r.right ? x - r.right : 0;
        if (g < gap) { gap = g; best = pos; }
      }
      target = blockIn(state.doc, best, dir);
    } else {
      const tablePos = $head.before(d - 2);
      const $out = state.doc.resolve(dir > 0 ? tablePos + table.nodeSize : tablePos);
      const out = Selection.findFrom($out, dir, true);
      if (!out) return true; // (the document's first or last table: the caret stays)
      target = out.$head;
    }
  } else {
    const $side = state.doc.resolve(dir > 0 ? $head.after() : $head.before());
    const next = dir > 0 ? $side.nodeAfter : $side.nodeBefore;
    if (next?.type.spec.tableRole !== 'table') return false;
    const cell = cellUnder(view, dir > 0 ? $side.pos : $side.pos - next.nodeSize, dir, x);
    if (cell == null) return false;
    target = blockIn(state.doc, cell, dir);
  }
  if (!target) return false;
  let pos;
  try { pos = nearestOnLine(view, target.start(), target.end(), dir, x); } catch { return false; }
  if (dispatch) {
    dispatch(state.tr.setSelection(TextSelection.create(state.doc, pos)).scrollIntoView());
    goal = { x, head: pos };
  }
  return true;
};
