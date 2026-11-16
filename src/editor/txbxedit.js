// Editing the text in a text box, as in Word: double-click the box and its text opens in a small
// editor over it (in a group of shapes, the box double-clicked). Clicking outside, Escape or Done puts the edited text back into the shape (and its
// VML fallback copy), redraws it and makes the change one undoable step in the document.
import { textSpansPlugin } from './textspans.js';
import { Plugin, EditorState, NodeSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo, redo } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import { baseKeymap } from 'prosemirror-commands';
import * as cmd from './commands.js';
import { csFormatPlugin } from './csformat.js';
import { zoomOf } from './zoom.js';

let open = null; // { main, tv, el, pos, node, k, doc0, outside }

export const editingTextBox = () => open;

/** Put the edited text back and close the editor. */
export function closeTextBox() {
  if (!open) return;
  const { main, tv, el, pos, node, k, doc0, outside } = open;
  open = null;
  main.odx.hfView = null;
  document.removeEventListener('mousedown', outside, true);
  const doc = tv.state.doc;
  // (A box as wide as its text: its new width, the widest line's and the insets, as Word fits it.)
  let fit = null;
  if (fitsText(node.attrs.xml)) {
    const z = zoomOf(main);
    const cs = getComputedStyle(el);
    let w = 0;
    for (const p of tv.dom.children) {
      const r = document.createRange();
      r.selectNodeContents(p);
      w = Math.max(w, r.getBoundingClientRect().width / z);
    }
    // (With room to spare: Word wraps a line that only just fits, 1pt over in its print. No wider than the column, where
    // Word wraps its lines: shapes.js draws them so.)
    fit = Math.min(colWidth(main), Math.ceil(w * 1.02 + (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.paddingRight) || 0) + 4));
  }
  tv.destroy();
  el.remove();
  if (!doc.eq(doc0) && main.state.doc.nodeAt(pos) === node) {
    let { xml, src } = main.odx.session.setTextBox(node.attrs.xml, doc, k);
    let size = {};
    if (fit && Math.abs(fit - node.attrs.width) >= 1) {
      const r = main.odx.session.resizeShape(xml, fit, node.attrs.height);
      if (r) ({ xml, src, ...size } = { ...r, src: r.src });
      // (About its centre, so it stays where it was and on the page: Word wraps text that would pass the page's edge.)
      const moved = r && main.odx.session.moveShape(xml, -Math.round((fit - node.attrs.width) / 2), 0);
      if (moved) ({ xml, src, ...size } = { ...moved, src: moved.src });
    }
    const tr = main.state.tr.setNodeMarkup(pos, null, { ...node.attrs, ...size, xml, src: src || node.attrs.src });
    main.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
  }
  main.focus();
}

/**
 * Open a text box of the drawing at `pos`: box `k` when given, else the one under the point (x, y) in client px when
 * given (a group can hold several), else the first. Returns its view, or null.
 */
export function editTextBox(main, pos, x = null, y = null, k0 = null) {
  closeTextBox();
  const node = main.state.doc.nodeAt(pos);
  const img = main.nodeDOM(pos);
  if (!node || node.type.name !== 'image' || !/txbxContent/.test(node.attrs.xml || '') || !img?.getBoundingClientRect) return null;
  const boxes = main.odx.session.textBoxes(node.attrs.xml);
  if (!boxes?.length) return null;
  const r = img.getBoundingClientRect();
  const at = (b) => ({ left: r.left + b.rect[0] * r.width, top: r.top + b.rect[1] * r.height, width: b.rect[2] * r.width, height: b.rect[3] * r.height });
  const inside = (b) => { const a = at(b); return x >= a.left && x <= a.left + a.width && y >= a.top && y <= a.top + a.height; };
  // The smallest box under the point, so a box inside another one wins.
  const under = x == null ? [] : boxes.map((b, k) => ({ b, k })).filter(({ b }) => inside(b)).sort((p, q) => p.b.rect[2] * p.b.rect[3] - q.b.rect[2] * q.b.rect[3]);
  if (x != null && !under.length && boxes.length > 1) return null;
  const { b: box, k } = k0 != null && boxes[k0] ? { b: boxes[k0], k: k0 } : under[0] || { b: boxes[0], k: 0 };
  const doc0 = main.odx.session.textBox(node.attrs.xml, k);
  if (!doc0) return null;
  const host = main.dom.parentElement;
  const o = host.getBoundingClientRect();
  const a = at(box);
  const [t, rr, b, l] = box.ins;
  const el = document.createElement('div');
  el.className = 'odx-hfedit odx-txbxedit';
  const z = zoomOf(main); // (in the page's own px, whatever the zoom)
  el.style.cssText = `left:${(a.left - o.left) / z}px;top:${(a.top - o.top) / z}px;width:${a.width / z}px;min-height:${a.height / z}px;padding:${t}px ${rr}px ${b}px ${l}px;box-sizing:border-box`;
  // A box as wide as its text (WordArt: wrap="none", resized to fit): the box widening as it's typed, up to the column's
  // width, where its lines wrap, as Word wraps them.
  if (fitsText(node.attrs.xml)) el.style.cssText += `;width:max-content;min-width:${Math.min(a.width / z, colWidth(main))}px;max-width:${colWidth(main)}px;white-space:pre-wrap;overflow-wrap:anywhere`;
  const tag = document.createElement('div');
  tag.className = 'odx-hfedit-tag';
  const done = document.createElement('button');
  done.type = 'button';
  done.textContent = 'Done';
  done.onmousedown = (e) => { e.preventDefault(); closeTextBox(); };
  done.onclick = (e) => { if (e.detail === 0) closeTextBox(); };
  tag.append('Text Box', done);
  el.append(tag);
  host.append(el);

  const keys = keymap({
    'Mod-z': undo, 'Mod-y': redo, 'Mod-Shift-z': redo,
    'Mod-b': cmd.toggle('bold'), 'Mod-i': cmd.toggle('italic'), 'Mod-u': cmd.toggle('underline'),
    'Mod-e': cmd.setAlign('center'), 'Mod-l': cmd.setAlign('left'), 'Mod-r': cmd.setAlign('right'),
    Tab: cmd.insertTab,
    Escape: () => { closeTextBox(); return true; },
  });
  const tv = new EditorView(el, {
    state: EditorState.create({ doc: doc0, plugins: [history(), keys, keymap(baseKeymap), textSpansPlugin, csFormatPlugin] }),
    attributes: { class: 'odx-hf odx-hfedit-content', spellcheck: 'true' },
    dispatchTransaction(tr) {
      tv.updateState(tv.state.apply(tr));
      main.odx.notify?.(tv);
    },
  });
  tv.odx = { ...main.odx, isStory: true, main };
  main.odx.hfView = tv;
  // A click anywhere else (but the toolbar, which formats the box's text) closes it, as in Word.
  const outside = (e) => {
    if (el.contains(e.target) || e.target.closest?.('.odx-toolbar, .odx-pop')) return;
    closeTextBox();
  };
  document.addEventListener('mousedown', outside, true);
  open = { main, tv, el, pos, node, k, doc0, outside };
  tv.focus();
  return tv;
}

/** The body's column width (px, the page's less its margins), which WordArt's box is no wider than. */
function colWidth(main) {
  const cs = getComputedStyle(main.dom);
  return main.dom.clientWidth - (parseFloat(cs.getPropertyValue('--odx-page-left')) || 96) - (parseFloat(cs.getPropertyValue('--odx-page-right')) || 96);
}

// A text box as wide as its text: not wrapped (wrap="none"), resized to fit it (spAutoFit), as WordArt's.
const fitsText = (xml) => /<wps:bodyPr\b[^>]*\bwrap="none"[\s\S]*?<a:spAutoFit\/>/.test(xml || '');

/** Double-clicking a text box edits its text. */
export function textBoxPlugin() {
  return new Plugin({
    props: {
      handleDoubleClickOn(view, pos, node, nodePos, event) {
        if (view.odx?.readOnly || node.type.name !== 'image' || !/txbxContent/.test(node.attrs.xml || '')) return false;
        return !!editTextBox(view, nodePos, event.clientX, event.clientY);
      },
    },
    view: () => ({ destroy: closeTextBox }),
  });
}
