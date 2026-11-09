// Resizing pictures, as in Word: select a picture and drag one of the eight handles around it. Corner handles
// keep its proportions; side handles stretch one way. The new size is one undoable step, saved in the
// picture's own drawing XML (see the writer).
import { Plugin, NodeSelection, TextSelection } from 'prosemirror-state';
import { editTextBox } from './txbxedit.js';
import { editChart, chartXmlAt } from './charts.js';

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const MIN = 8; // px

export function imageResizePlugin() {
  return new Plugin({
    view(view) {
      const box = document.createElement('div');
      box.className = 'odx-imgsize';
      box.setAttribute('aria-hidden', 'true');
      for (const h of HANDLES) {
        const d = document.createElement('div');
        d.className = `odx-imgsize-h odx-imgsize-${h}`;
        d.dataset.h = h;
        box.append(d);
      }
      let target = null; // { pos, node, img }
      const place = () => {
        if (!target || !target.img.isConnected) { box.remove(); return; }
        const host = view.dom.parentElement;
        const r = target.img.getBoundingClientRect();
        const o = host.getBoundingClientRect();
        box.style.cssText = `left:${r.left - o.left}px;top:${r.top - o.top}px;width:${r.width}px;height:${r.height}px`;
        if (box.parentElement !== host) host.append(box);
      };
      const select = () => {
        const sel = view.state.selection;
        const node = sel instanceof NodeSelection && sel.node.type.name === 'image' ? sel.node : null;
        const img = node && view.nodeDOM(sel.from);
        target = node && !view.odx?.readOnly && img && img.tagName === 'IMG' ? { pos: sel.from, node, img } : null;
        box.classList.toggle('odx-imgsize-move', !!target && floating(target.node));
        place();
        if (target) setTimeout(place, 300); // again once pagination has moved things
      };
      box.addEventListener('pointerdown', (e) => {
        const h = e.target.dataset?.h;
        if (!h || !target) return;
        e.preventDefault();
        e.stopPropagation();
        const { img, pos, node } = target;
        const w0 = img.getBoundingClientRect().width;
        const h0 = img.getBoundingClientRect().height;
        const x0 = e.clientX;
        const y0 = e.clientY;
        let w = w0;
        let hh = h0;
        const move = (ev) => {
          const dx = (ev.clientX - x0) * (h.includes('w') ? -1 : h.includes('e') ? 1 : 0);
          const dy = (ev.clientY - y0) * (h.includes('n') ? -1 : h.includes('s') ? 1 : 0);
          if (h.length === 2) {
            // Corners keep the proportions: follow whichever way moved further.
            const k = Math.max((w0 + dx) / w0, (h0 + dy) / h0);
            w = Math.max(MIN, w0 * k);
            hh = Math.max(MIN, h0 * k);
          } else {
            w = Math.max(MIN, w0 + dx);
            hh = Math.max(MIN, h0 + dy);
          }
          img.style.width = `${w}px`;
          img.style.height = `${hh}px`;
          place();
        };
        const up = () => {
          box.removeEventListener('pointermove', move);
          box.removeEventListener('pointerup', up);
          box.releasePointerCapture?.(e.pointerId);
          const width = Math.round(w);
          const height = Math.round(hh);
          if (width === node.attrs.width && height === node.attrs.height) return;
          const tr = view.state.tr.setNodeMarkup(pos, null, sized(view, node, width, height));
          view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
          view.focus();
        };
        box.setPointerCapture?.(e.pointerId);
        box.addEventListener('pointermove', move);
        box.addEventListener('pointerup', up);
      });
      // A floating shape (a text box placed on the page) moves by dragging it.
      box.addEventListener('pointerdown', (e) => {
        if (e.target !== box || e.button !== 0 || !target || !floating(target.node)) return;
        e.preventDefault();
        e.stopPropagation();
        const { img, pos, node } = target;
        const x0 = e.clientX, y0 = e.clientY;
        let dx = 0, dy = 0;
        const move = (ev) => { dx = ev.clientX - x0; dy = ev.clientY - y0; img.style.translate = `${dx}px ${dy}px`; place(); };
        const up = () => {
          box.removeEventListener('pointermove', move);
          box.removeEventListener('pointerup', up);
          box.releasePointerCapture?.(e.pointerId);
          img.style.translate = '';
          if (Math.abs(dx) < 2 && Math.abs(dy) < 2) { place(); return; }
          moveBy(view, pos, node, dx, dy);
        };
        box.setPointerCapture?.(e.pointerId);
        box.addEventListener('pointermove', move);
        box.addEventListener('pointerup', up);
      });
      // The overlay is over the shape while it moves: what the shape itself does on a double-click (edit its text, or
      // a chart's data) and on a right-click (the menu) still happen.
      box.addEventListener('dblclick', (e) => {
        if (e.target !== box || !target) return;
        e.preventDefault();
        const { pos, node } = target;
        if (/txbxContent/.test(node.attrs.xml || '')) editTextBox(view, pos, e.clientX, e.clientY);
        else if (chartXmlAt(view, node)) editChart(view, pos);
      });
      box.addEventListener('contextmenu', (e) => {
        if (e.target !== box || !target) return;
        e.preventDefault();
        target.img.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: e.clientX, clientY: e.clientY, button: 2 }));
      });
      const onScroll = () => place();
      window.addEventListener('resize', onScroll);
      select();
      return {
        update: select,
        destroy() { box.remove(); window.removeEventListener('resize', onScroll); },
      };
    },
  });
}

// A shape (a text box, a drawing): its own frame and text are resized too, the text wrapped again (docx.js
// resizeShape), not the picture of it stretched. A picture just takes the new size.
const isShape = (node) => /<(?:wps|wpg):|txbxContent/.test(node.attrs.xml || '');
function sized(view, node, width, height) {
  const r = isShape(node) && view.odx?.session?.resizeShape(node.attrs.xml, width, height);
  return r ? { ...node.attrs, ...r } : { ...node.attrs, width, height };
}
// (Placed on the page: absolutely, or floated with text wrapping around it; by offsets, which a move changes.)
const floating = (node) => isShape(node) && /position:\s*absolute|float:\s*(left|right)/.test(node.attrs.place || '') && /<wp:posOffset>/.test(node.attrs.xml || '');
function moveBy(view, pos, node, dx, dy) {
  const r = view.odx?.session?.moveShape(node.attrs.xml, dx, dy);
  if (!r) return false;
  const tr = view.state.tr.setNodeMarkup(pos, null, { ...node.attrs, ...r });
  view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
  return true;
}

/**
 * The keyboard's way to move and resize a selected shape (2.5.7, dragging movements): the arrow keys move a floating one
 * 8px (Ctrl: 1px), as Word nudges it; Shift+arrows make any shape wider, narrower, taller or shorter; Esc goes back to
 * the text after it. (Before the keymap, so the arrows aren't taken as caret moves.)
 */
export function shapeKeysPlugin() {
  return new Plugin({
    props: {
      handleKeyDown(view, e) {
        const sel = view.state.selection;
        if (!(sel instanceof NodeSelection) || sel.node.type.name !== 'image' || !isShape(sel.node) || view.odx?.readOnly) return false;
        if (e.key === 'Escape') { view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, sel.to))); return true; }
        const dir = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[e.key];
        if (!dir || e.altKey) return false;
        const step = e.ctrlKey || e.metaKey ? 1 : 8;
        if (e.shiftKey) {
          const { width, height } = sel.node.attrs;
          const w = Math.max(8, Math.round(width + dir[0] * step)), h = Math.max(8, Math.round(height + dir[1] * step));
          const tr = view.state.tr.setNodeMarkup(sel.from, null, sized(view, sel.node, w, h));
          view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, sel.from)));
          return true;
        }
        if (!floating(sel.node)) return false;
        return moveBy(view, sel.from, sel.node, dir[0] * step, dir[1] * step);
      },
    },
  });
}
