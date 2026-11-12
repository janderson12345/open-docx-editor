// Resizing pictures and shapes, as in Word: select one and drag one of the eight handles around it. A picture's corner
// handles keep its proportions; a shape's resize it freely (Shift keeps its proportions, as does a shape locked to
// them); side handles stretch one way. A line has a handle at each end instead, which moves that end (Shift: at a
// multiple of 15 degrees). A shape's yellow handles change its adjust values (an arrow's head, a rounded rectangle's
// corners, a callout's pointer). Each change is one undoable step, saved in the drawing's own XML (see the writer).
import { Plugin, NodeSelection, TextSelection } from 'prosemirror-state';
import { editTextBox } from './txbxedit.js';
import { editChart, chartXmlAt } from './charts.js';
import { zoomOf } from './zoom.js';
import { PRESETS } from '../ooxml/presets.js';
import { adjustHandles, dragAdjust } from '../ooxml/geometry.js';

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const MIN = 8; // px
const LINES = /^(?:line|straightConnector1)$/;

/**
 * What a selected drawing's handles need to know of its shape: a single preset shape (not a group), unrotated, with
 * its definition, own adjust values, size (EMU) and flips. Null for anything else (pictures, groups, charts).
 */
export function shapeInfo(xml) {
  if (!xml || !/<wps:wsp\b/.test(xml) || /<wpg:/.test(xml)) return null;
  const sp = /<wps:spPr\b[^>]*>[\s\S]*?<\/wps:spPr>/.exec(xml)?.[0];
  const xf = sp && /<a:xfrm\b([^>]*)>[\s\S]*?<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(sp);
  const prst = sp && /<a:prstGeom\b[^>]*\bprst="([^"]+)"/.exec(sp)?.[1];
  if (!xf || !prst || !PRESETS[prst] || Number(/\brot="(-?\d+)"/.exec(xf[1])?.[1] || 0) % 21600000) return null;
  const av = /<a:prstGeom\b[^>]*>[\s\S]*?<\/a:prstGeom>/.exec(sp)?.[0] || '';
  const adjust = Object.fromEntries([...av.matchAll(/<a:gd\b[^>]*\bname="([^"]+)"[^>]*\bfmla="([^"]+)"/g)].map((m) => [m[1], m[2]]));
  const on = (k) => new RegExp(`\\b${k}="(?:1|true)"`).test(xf[1]);
  return { prst, def: PRESETS[prst], adjust, cx: Number(xf[2]), cy: Number(xf[3]), flipH: on('flipH'), flipV: on('flipV'), line: LINES.test(prst) };
}

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
      let target = null; // { pos, node, img, info }
      let extra = []; // a line's end handles, a shape's adjust handles
      // The drawing's own box (px), in the page's own px whatever the zoom: its picture less the room drawn around it
      // for arrowheads (--odx-pad).
      const padOf = (node) => Number(/--odx-pad:([\d.]+)px/.exec(node.attrs.place || '')?.[1]) || 0;
      const boxOf = () => {
        const z = zoomOf(view);
        const r = target.img.getBoundingClientRect();
        const pad = padOf(target.node);
        return { left: r.left / z + pad, top: r.top / z + pad, width: Math.max(0, r.width / z - 2 * pad), height: Math.max(0, r.height / z - 2 * pad), z };
      };
      // Where the extra handles go in a w x h box (px), for the shape `info`.
      const spots = (info, w, h) => {
        if (!info) return [];
        const fx = (x) => (info.flipH ? w - x : x), fy = (y) => (info.flipV ? h - y : y);
        if (info.line) return [{ kind: 'end', i: 0, x: fx(0), y: fy(0) }, { kind: 'end', i: 1, x: fx(w), y: fy(h) }];
        if (!info.cx || !info.cy) return [];
        return adjustHandles(info.def, info.cx, info.cy, info.adjust).map((p, i) => ({ kind: 'adj', i, x: fx((p.x / info.cx) * w), y: fy((p.y / info.cy) * h) }));
      };
      const drawExtra = (info, w, h) => {
        const want = spots(info, w, h);
        while (extra.length > want.length) extra.pop().remove();
        want.forEach((s, k) => {
          let d = extra[k];
          if (!d) { d = document.createElement('div'); box.append(d); extra.push(d); }
          d.className = `odx-imgsize-${s.kind}`;
          d.dataset.kind = s.kind;
          d.dataset.i = String(s.i);
          d.style.left = `${s.x}px`;
          d.style.top = `${s.y}px`;
        });
      };
      const place = (info = target?.info, size = null) => {
        if (!target || !target.img.isConnected) { box.remove(); return; }
        const host = view.dom.parentElement;
        const o = host.getBoundingClientRect();
        const b = boxOf();
        const w = size ? size.w : b.width, h = size ? size.h : b.height;
        box.style.cssText = `left:${b.left - o.left / b.z}px;top:${b.top - o.top / b.z}px;width:${w}px;height:${h}px`;
        drawExtra(info, w, h);
        if (box.parentElement !== host) host.append(box);
      };
      const select = () => {
        const sel = view.state.selection;
        const node = sel instanceof NodeSelection && sel.node.type.name === 'image' ? sel.node : null;
        const img = node && view.nodeDOM(sel.from);
        target = node && !view.odx?.readOnly && img && img.tagName === 'IMG' ? { pos: sel.from, node, img, info: isShape(node) ? shapeInfo(node.attrs.xml) : null } : null;
        box.classList.toggle('odx-imgsize-move', !!target && floating(target.node));
        box.classList.toggle('odx-imgsize-line', !!target?.info?.line);
        if (target) target.img.classList.add('odx-sized');
        place();
        if (target) setTimeout(place, 300); // again once pagination has moved things
      };
      // The drawing redrawn while a handle is dragged (its picture and size), and the step that keeps the result.
      const preview = (r) => {
        const { img, node } = target;
        const pad = padOf({ attrs: { place: r.place ?? node.attrs.place } });
        if (r.src) img.src = r.src;
        img.style.width = `${r.width + 2 * pad}px`;
        img.style.height = `${r.height + 2 * pad}px`;
      };
      const commit = (pos, node, r) => {
        const tr = view.state.tr.setNodeMarkup(pos, null, { ...node.attrs, ...r });
        view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
        view.focus();
      };
      // Dragging a handle: listened for on the overlay, which holds the pointer until the button is let go.
      const drag = (e, onMove, onUp) => {
        e.preventDefault();
        e.stopPropagation();
        const move = (ev) => onMove(ev);
        const up = (ev) => {
          box.removeEventListener('pointermove', move);
          box.removeEventListener('pointerup', up);
          box.releasePointerCapture?.(e.pointerId);
          onUp(ev);
        };
        box.setPointerCapture?.(e.pointerId);
        box.addEventListener('pointermove', move);
        box.addEventListener('pointerup', up);
      };
      box.addEventListener('pointerdown', (e) => {
        const h = e.target.dataset?.h;
        if (!h || !target) return;
        const { img, pos, node, info } = target;
        // (In the page's own px: the screen's divided by the zoom.)
        const b0 = boxOf();
        const z = b0.z;
        const w0 = b0.width;
        const h0 = b0.height;
        const x0 = e.clientX;
        const y0 = e.clientY;
        const pad = padOf(node);
        // Proportions kept: a picture's always (Word locks them), a shape's when it's locked to them or with Shift.
        const locked = !isShape(node) || /\bnoChangeAspect="(?:1|true)"/.test(node.attrs.xml || '');
        let w = w0;
        let hh = h0;
        drag(e, (ev) => {
          const dx = ((ev.clientX - x0) / z) * (h.includes('w') ? -1 : h.includes('e') ? 1 : 0);
          const dy = ((ev.clientY - y0) / z) * (h.includes('n') ? -1 : h.includes('s') ? 1 : 0);
          if (h.length === 2 && (locked || ev.shiftKey) && w0 && h0) {
            // Proportions kept: follow whichever way moved further.
            const k = Math.max((w0 + dx) / w0, (h0 + dy) / h0);
            w = Math.max(MIN, w0 * k);
            hh = Math.max(MIN, h0 * k);
          } else {
            w = Math.max(MIN, w0 + dx);
            hh = Math.max(MIN, h0 + dy);
          }
          img.style.width = `${w + 2 * pad}px`;
          img.style.height = `${hh + 2 * pad}px`;
          place(info ? { ...info, cx: Math.round(w * 9525), cy: Math.round(hh * 9525) } : null, { w, h: hh });
        }, () => {
          const width = Math.round(w);
          const height = Math.round(hh);
          if (width === node.attrs.width && height === node.attrs.height) { img.style.width = ''; img.style.height = ''; place(); return; }
          commit(pos, node, sized(view, node, width, height));
        });
      });
      // A line's end: dragged anywhere, the other end staying put. The line's box is the two ends' span, flipped to
      // run from its start to its end; a floating line moves with the box's corner.
      box.addEventListener('pointerdown', (e) => {
        if (e.target.dataset?.kind !== 'end' || !target?.info) return;
        const { img, pos, node, info } = target;
        const session = view.odx?.session;
        if (!session) return;
        const b0 = boxOf();
        const z = b0.z;
        const ends = spots(info, b0.width, b0.height); // [start, end] in the box (px)
        const moving = Number(e.target.dataset.i);
        const fixed = ends[1 - moving];
        const grab = ends[moving];
        const x0 = e.clientX, y0 = e.clientY;
        let result = null;
        let frame = 0;
        drag(e, (ev) => {
          let px = grab.x + (ev.clientX - x0) / z, py = grab.y + (ev.clientY - y0) / z;
          if (ev.shiftKey) {
            // At a multiple of 15 degrees from the other end, as long as it was.
            const len = Math.hypot(px - fixed.x, py - fixed.y);
            const ang = Math.round(Math.atan2(py - fixed.y, px - fixed.x) / (Math.PI / 12)) * (Math.PI / 12);
            px = fixed.x + len * Math.cos(ang);
            py = fixed.y + len * Math.sin(ang);
          }
          const [start, end] = moving === 0 ? [{ x: px, y: py }, fixed] : [fixed, { x: px, y: py }];
          const left = Math.min(start.x, end.x), top = Math.min(start.y, end.y);
          const width = Math.max(1, Math.round(Math.abs(end.x - start.x))), height = Math.max(1, Math.round(Math.abs(end.y - start.y)));
          cancelAnimationFrame(frame);
          frame = requestAnimationFrame(() => {
            result = session.lineShape(node.attrs.xml, width, height, start.x > end.x, start.y > end.y, floating(node) ? left : 0, floating(node) ? top : 0);
            if (!result) return;
            preview(result);
            if (floating(node)) img.style.translate = `${left}px ${top}px`;
            place({ ...info, flipH: start.x > end.x, flipV: start.y > end.y }, { w: width, h: height });
            box.style.translate = img.style.translate;
          });
        }, () => {
          cancelAnimationFrame(frame);
          img.style.translate = '';
          box.style.translate = '';
          if (!result) { place(); return; }
          commit(pos, node, result);
        });
      });
      // A yellow handle: its adjust values found for where it's dragged, the shape redrawn as it goes.
      box.addEventListener('pointerdown', (e) => {
        if (e.target.dataset?.kind !== 'adj' || !target?.info) return;
        const { pos, node, info } = target;
        const session = view.odx?.session;
        if (!session) return;
        const b0 = boxOf();
        const z = b0.z;
        const i = Number(e.target.dataset.i);
        const hb = e.target.getBoundingClientRect();
        // (From the handle's centre, wherever on it the pointer went down.)
        const gx = (hb.left + hb.width / 2) / z - b0.left, gy = (hb.top + hb.height / 2) / z - b0.top;
        const x0 = e.clientX, y0 = e.clientY;
        let result = null;
        let adjust = info.adjust;
        let frame = 0;
        drag(e, (ev) => {
          const bx = gx + (ev.clientX - x0) / z, by = gy + (ev.clientY - y0) / z;
          // (In the shape's own units, unflipped.)
          const ux = (info.flipH ? b0.width - bx : bx) / (b0.width || 1) * info.cx;
          const uy = (info.flipV ? b0.height - by : by) / (b0.height || 1) * info.cy;
          cancelAnimationFrame(frame);
          frame = requestAnimationFrame(() => {
            const values = dragAdjust(info.def, info.cx, info.cy, info.adjust, i, ux, uy);
            result = session.adjustShape(node.attrs.xml, values);
            if (!result) return;
            adjust = { ...info.adjust, ...values };
            preview(result);
            place({ ...info, adjust });
          });
        }, () => {
          cancelAnimationFrame(frame);
          if (!result) { place(); return; }
          commit(pos, node, result);
        });
      });
      // A floating shape (a text box placed on the page) moves by dragging it.
      box.addEventListener('pointerdown', (e) => {
        if (e.target !== box || e.button !== 0 || !target || !floating(target.node)) return;
        e.preventDefault();
        e.stopPropagation();
        const { img, pos, node } = target;
        const x0 = e.clientX, y0 = e.clientY;
        const z = zoomOf(view);
        let dx = 0, dy = 0;
        const move = (ev) => { dx = (ev.clientX - x0) / z; dy = (ev.clientY - y0) / z; img.style.translate = `${dx}px ${dy}px`; place(); };
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
