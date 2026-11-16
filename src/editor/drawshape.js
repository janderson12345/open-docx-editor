// Insert > Shapes with the mouse, as in Word: the pointer turns into a crosshair, and the shape is drawn by dragging
// on the page, from one corner to the other (Shift: a square, a circle, or a line at 15-degree steps; Ctrl: from its
// centre out). A click without a drag puts the shape there at Word's own size. The shape floats in front of the text
// where it was drawn, anchored to the paragraph it starts beside, as Word's are. Esc, or a press off the page, stops.
import { NodeSelection, Selection } from 'prosemirror-state';
import { closeHistory } from 'prosemirror-history';
import { schema } from './schema.js';
import { PRESETS } from '../ooxml/presets.js';
import { geometryPaths } from '../ooxml/geometry.js';
import { zoomOf } from './zoom.js';
import { padOf } from '../ooxml/place.js';

const SVG = 'http://www.w3.org/2000/svg';
// (A click's size: Word's click-to-insert, 1in by 0.75in; a line 1.5in across.)
const CLICK = { w: 96, h: 72, line: 144 };

/** The box (client px) dragged out from (x0, y0) to (x1, y1), with Shift and Ctrl as Word takes them. */
export function dragBox(x0, y0, x1, y1, { line = false, shift = false, ctrl = false } = {}) {
  let dx = x1 - x0, dy = y1 - y0;
  if (shift && line) {
    // (At 15-degree steps from where it starts.)
    const len = Math.hypot(dx, dy);
    const a = Math.round(Math.atan2(dy, dx) / (Math.PI / 12)) * (Math.PI / 12);
    dx = Math.round(len * Math.cos(a) * 100) / 100;
    dy = Math.round(len * Math.sin(a) * 100) / 100;
  } else if (shift) {
    const s = Math.max(Math.abs(dx), Math.abs(dy));
    dx = Math.sign(dx || 1) * s;
    dy = Math.sign(dy || 1) * s;
  }
  if (ctrl && !line) { x0 -= dx; y0 -= dy; dx *= 2; dy *= 2; }
  return { left: Math.min(x0, x0 + dx), top: Math.min(y0, y0 + dy), width: Math.abs(dx), height: Math.abs(dy), flipH: dx < 0, flipV: dy < 0 };
}

/**
 * Starts drawing the preset shape `prst` (named `name`; `arrow`: a line with an arrowhead) in `view`. `done(ok)` is
 * called once it's drawn (true) or given up (false). Returns a function that stops it.
 */
export function drawShape(view, { prst, name, arrow = false }, done = () => {}) {
  const line = prst === 'line';
  const layer = document.createElement('div');
  layer.className = 'odx-draw';
  layer.setAttribute('aria-hidden', 'true');
  const svg = document.createElementNS(SVG, 'svg');
  svg.setAttribute('class', 'odx-draw-shape');
  layer.append(svg);
  document.body.append(layer);
  let start = null;
  let finished = false;
  const stop = (ok) => {
    if (finished) return;
    finished = true;
    layer.remove();
    window.removeEventListener('keydown', onKey, true);
    done(ok);
  };
  const onKey = (e) => {
    if (e.key !== 'Escape') return;
    e.preventDefault();
    e.stopPropagation();
    stop(false);
  };
  window.addEventListener('keydown', onKey, true);
  // (The pages: where a shape can be drawn, below the last line too, as far down as the page is drawn.)
  const onPage = (x, y) => {
    const r = view.dom.getBoundingClientRect();
    const host = view.dom.parentElement?.getBoundingClientRect() || r;
    return x >= r.left && x <= r.right && y >= r.top && y <= Math.max(r.bottom, host.bottom);
  };
  const show = (b) => {
    const w = Math.max(1, b.width), h = Math.max(1, b.height);
    svg.style.cssText = `left:${b.left - 1}px;top:${b.top - 1}px;width:${w + 2}px;height:${h + 2}px`;
    svg.setAttribute('viewBox', `-1 -1 ${w + 2} ${h + 2}`);
    svg.replaceChildren();
    const ds = line
      ? [`M${b.flipH ? w : 0} ${b.flipV ? h : 0}L${b.flipH ? 0 : w} ${b.flipV ? 0 : h}`]
      : geometryPaths(PRESETS[prst], w, h).map((p) => p.d);
    for (const d of ds) {
      const p = document.createElementNS(SVG, 'path');
      p.setAttribute('d', d);
      svg.append(p);
    }
  };
  layer.addEventListener('pointerdown', (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    if (!onPage(e.clientX, e.clientY)) { stop(false); return; }
    start = { x: e.clientX, y: e.clientY };
    layer.setPointerCapture?.(e.pointerId);
  });
  layer.addEventListener('pointermove', (e) => {
    if (!start) return;
    show(dragBox(start.x, start.y, e.clientX, e.clientY, { line, shift: e.shiftKey, ctrl: e.ctrlKey || e.metaKey }));
  });
  layer.addEventListener('pointerup', (e) => {
    if (!start) return;
    layer.releasePointerCapture?.(e.pointerId);
    const z = zoomOf(view);
    let b = dragBox(start.x, start.y, e.clientX, e.clientY, { line, shift: e.shiftKey, ctrl: e.ctrlKey || e.metaKey });
    // (A click, not a drag: the shape at Word's size, its top left corner there.)
    if (Math.hypot(e.clientX - start.x, e.clientY - start.y) < 3) {
      b = { left: start.x, top: start.y, width: (line ? CLICK.line : CLICK.w) * z, height: line ? 0 : CLICK.h * z, flipH: false, flipV: false };
    }
    // (The layer off first: the text under the corner is found by what's at that point on the page.)
    layer.remove();
    stop(place(view, { prst, name, arrow, line }, b, z));
  });
  return () => stop(false);
}

/** The shape drawn over client box `b` (zoom `z`): inserted floating, anchored to the paragraph it starts beside. */
function place(view, { prst, name, arrow, line }, b, z) {
  const session = view.odx?.session;
  if (!session) return false;
  // Its paragraph: the one beside its top left corner (above or below the text, the nearest).
  const r = view.dom.getBoundingClientRect();
  const at = view.posAtCoords({ left: Math.min(Math.max(b.left, r.left + 1), r.right - 1), top: Math.min(Math.max(b.top, r.top + 1), r.bottom - 1) });
  if (!at) return false;
  // (Below the last line, or between blocks: the nearest text, before it.)
  const $p = Selection.near(view.state.doc.resolve(at.pos), -1).$from;
  let d = $p.depth;
  while (d > 0 && !$p.node(d).isTextblock) d--;
  if (!$p.node(d).isTextblock) return false;
  const pos = $p.start(d);
  const dom = view.nodeDOM($p.before(d));
  const para = dom?.nodeType === 1 ? dom : view.dom;
  const left = parseFloat(getComputedStyle(para).getPropertyValue('--odx-page-left')) || 96;
  const top = view.coordsAtPos(pos).top;
  const w = Math.max(line ? 0 : 1, Math.round(b.width / z)), h = Math.max(line ? 0 : 1, Math.round(b.height / z));
  const want = { x: (b.left - r.left) / z - left, y: (b.top - top) / z };
  const made = session.addShape(prst, Math.max(1, w), Math.max(1, h), { arrow, float: { x: Math.round(want.x), y: Math.round(want.y) }, flipH: b.flipH, flipV: b.flipV });
  if (!made) return false;
  const node = schema.nodes.image.create({ ...made, alt: name });
  // (A step of its own for Undo, as each shape drawn is in Word; the fit below joins it.)
  const tr = closeHistory(view.state.tr).insert(pos, node);
  view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)).scrollIntoView());
  view.focus();
  // Then put it exactly where it was drawn: the offsets measured from where the page shows its paragraph (one step with
  // the insertion, for Undo: the same place, straight after).
  const fit = () => {
    const n = view.state.doc.nodeAt(pos);
    const img = view.nodeDOM(pos);
    if (!n || n.attrs.xml !== made.xml || !img?.getBoundingClientRect) return;
    const pad = padOf(n.attrs.place);
    const q = img.getBoundingClientRect();
    const dx = Math.round((b.left - q.left) / z - pad), dy = Math.round((b.top - q.top) / z - pad);
    if (!dx && !dy) return;
    const moved = session.moveShape(n.attrs.xml, dx, dy);
    if (!moved) return;
    const t2 = view.state.tr.setNodeMarkup(pos, null, { ...n.attrs, ...moved });
    view.dispatch(t2.setSelection(NodeSelection.create(t2.doc, pos)));
  };
  requestAnimationFrame(() => requestAnimationFrame(fit));
  return true;
}
