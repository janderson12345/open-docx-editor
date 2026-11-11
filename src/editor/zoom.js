// Zoom, as Word's View > Zoom: the pages drawn larger or smaller (CSS zoom on the page), at a percentage or fitted to
// the window's width (Page Width) or to show a whole page (One Page), fitted again as the window resizes. Only what's
// shown changes: the layout measures the pages at 100% (pages.js unzoomed), so lines and page breaks stay as they
// are, and printing and PDF export are at full size.
import { announce } from './live.js';

export const ZOOM_STEPS = [0.1, 0.25, 0.5, 0.75, 1, 1.25, 1.5, 2, 3, 4, 5];
const MIN = 0.1, MAX = 5;

/**
 * The zoom of one editor: `mount` the scrolling canvas, `sheet` the page column, `page` its size in px.
 * `onChange(zoom)` after each change. Returns { value, mode, set(value | 'pageWidth' | 'onePage'), step(dir), destroy() }.
 */
export function createZoom(mount, sheet, page, onChange = () => {}) {
  let value = 1;
  let mode = null; // 'pageWidth' | 'onePage' | null (a percentage)
  const fitted = (m) => {
    const cs = getComputedStyle(mount);
    const w = mount.clientWidth - (parseFloat(cs.paddingLeft) || 0) - (parseFloat(cs.paddingRight) || 0);
    const h = mount.clientHeight - (parseFloat(cs.paddingTop) || 0) - (parseFloat(cs.paddingBottom) || 0);
    const byWidth = w / page.width;
    return m === 'onePage' ? Math.min(byWidth, h / page.height) : byWidth;
  };
  // The page moves left to make room for the comments column beside it (editor.css odx-has-markup), by up to half that
  // column's width, but only as far as the window has room: never past the canvas's left edge, where scrolling can't
  // reach. (A zoomed element's own lengths are zoomed too: given in the page's px.)
  const shift = () => {
    const cs = getComputedStyle(mount);
    const w = mount.clientWidth - (parseFloat(cs.paddingLeft) || 0) - (parseFloat(cs.paddingRight) || 0);
    const room = Math.max(0, (w - page.width * value) / 2);
    sheet.style.setProperty('--odx-markup-left', `${-Math.min(140, room) / value}px`);
  };
  const apply = (z, say = true) => {
    const next = Math.round(Math.min(MAX, Math.max(MIN, z)) * 1000) / 1000;
    if (next === value && sheet.style.zoom) return;
    value = next;
    sheet.style.zoom = value === 1 ? '' : String(value);
    shift();
    onChange(api);
    if (say) announce(`Zoom ${Math.round(value * 100)}%`);
  };
  const api = {
    get value() { return value; },
    get mode() { return mode; },
    /** A zoom (1 is 100%), or 'pageWidth' / 'onePage' to fit the window and keep fitting it. */
    set(z, quiet = false) {
      if (z === 'pageWidth' || z === 'onePage') { mode = z; apply(fitted(z), !quiet); }
      else { mode = null; apply(Number(z) || 1, !quiet); }
    },
    /** The next zoom step up (dir 1) or down (-1), as Word's zoom buttons go. */
    step(dir) {
      const cur = Math.round(value * 100) / 100;
      const next = dir > 0 ? ZOOM_STEPS.find((s) => s > cur + 1e-6) : [...ZOOM_STEPS].reverse().find((s) => s < cur - 1e-6);
      api.set(next ?? (dir > 0 ? MAX : MIN));
    },
    /** Drawn at 100% for a while (a PDF export measuring the pages), then back. */
    suspend() { sheet.style.zoom = ''; },
    resume() { sheet.style.zoom = value === 1 ? '' : String(value); },
    destroy() { ro?.disconnect(); mount.removeEventListener('wheel', onWheel); },
  };
  // Ctrl+wheel zooms the pages (not the browser's page), in Word's 10% steps.
  const onWheel = (e) => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    mode = null;
    apply(Math.round((value + (e.deltaY < 0 ? 0.1 : -0.1)) * 10) / 10);
  };
  mount.addEventListener('wheel', onWheel, { passive: false });
  // Fitted zooms follow the window.
  const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(() => { if (mode) apply(fitted(mode), false); shift(); }) : null;
  ro?.observe(mount);
  shift();
  return api;
}

/**
 * The zoom an editor view's pages are drawn at now (1 when it has none; 1 too while the layout measures them at 100%).
 */
export const zoomOf = (view) => parseFloat((view?.odx?.main || view)?.dom?.parentElement?.style.zoom) || 1;
