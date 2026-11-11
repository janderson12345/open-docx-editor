// View > Ruler: a horizontal ruler over the page with the margins shaded and
// the current paragraph's indents as markers you can drag, as in Word.
import * as cmd from './commands.js';
import { TWIPS_PER_PX } from '../ooxml/units.js';
import { zoomOf } from './zoom.js';

const TW = 1 / TWIPS_PER_PX; // px per twip (96 dpi)

export function createRuler(getMain) {
  const bar = document.createElement('div');
  bar.className = 'odx-ruler';
  const scale = document.createElement('div');
  scale.className = 'odx-ruler-scale';
  const mk = (cls, title) => { const m = document.createElement('div'); m.className = `odx-rm ${cls}`; m.title = title; scale.append(m); return m; };
  const first = mk('odx-rm-first', 'First Line Indent');
  const hang = mk('odx-rm-hang', 'Hanging Indent');
  const left = mk('odx-rm-left', 'Left Indent');
  const right = mk('odx-rm-right', 'Right Indent');
  bar.append(scale);
  let shown = false;
  let geo = null; // { page, marginL, marginR }
  let ind = { left: 0, first: 0, right: 0 }; // twips; first is relative to left

  function drawScale(page) {
    const key = `${page.width}:${page.left}:${page.right}`;
    if (scale.dataset.key === key) return;
    scale.dataset.key = key;
    scale.querySelectorAll('.odx-rt').forEach((t) => t.remove());
    scale.style.width = `${page.width}px`;
    scale.style.setProperty('--ml', `${page.left}px`);
    scale.style.setProperty('--mr', `${page.right}px`);
    // Ticks every 1/8", numbered every inch from the left margin.
    for (let x = -Math.floor(page.left / 12) * 12; page.left + x <= page.width; x += 12) {
      const t = document.createElement('div');
      const inch = Math.round(x / 96 * 8);
      t.className = `odx-rt${inch % 8 === 0 ? ' odx-rt-n' : inch % 4 === 0 ? ' odx-rt-h' : ''}`;
      t.style.left = `${page.left + x}px`;
      if (inch % 8 === 0 && inch) t.textContent = String(Math.abs(inch / 8));
      scale.prepend(t);
    }
  }

  function place() {
    if (!geo) return;
    const { page } = geo;
    const L = page.left + ind.left * TW;
    first.style.left = `${L + ind.first * TW}px`;
    hang.style.left = `${L}px`;
    left.style.left = `${L}px`;
    right.style.left = `${page.width - page.right - ind.right * TW}px`;
  }

  function update(view) {
    if (!shown) return;
    const main = view?.odx?.main || view;
    if (!main || !main.odx) return;
    const { session } = main.odx;
    const page = session.page();
    const sheet = main.dom.parentElement;
    const canvas = sheet.parentElement;
    geo = { page };
    drawScale(page);
    // Line up with the page (which may be shifted for the markup column).
    const cr = canvas.getBoundingClientRect();
    // (Its scale at the page's zoom too: the ruler's marks are in the page's px.)
    const z = zoomOf(main);
    scale.style.zoom = z === 1 ? '' : String(z);
    bar.style.paddingLeft = `${sheet.getBoundingClientRect().left - cr.left + canvas.scrollLeft}px`;
    const p = cmd.selectedParagraphs(view.state)[0]?.node.attrs;
    if (p) {
      const v = p.view || {};
      const st = (k) => session.styleProp(p.styleId, k);
      ind = { left: v.indLeft ?? st('indLeft') ?? 0, first: v.indFirst ?? st('indFirst') ?? 0, right: v.indRight ?? st('indRight') ?? 0 };
    }
    place();
  }

  // Dragging a marker sets that indent on the selected paragraphs.
  let drag = null;
  const startDrag = (which) => (e) => {
    e.preventDefault();
    drag = { which, x0: e.clientX, ind0: { ...ind } };
  };
  first.addEventListener('mousedown', startDrag('first'));
  hang.addEventListener('mousedown', startDrag('hang'));
  left.addEventListener('mousedown', startDrag('left'));
  right.addEventListener('mousedown', startDrag('right'));
  const snap = (tw) => Math.round(tw / 90) * 90; // 1/16"
  document.addEventListener('mousemove', (e) => {
    if (!drag) return;
    const d = snap((e.clientX - drag.x0) / zoomOf(getMain()) / TW); // (screen px back to the page's, whatever the zoom)
    const o = drag.ind0;
    if (drag.which === 'first') ind = { ...o, first: o.first + d };
    else if (drag.which === 'hang') ind = { ...o, left: o.left + d, first: o.first - d }; // first line stays put
    else if (drag.which === 'left') ind = { ...o, left: o.left + d };
    else ind = { ...o, right: o.right - d };
    place();
  });
  document.addEventListener('mouseup', () => {
    if (!drag) return;
    drag = null;
    const view = getMain();
    if (!view) return;
    cmd.setIndent({ left: ind.left, right: ind.right, firstLine: ind.first })(view.state, view.dispatch);
    view.focus();
  });

  return {
    el: bar,
    get shown() { return shown; },
    toggle(view) {
      shown = !shown;
      const main = view?.odx?.main || view;
      if (shown && main) main.dom.parentElement.parentElement.prepend(bar);
      else bar.remove();
      update(view);
      return shown;
    },
    update,
  };
}
