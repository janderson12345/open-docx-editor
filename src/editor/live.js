// Screen-reader announcements (WCAG 4.1.3 status messages): one visually hidden live region for the page.

export function liveRegion() {
  if (typeof document === 'undefined') return null;
  let live = document.querySelector('body > .odx-live');
  if (!live) {
    live = document.createElement('div');
    live.className = 'odx-live odx-sr';
    document.body.append(live);
  }
  return live;
}

/** Says `msg`: a fresh role=alert each time, so NVDA reads it even while it is reading the new selection. */
export function announce(msg) {
  const live = liveRegion();
  if (!live) return;
  setTimeout(() => {
    const a = document.createElement('div');
    a.setAttribute('role', 'alert');
    a.textContent = msg;
    live.replaceChildren(a);
  }, 150);
}
