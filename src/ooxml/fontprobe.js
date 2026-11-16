// Whether the browser has a font: document.fonts.check() says yes for a font that isn't there at all, so availability
// is found the usual way: text set in the font is measured against the same text in generic families, and a font that
// draws differently from each fallback it would have is there. One way of asking for the three places that ask
// (fontcheck.js, svgfonts.js, props.js), each with the probe it was fitted with.

let canvas = null;

/**
 * Whether `family` draws differently from one of `generics` (so it is installed, or loaded by an @font-face rule):
 * `probe` text at `size` px. False when there is no canvas (no page).
 */
export function drawsOwn(family, { probe = 'mmmmmmmmmmlli1WQ@#', size = 40, generics = ['monospace', 'serif'] } = {}) {
  try {
    if (!canvas) canvas = document.createElement('canvas').getContext('2d');
    const quoted = `"${String(family).replace(/["\\]/g, '')}"`;
    return generics.some((g) => {
      canvas.font = `${size}px ${g}`;
      const w = canvas.measureText(probe).width;
      canvas.font = `${size}px ${quoted}, ${g}`;
      return canvas.measureText(probe).width !== w;
    });
  } catch {
    return false;
  }
}
