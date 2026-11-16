// A drawing's place on its line or page (image nodes' `place`, reader.js readDrawing) is CSS, and carries custom
// properties the layout reads back: --odx-pad (the room drawn round a shape's outline, reader.js padded), --odx-lrel /
// --odx-ly / --odx-lyalign and --odx-arel / --odx-ay / --odx-ayalign (anchorCss: placed against the page), --odx-grown
// (a text box grown to fit its text). Read here, in one place, rather than by a regex at each reader: six of them had
// read --odx-pad, one allowing a space after the colon and five not.

/** The value of custom property `--name` in a place's (or a style attribute's) CSS, trimmed; null when it has none. */
export function placeVar(place, name) {
  if (!place) return null;
  const want = `--${name}`;
  for (const part of String(place).split(';')) {
    const i = part.indexOf(':');
    if (i > 0 && part.slice(0, i).trim() === want) return part.slice(i + 1).trim();
  }
  return null;
}

/** A place's --odx-pad in px (0 for none). */
export const padOf = (place) => parseFloat(placeVar(place, 'odx-pad')) || 0;
