// Wrap Text (Layout > Wrap Text and the picture's context menu): the selected picture or shape in line with the text or
// floating, the text square round it, tight, through, above and below it, or none (behind or in front), as in Word.
// Its XML is rewritten (ooxml/wrap.js) and its place on the page read again from that, so what's drawn is what's saved.
import { NodeSelection } from 'prosemirror-state';
import { wrapOf, setWrap } from '../ooxml/wrap.js';

/** The selected picture or shape whose wrapping can be chosen: { node, pos }, or null. */
export function wrappable(state) {
  const { node, from } = state.selection;
  return node?.type.name === 'image' && /<wp:(?:inline|anchor)\b/.test(node.attrs.xml || '') ? { node, pos: from } : null;
}

/** The selected picture's wrapping (ooxml/wrap.js WRAPS' names), or null. */
export const currentWrap = (state) => { const w = wrappable(state); return w ? wrapOf(w.node.attrs.xml) : null; };

/** Wrap the selected picture `kind`. */
export const setWrapText = (kind) => (state, dispatch, view) => {
  const w = wrappable(state);
  if (!w || !view?.odx?.session) return false;
  if (dispatch) {
    const xml = setWrap(w.node.attrs.xml, kind);
    let place = null;
    view.odx.session._fragment(`<w:p>${xml}</w:p>`).descendants((n) => { if (place == null && n.type.name === 'image') place = n.attrs.place || ''; });
    const tr = state.tr.setNodeMarkup(w.pos, null, { ...w.node.attrs, xml, place: place ?? '' });
    dispatch(tr.setSelection(NodeSelection.create(tr.doc, w.pos)).scrollIntoView());
  }
  return true;
};
