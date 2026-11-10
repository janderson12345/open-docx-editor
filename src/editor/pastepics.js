// Pictures in pasted content. One given in full (a data: URL: copied from a web page or another editor) is stored in
// the package as a picture of the document's own, so it is shown and saved like one inserted from a file; a copy of one
// of the document's pictures gets drawing ids of its own (Word wants each wp:docPr id once). And pictures that are links.
import { Plugin } from 'prosemirror-state';
import { Fragment, Slice } from 'prosemirror-model';

/** A data: URL's bytes and type, or null. */
function fromDataUrl(url) {
  const m = /^data:(image\/[\w+.-]+)(;base64)?,(.*)$/s.exec(url || '');
  if (!m) return null;
  try {
    const text = m[2] ? atob(m[3]) : decodeURIComponent(m[3]);
    const bytes = new Uint8Array(text.length);
    for (let i = 0; i < text.length; i++) bytes[i] = text.charCodeAt(i) & 0xff;
    return { bytes, mime: m[1].toLowerCase() };
  } catch { return null; }
}

/** A PNG's or GIF's own size in px from its header, or null. */
function naturalSize({ bytes, mime }) {
  const u32 = (i) => ((bytes[i] << 24) | (bytes[i + 1] << 16) | (bytes[i + 2] << 8) | bytes[i + 3]) >>> 0;
  if (mime === 'image/png' && bytes.length > 24) return { w: u32(16), h: u32(20) };
  if (mime === 'image/gif' && bytes.length > 10) return { w: bytes[6] | (bytes[7] << 8), h: bytes[8] | (bytes[9] << 8) };
  return null;
}

export function pastePictures(session) {
  const fresh = (xml) => {
    session._picId = (session._picId || 1000) + 1;
    const id = session._picId;
    return xml.replace(/(<wp:docPr\b[^>]*\bid=")\d+"/, `$1${id}"`).replace(/(<pic:cNvPr\b[^>]*\bid=")\d+"/, `$1${id}"`);
  };
  const fix = (n) => {
    if (n.type.name !== 'image') return n.isLeaf ? n : n.copy(map(n.content));
    if (n.attrs.xml) return n.type.create({ ...n.attrs, xml: fresh(n.attrs.xml) }, null, n.marks);
    const pic = fromDataUrl(n.attrs.src);
    if (!pic || !session.addImage) return n;
    const nat = naturalSize(pic);
    let w = n.attrs.width || nat?.w || 240;
    let h = n.attrs.height || (nat ? Math.round((nat.h * w) / nat.w) : 180);
    if (w > 624) { h = Math.round((h * 624) / w); w = 624; } // (no wider than a Letter page's text)
    const { xml, src } = session.addImage(pic.bytes, pic.mime, w, h);
    return n.type.create({ ...n.attrs, xml, src: src || n.attrs.src, width: w, height: h }, null, n.marks);
  };
  const map = (frag) => {
    const out = [];
    frag.forEach((n) => out.push(fix(n)));
    return Fragment.fromArray(out);
  };
  return new Plugin({ props: { transformPasted: (slice) => new Slice(map(slice.content), slice.openStart, slice.openEnd) } });
}

/**
 * Ctrl+click (Cmd+click) on a picture that is a link follows it, as Word does (and as the browser does for text links):
 * in the body, or in a header or footer (a footer logo), so on the whole page.
 */
export const pictureLinks = new Plugin({
  view(view) {
    const host = view.dom.closest('.odx-page') || view.dom.parentElement || view.dom;
    const down = (e) => {
      const href = (e.ctrlKey || e.metaKey) && e.target?.closest?.('img[data-href]')?.getAttribute('data-href');
      if (!href) return;
      e.preventDefault();
      e.stopPropagation();
      window.open(href, '_blank', 'noopener');
    };
    host.addEventListener('mousedown', down, true);
    return { destroy: () => host.removeEventListener('mousedown', down, true) };
  },
});
