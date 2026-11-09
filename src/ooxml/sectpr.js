// Page setup edits of a section's w:sectPr, as Layout's menus make them: each takes the w:sectPr XML and returns it
// changed (or the same string when nothing changes).

/** A Letter page with one-inch margins: the section of a document that has none. */
export const DEFAULT_SECTPR = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

/** Sets (or, for null, removes) attributes of the empty element `w:<tag>`. */
export const setAttrs = (xml, tag, vals) => xml.replace(new RegExp(`<w:${tag}\\b[^>]*/>`), (m) => {
  let out = m;
  for (const [k, val] of Object.entries(vals)) {
    const re = new RegExp(`\\sw:${k}="[^"]*"`);
    if (val == null) out = out.replace(re, '');
    else out = re.test(out) ? out.replace(re, ` w:${k}="${val}"`) : out.replace(/\s*\/>$/, ` w:${k}="${val}"/>`);
  }
  return out;
});

/** The page's width and height in twips (Letter when not given). */
export const pageSize = (xml) => ({ w: Number(/<w:pgSz\b[^>]*\sw:w="(\d+)"/.exec(xml)?.[1] || 12240), h: Number(/<w:pgSz\b[^>]*\sw:h="(\d+)"/.exec(xml)?.[1] || 15840) });

export const setMargins = (xml, top, right, bottom, left) => setAttrs(xml, 'pgMar', { top, right, bottom, left });

/** Portrait or landscape: the longer side becomes the height or the width. */
export const setOrientation = (xml, landscape) => {
  const { w, h } = pageSize(xml);
  const [nw, nh] = landscape ? [Math.max(w, h), Math.min(w, h)] : [Math.min(w, h), Math.max(w, h)];
  return setAttrs(xml, 'pgSz', { w: nw, h: nh, orient: landscape ? 'landscape' : null });
};

/** A paper size given portrait (w < h); a landscape page keeps its orientation. */
export const setPageSize = (xml, w, h) => {
  const { w: ow, h: oh } = pageSize(xml);
  return setAttrs(xml, 'pgSz', ow > oh ? { w: h, h: w } : { w, h });
};

/** Equal columns half an inch apart (w:cols), as Word's One / Two / Three. */
export const setColumns = (xml, num) => {
  const cols = num > 1 ? `<w:cols w:space="720" w:num="${num}"/>` : '<w:cols w:space="720"/>';
  if (/<w:cols\b/.test(xml)) return xml.replace(/<w:cols\b[^>]*\/>|<w:cols\b[^>]*>[\s\S]*?<\/w:cols>/, cols);
  // Its place in w:sectPr: before these, else at the end.
  const after = /<w:(?:formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)\b/.exec(xml);
  return after ? xml.slice(0, after.index) + cols + xml.slice(after.index) : xml.replace(/<\/w:sectPr>$/, `${cols}</w:sectPr>`);
};

/** How many columns the section has. */
export const columnCount = (xml) => Number(/<w:cols\b[^>]*\sw:num="(\d+)"/.exec(xml || '')?.[1] || 1);
