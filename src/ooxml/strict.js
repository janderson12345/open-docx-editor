// Strict Office Open XML (ISO/IEC 29500 Strict, Word's "Strict Open XML Document"): the same markup as the
// Transitional form everything else here reads, under other namespaces, with percentages written as "50%". A Strict
// package is turned into the Transitional form as it opens, so it's read and edited like any other, and saved as
// Transitional (which every version of Word opens).
import { strFromU8, strToU8 } from 'fflate';

const S = 'http://purl.oclc.org/ooxml/';
const T = 'http://schemas.openxmlformats.org/';
// Strict namespace (or relationship type) prefixes and their Transitional equivalents, longest first.
const MAP = [
  ['officeDocument/relationships/extendedProperties', 'officeDocument/2006/relationships/extended-properties'],
  ['officeDocument/relationships/customProperties', 'officeDocument/2006/relationships/custom-properties'],
  ['officeDocument/relationships', 'officeDocument/2006/relationships'],
  ['officeDocument/extendedProperties', 'officeDocument/2006/extended-properties'],
  ['officeDocument/customProperties', 'officeDocument/2006/custom-properties'],
  ['officeDocument/docPropsVTypes', 'officeDocument/2006/docPropsVTypes'],
  ['officeDocument/sharedTypes', 'officeDocument/2006/sharedTypes'],
  ['officeDocument/bibliography', 'officeDocument/2006/bibliography'],
  ['officeDocument/customXml', 'officeDocument/2006/customXml'],
  ['officeDocument/math', 'officeDocument/2006/math'],
  ['wordprocessingml/main', 'wordprocessingml/2006/main'],
  ['drawingml/wordprocessingDrawing', 'drawingml/2006/wordprocessingDrawing'],
  ['drawingml/chartDrawing', 'drawingml/2006/chartDrawing'],
  ['drawingml/chart', 'drawingml/2006/chart'],
  ['drawingml/picture', 'drawingml/2006/picture'],
  ['drawingml/diagram', 'drawingml/2006/diagram'],
  ['drawingml/main', 'drawingml/2006/main'],
  ['spreadsheetml/main', 'spreadsheetml/2006/main'],
];

/** Whether the package is Strict: its main relationships or parts use the Strict namespaces. */
export const isStrict = (files) => Object.entries(files).some(([p, b]) => /(\.rels|document\.xml)$/.test(p) && strFromU8(b.subarray(0, 4096)).includes(S));

/** The package's parts in the Transitional form (a new map; parts with nothing Strict in them are the same bytes). */
export function fromStrict(files) {
  const out = {};
  for (const [path, bytes] of Object.entries(files)) {
    if (!/\.(xml|rels)$/i.test(path)) { out[path] = bytes; continue; }
    const text = strFromU8(bytes);
    if (!text.includes(S)) { out[path] = bytes; continue; }
    let t = text;
    for (const [s, tr] of MAP) t = t.split(S + s).join(T + tr);
    t = t.replace(/\s(?:w:)?conformance="strict"/g, '');
    // Percentages: WordprocessingML widths in fiftieths of a percent, DrawingML values in thousandths. (Tag by tag, each
    // looked at once: a pattern spanning tags took minutes over a crafted part full of unclosed ones.)
    // Charts' (gap width, overlap, hole size) are plain percentages.
    t = t.replace(/<(w|a|c):[^<>]*>/g, (tag, ns) => (!tag.includes('%"') ? tag : ns === 'w'
      ? tag.replace(/(\sw:w=")(-?\d+(?:\.\d+)?)%"/g, (_, a, v) => `${a}${Math.round(Number(v) * 50)}"`)
      : ns === 'c' ? tag.replace(/(\sval=")(-?\d+(?:\.\d+)?)%"/g, (_, a, v) => `${a}${Math.round(Number(v))}"`)
        : tag.replace(/(\s(?:val|pos|amt|thresh)=")(-?\d+(?:\.\d+)?)%"/g, (_, a, v) => `${a}${Math.round(Number(v) * 1000)}"`)));
    out[path] = strToU8(t);
  }
  return out;
}
