// A picture's or shape's text wrapping (Word's Wrap Text): in line with the text (wp:inline), or floating (wp:anchor)
// with the text square around it, tight to its outline, through, above and below it, or none (behind or in front of
// the text). The run XML is rewritten in the order Word writes it, so Word reads what was chosen.

export const WRAPS = [
  ['inline', 'In Line with Text'], ['square', 'Square'], ['tight', 'Tight'], ['through', 'Through'],
  ['topAndBottom', 'Top and Bottom'], ['behind', 'Behind Text'], ['front', 'In Front of Text'],
];

const WRAP_EL = /<wp:wrap(None|Square|Tight|Through|TopAndBottom)\b[^>]*?(?:\/>|>[\s\S]*?<\/wp:wrap\1>)/;
// (The polygon Word draws round a picture with no outline of its own: its box.)
const POLY = '<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>';

/** The wrapping of the first drawing in a run's XML (one of WRAPS' names), or null when it has none. */
export function wrapOf(runXml) {
  if (/<wp:inline\b/.test(runXml)) return 'inline';
  const anchor = /<wp:anchor\b[^>]*>/.exec(runXml || '');
  if (!anchor) return null;
  const w = WRAP_EL.exec(runXml)?.[1] || 'None';
  if (w === 'None') return /\sbehindDoc="(?:1|true|on)"/.test(anchor[0]) ? 'behind' : 'front';
  return w[0].toLowerCase() + w.slice(1);
}

const wrapEl = (kind) => ({
  square: '<wp:wrapSquare wrapText="bothSides"/>',
  tight: `<wp:wrapTight wrapText="bothSides">${POLY}</wp:wrapTight>`,
  through: `<wp:wrapThrough wrapText="bothSides">${POLY}</wp:wrapThrough>`,
  topAndBottom: '<wp:wrapTopAndBottom/>',
  behind: '<wp:wrapNone/>',
  front: '<wp:wrapNone/>',
}[kind]);

/** The parts a drawing keeps whichever way it's placed. */
function parts(body) {
  const one = (re) => re.exec(body)?.[0] || '';
  return {
    extent: one(/<wp:extent\b[^>]*\/>/),
    effect: one(/<wp:effectExtent\b[^>]*\/>/),
    docPr: one(/<wp:docPr\b[^>]*?(?:\/>|>[\s\S]*?<\/wp:docPr>)/),
    frame: one(/<wp:cNvGraphicFramePr\b[^>]*?(?:\/>|>[\s\S]*?<\/wp:cNvGraphicFramePr>)/),
    graphic: one(/<a:graphic\b[\s\S]*?<\/a:graphic>/),
  };
}

/** The run XML with its drawings wrapped `kind` (one of WRAPS' names). */
export function setWrap(runXml, kind) {
  if (!WRAPS.some(([k]) => k === kind)) throw new Error(`no such wrapping: ${kind}`);
  const behind = kind === 'behind' ? '1' : '0';
  return runXml
    .replace(/<wp:inline\b[^>]*>([\s\S]*?)<\/wp:inline>/g, (all, body) => {
      if (kind === 'inline') return all;
      const p = parts(body);
      // Floating from where it was: at the paragraph's top, the column's left edge (Word starts from there too).
      return `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="${behind}" locked="0" layoutInCell="1" allowOverlap="1">`
        + '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        + `${p.extent}${p.effect || '<wp:effectExtent l="0" t="0" r="0" b="0"/>'}${wrapEl(kind)}${p.docPr}${p.frame}${p.graphic}</wp:anchor>`;
    })
    .replace(/<wp:anchor\b([^>]*)>([\s\S]*?)<\/wp:anchor>/g, (all, attrs, body) => {
      if (kind === 'inline') {
        const p = parts(body);
        return `<wp:inline distT="0" distB="0" distL="0" distR="0">${p.extent}${p.effect}${p.docPr}${p.frame}${p.graphic}</wp:inline>`;
      }
      const open = /\sbehindDoc="/.test(attrs) ? attrs.replace(/\sbehindDoc="[^"]*"/, ` behindDoc="${behind}"`) : `${attrs} behindDoc="${behind}"`;
      const inner = WRAP_EL.test(body) ? body.replace(WRAP_EL, wrapEl(kind)) : body.replace(/(<wp:effectExtent\b[^>]*\/>|<wp:extent\b[^>]*\/>)(?![\s\S]*<wp:effectExtent)/, `$1${wrapEl(kind)}`);
      return `<wp:anchor${open}>${inner}</wp:anchor>`;
    });
}
