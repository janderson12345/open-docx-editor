// Insert > WordArt's styles: text in a floating text box with Word 2010's text effects (w14), in the theme's colours.
// Each style is named by what it does, as Word's gallery names its own (Fill, Outline, Shadow, Glow...); the effects
// are written as Word writes them (odx-test/mkeffects14.ps1: Word's own XML for each), in the schema's order (glow,
// shadow, reflection, outline, fill).

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const clr = (scheme, mods = '') => (mods ? `<w14:schemeClr w14:val="${scheme}">${mods}</w14:schemeClr>` : `<w14:schemeClr w14:val="${scheme}"/>`);
const alpha = (v) => `<w14:alpha w14:val="${v}"/>`;
const lum = (mod, off) => `<w14:lumMod w14:val="${mod}"/>${off ? `<w14:lumOff w14:val="${off}"/>` : ''}`;
const fill = (c) => `<w14:textFill><w14:solidFill>${c}</w14:solidFill></w14:textFill>`;
const outline = (c, w = 9525) => `<w14:textOutline w14:w="${w}" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill>${c}</w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>`;
// (Word's "Offset: Bottom Right" shadow, 60% transparent black; and a hard one, no blur, in a colour.)
const shadow = '<w14:shadow w14:blurRad="38100" w14:dist="19050" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl">'
  + `<w14:srgbClr w14:val="000000">${alpha(60000)}</w14:srgbClr></w14:shadow>`;
const hard = (c) => `<w14:shadow w14:blurRad="0" w14:dist="38100" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl">${c}</w14:shadow>`;
const glow = (c) => `<w14:glow w14:rad="63500">${c}</w14:glow>`;
// (Word's "Tight Reflection: Touching".)
const reflection = '<w14:reflection w14:blurRad="6350" w14:stA="55000" w14:stPos="0" w14:endA="300" w14:endPos="45500" w14:dist="0" w14:dir="5400000" w14:fadeDir="5400000" w14:sx="100000" w14:sy="-100000" w14:kx="0" w14:ky="0" w14:algn="bl"/>';
const gradient = (a, b) => `<w14:textFill><w14:gradFill><w14:gsLst><w14:gs w14:pos="0">${a}</w14:gs><w14:gs w14:pos="100000">${b}</w14:gs></w14:gsLst><w14:lin w14:ang="5400000" w14:scaled="0"/></w14:gradFill></w14:textFill>`;

/** [name, the colour it's in (for w:color: a theme colour's name), its w14 effects] */
export const WORDART = [
  ['Fill: Text 1; Shadow', 'tx1', shadow + fill(clr('tx1'))],
  ['Fill: Accent 1; Shadow', 'accent1', shadow + fill(clr('accent1'))],
  ['Fill: Accent 2, lighter; Outline: Accent 2', 'accent2', outline(clr('accent2')) + fill(clr('accent2', lum(40000, 60000)))],
  ['Fill: White; Outline: Accent 1; Glow: Accent 1', 'bg1', glow(clr('accent1', alpha(60000))) + outline(clr('accent1')) + fill(clr('bg1'))],
  ['Fill: Accent 4; Outline: Accent 4, darker', 'accent4', outline(clr('accent4', lum(50000))) + fill(clr('accent4'))],
  ['Gradient Fill: Gray', 'tx1', gradient(clr('tx1', lum(50000, 50000)), clr('tx1', lum(85000, 15000)))],
  ['Gradient Fill: Accent 1; Reflection', 'accent1', reflection + gradient(clr('accent1', lum(60000, 40000)), clr('accent1', lum(75000)))],
  ['Gradient Fill: Accent 4; Outline: Accent 4', 'accent4', outline(clr('accent4')) + gradient(clr('accent4', lum(40000, 60000)), clr('accent4'))],
  ['Fill: White; Outline: Accent 1; Shadow', 'bg1', shadow + outline(clr('accent1')) + fill(clr('bg1'))],
  ['Fill: Text 1; Outline: White; Shadow', 'tx1', shadow + outline(clr('bg1')) + fill(clr('tx1'))],
  ['Fill: Accent 1; Outline: White; Hard Shadow: Accent 1', 'accent1', hard(clr('accent1')) + outline(clr('bg1')) + fill(clr('accent1'))],
  ['Fill: White; Outline: Accent 2; Hard Shadow: Accent 2', 'bg1', hard(clr('accent2')) + outline(clr('accent2')) + fill(clr('bg1'))],
  ['Fill: Accent 2; Reflection', 'accent2', reflection + fill(clr('accent2'))],
  ['Fill: Text 1; Glow: Accent 4', 'tx1', glow(clr('accent4', alpha(60000))) + fill(clr('tx1'))],
  ['No Fill; Outline: Text 1', 'tx1', outline(clr('tx1'), 12700) + '<w14:textFill><w14:noFill/></w14:textFill>'],
];

/** The w:rPr of WordArt style `i` at `size` half-points, its w:color the theme colour's value (`colors`: the theme's). */
export function wordArtRPr(i, size, colors = {}) {
  const [, scheme, fx] = WORDART[i] || WORDART[0];
  const key = { tx1: 'dk1', bg1: 'lt1' }[scheme] || scheme;
  const hex = (colors[key] || { dk1: '000000', lt1: 'FFFFFF', accent1: '156082', accent2: 'E97132', accent4: '0F9ED5' }[key] || '000000').toUpperCase();
  const theme = { tx1: 'text1', bg1: 'background1' }[scheme] || scheme;
  return `<w:rPr xmlns:w14="${W14}"><w:color w:val="${hex}" w:themeColor="${theme}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/>${fx}</w:rPr>`;
}

/** WordArt style `i`'s effects as [localName, xml] pairs, as a run's raw properties are kept (for drawing its preview). */
export function wordArtItems(i) {
  const fx = WORDART[i]?.[2] || '';
  const out = [];
  let depth = 0, start = 0, name = null;
  for (const m of fx.matchAll(/<(\/?)w14:(\w+)[^>]*?(\/?)>/g)) {
    if (!m[1] && depth === 0) { start = m.index; name = m[2]; }
    depth += m[1] ? -1 : m[3] ? 0 : 1;
    if (depth === 0) out.push([name, fx.slice(start, m.index + m[0].length)]);
  }
  return out;
}
