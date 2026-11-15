// Insert > WordArt's styles: text in a floating text box with Word 2010's text effects (w14), in the theme's colours.
// Each is Word's own, in its gallery's order, as Word writes it (its run properties copied from a document with all
// fifteen inserted from Word's gallery); a theme colour's value is filled in from the document's theme. The names say
// what each does, as Word's gallery names its own (Fill, Outline, Shadow, Glow...), by theme colour.
import { modify } from './shapes.js';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

/** [name, its w:rPr children as Word writes them ({color}: the theme colour's value)] */
export const WORDART = [
  ['Fill: Text 1; Shadow', '<w:color w:val="{color}" w:themeColor="text1"/><w:sz w:val="72"/><w:szCs w:val="72"/><w14:shadow w14:blurRad="38100" w14:dist="19050" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl"><w14:schemeClr w14:val="dk1"><w14:alpha w14:val="60000"/></w14:schemeClr></w14:shadow><w14:textOutline w14:w="0" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:noFill/><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>'],
  ['Fill: Accent 1; Shadow', '<w:color w:val="{color}" w:themeColor="accent1"/><w:sz w:val="72"/><w:szCs w:val="72"/><w14:shadow w14:blurRad="38100" w14:dist="25400" w14:dir="5400000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="ctr"><w14:srgbClr w14:val="6E747A"><w14:alpha w14:val="57000"/></w14:srgbClr></w14:shadow><w14:textOutline w14:w="0" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:noFill/><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>'],
  ['Fill: Accent 2, lighter; Outline: Accent 2', '<w:b/><w:color w:val="{color}" w:themeColor="accent2" w:themeTint="66"/><w:sz w:val="72"/><w:szCs w:val="72"/><w14:textOutline w14:w="11112" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill><w14:schemeClr w14:val="accent2"/></w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>'],
  ['Fill: White; Outline: Accent 5; Shadow', '<w:b/><w:outline/><w:color w:val="{color}" w:themeColor="accent5"/><w:sz w:val="72"/><w:szCs w:val="72"/><w14:shadow w14:blurRad="38100" w14:dist="22860" w14:dir="5400000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl"><w14:srgbClr w14:val="000000"><w14:alpha w14:val="70000"/></w14:srgbClr></w14:shadow><w14:textOutline w14:w="10160" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill><w14:schemeClr w14:val="accent5"/></w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline><w14:textFill><w14:solidFill><w14:srgbClr w14:val="FFFFFF"/></w14:solidFill></w14:textFill>'],
  ['Fill: Accent 4; Soft Bevel', '<w:b/><w:color w:val="{color}" w:themeColor="accent4"/><w:sz w:val="72"/><w:szCs w:val="72"/><w14:textOutline w14:w="0" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:noFill/><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline><w14:props3d w14:extrusionH="57150" w14:contourW="0" w14:prstMaterial="softEdge"><w14:bevelT w14:w="25400" w14:h="38100" w14:prst="circle"/></w14:props3d>'],
  ['Gradient Fill: Gray', '<w:sz w:val="72"/><w:szCs w:val="72"/><w14:textOutline w14:w="0" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:noFill/><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline><w14:textFill><w14:gradFill><w14:gsLst><w14:gs w14:pos="21000"><w14:srgbClr w14:val="53575C"/></w14:gs><w14:gs w14:pos="88000"><w14:srgbClr w14:val="C5C7CA"/></w14:gs></w14:gsLst><w14:lin w14:ang="5400000" w14:scaled="0"/></w14:gradFill></w14:textFill>'],
  ['Gradient Fill: Accent 5; Reflection', '<w:sz w:val="72"/><w:szCs w:val="72"/><w14:reflection w14:blurRad="6350" w14:stA="53000" w14:stPos="0" w14:endA="300" w14:endPos="35500" w14:dist="0" w14:dir="5400000" w14:fadeDir="5400000" w14:sx="100000" w14:sy="-90000" w14:kx="0" w14:ky="0" w14:algn="bl"/><w14:textOutline w14:w="0" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:noFill/><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline><w14:textFill><w14:gradFill><w14:gsLst><w14:gs w14:pos="0"><w14:schemeClr w14:val="accent5"><w14:lumMod w14:val="50000"/></w14:schemeClr></w14:gs><w14:gs w14:pos="50000"><w14:schemeClr w14:val="accent5"/></w14:gs><w14:gs w14:pos="100000"><w14:schemeClr w14:val="accent5"><w14:lumMod w14:val="60000"/><w14:lumOff w14:val="40000"/></w14:schemeClr></w14:gs></w14:gsLst><w14:lin w14:ang="5400000" w14:scaled="0"/></w14:gradFill></w14:textFill>'],
  ['Gradient Fill: Accent 4; Outline: Accent 4', '<w:b/><w:sz w:val="72"/><w:szCs w:val="72"/><w14:textOutline w14:w="12700" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill><w14:schemeClr w14:val="accent4"/></w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline><w14:textFill><w14:gradFill><w14:gsLst><w14:gs w14:pos="0"><w14:schemeClr w14:val="accent4"/></w14:gs><w14:gs w14:pos="4000"><w14:schemeClr w14:val="accent4"><w14:lumMod w14:val="60000"/><w14:lumOff w14:val="40000"/></w14:schemeClr></w14:gs><w14:gs w14:pos="87000"><w14:schemeClr w14:val="accent4"><w14:lumMod w14:val="20000"/><w14:lumOff w14:val="80000"/></w14:schemeClr></w14:gs></w14:gsLst><w14:lin w14:ang="5400000" w14:scaled="0"/></w14:gradFill></w14:textFill>'],
  ['Fill: White; Outline: Accent 1; Glow: Accent 1', '<w:b/><w:color w:val="70AD47"/><w:spacing w:val="10"/><w:sz w:val="72"/><w:szCs w:val="72"/><w14:glow w14:rad="38100"><w14:schemeClr w14:val="accent1"><w14:alpha w14:val="60000"/></w14:schemeClr></w14:glow><w14:textOutline w14:w="9525" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill><w14:schemeClr w14:val="accent1"/></w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline><w14:textFill><w14:solidFill><w14:srgbClr w14:val="70AD47"><w14:tint w14:val="1000"/></w14:srgbClr></w14:solidFill></w14:textFill>'],
  ['Fill: Accent 3; Sharp Bevel', '<w:b/><w:color w:val="{color}" w:themeColor="accent3"/><w:sz w:val="72"/><w:szCs w:val="72"/><w14:textOutline w14:w="0" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:noFill/><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline><w14:props3d w14:extrusionH="57150" w14:contourW="0" w14:prstMaterial="matte"><w14:bevelT w14:w="63500" w14:h="12700" w14:prst="angle"/><w14:contourClr><w14:schemeClr w14:val="bg1"><w14:lumMod w14:val="65000"/></w14:schemeClr></w14:contourClr></w14:props3d>'],
  ['Fill: Text 1; Outline: Background 1; Shadow', '<w:b/><w:color w:val="{color}" w:themeColor="text1"/><w:sz w:val="72"/><w:szCs w:val="72"/><w14:shadow w14:blurRad="12700" w14:dist="38100" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl"><w14:schemeClr w14:val="bg1"><w14:lumMod w14:val="50000"/></w14:schemeClr></w14:shadow><w14:textOutline w14:w="9525" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill><w14:schemeClr w14:val="bg1"/></w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>'],
  ['Fill: Text 1, lighter; Outline: Background 1; Hard Shadow: Accent 5', '<w:b/><w:color w:val="{color}" w:themeColor="text1" w:themeTint="D9"/><w:sz w:val="72"/><w:szCs w:val="72"/><w14:shadow w14:blurRad="0" w14:dist="38100" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="bl"><w14:schemeClr w14:val="accent5"/></w14:shadow><w14:textOutline w14:w="6731" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill><w14:schemeClr w14:val="bg1"/></w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>'],
  ['Fill: Accent 5; Outline: Background 1; Shadow: Accent 5', '<w:b/><w:color w:val="{color}" w:themeColor="accent5"/><w:sz w:val="72"/><w:szCs w:val="72"/><w14:shadow w14:blurRad="12700" w14:dist="38100" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl"><w14:schemeClr w14:val="accent5"><w14:lumMod w14:val="60000"/><w14:lumOff w14:val="40000"/></w14:schemeClr></w14:shadow><w14:textOutline w14:w="9525" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill><w14:schemeClr w14:val="bg1"/></w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>'],
  ['Fill: White; Outline: Accent 2; Hard Shadow: Accent 2', '<w:b/><w:outline/><w:color w:val="{color}" w:themeColor="accent2"/><w:sz w:val="72"/><w:szCs w:val="72"/><w14:shadow w14:blurRad="0" w14:dist="38100" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl"><w14:schemeClr w14:val="accent2"/></w14:shadow><w14:textOutline w14:w="6604" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill><w14:schemeClr w14:val="accent2"/></w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline><w14:textFill><w14:solidFill><w14:srgbClr w14:val="FFFFFF"/></w14:solidFill></w14:textFill>'],
  ['Fill: Background 2; Inner Shadow', '<w:b/><w:color w:val="{color}" w:themeColor="background2"/><w:spacing w:val="10"/><w:sz w:val="72"/><w:szCs w:val="72"/><w14:shadow w14:blurRad="63500" w14:dist="50800" w14:dir="13500000" w14:sx="0" w14:sy="0" w14:kx="0" w14:ky="0" w14:algn="none"><w14:srgbClr w14:val="000000"><w14:alpha w14:val="50000"/></w14:srgbClr></w14:shadow><w14:textOutline w14:w="0" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:noFill/><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>'],
];

const THEME = { text1: 'dk1', background1: 'lt1', text2: 'dk2', background2: 'lt2' };
const OFFICE = { dk1: '000000', lt1: 'FFFFFF', dk2: '0E2841', lt2: 'E8E8E8', accent1: '156082', accent2: 'E97132', accent3: '196B24', accent4: '0F9ED5', accent5: 'A02B93', accent6: '4EA72E' };
/** A w:color's theme colour as its value: the theme's colour, made lighter (w:themeTint) or darker (w:themeShade). */
function themeHex(attrs, colors) {
  const name = /w:themeColor="(\w+)"/.exec(attrs)?.[1];
  const key = THEME[name] || name;
  let hex = (colors?.[key] || OFFICE[key] || '000000').toUpperCase();
  const tint = /w:themeTint="([0-9A-Fa-f]{2})"/.exec(attrs)?.[1];
  const shade = /w:themeShade="([0-9A-Fa-f]{2})"/.exec(attrs)?.[1];
  if (tint) { const t = parseInt(tint, 16) / 255; hex = modify(modify(hex, 'lumMod', t), 'lumOff', 1 - t); }
  if (shade) hex = modify(hex, 'lumMod', parseInt(shade, 16) / 255);
  return hex;
}

/** The w:rPr of WordArt style `i` at `size` half-points, its theme colour's value from `colors` (the theme's). */
export function wordArtRPr(i, size = 72, colors = {}) {
  const inner = (WORDART[i] || WORDART[0])[1]
    .replace(/<w:color w:val="\{color\}"([^>]*)\/>/, (m, attrs) => `<w:color w:val="${themeHex(attrs, colors)}"${attrs}/>`)
    .replace(/<w:sz w:val="\d+"\/>/, `<w:sz w:val="${size}"/>`).replace(/<w:szCs w:val="\d+"\/>/, `<w:szCs w:val="${size}"/>`);
  return `<w:rPr xmlns:w14="${W14}">${inner}</w:rPr>`;
}

/** How style `i`'s letters are set: bold, and the extra space after each (w:spacing, pt). */
export function wordArtLook(i) {
  const x = (WORDART[i] || WORDART[0])[1];
  return { bold: /<w:b\/>/.test(x), spacing: Number(/<w:spacing w:val="(-?\d+)"/.exec(x)?.[1] || 0) / 20 };
}

/** WordArt style `i`'s w14 effects as [localName, xml] pairs, as a run's raw properties are kept (for drawing its preview). */
export function wordArtItems(i) {
  const fx = (WORDART[i] || WORDART[0])[1];
  const out = [];
  let depth = 0, start = 0, name = null;
  for (const m of fx.matchAll(/<(\/?)w14:(\w+)[^>]*?(\/?)>/g)) {
    if (!m[1] && depth === 0) { start = m.index; name = m[2]; }
    depth += m[1] ? -1 : m[3] ? 0 : 1;
    if (depth === 0) out.push([name, fx.slice(start, m.index + m[0].length)]);
  }
  return out;
}

/** Style `i`'s colour for its preview (CSS), from the theme's colours. */
export function wordArtColor(i, colors = {}) {
  const m = /<w:color w:val="([^"]+)"([^>]*)\/>/.exec((WORDART[i] || WORDART[0])[1]);
  if (!m) return '#000000';
  return `#${m[1] === '{color}' ? themeHex(m[2], colors) : m[1]}`;
}
