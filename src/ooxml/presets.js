// Preset shape definitions (a:prstGeom prst="..."), after ECMA-376 Part 1's presetShapeDefinitions, each checked
// against Word's own drawing of it (a sheet of every shape printed by Word, compared shape by shape). Written one guide
// or path a line:
//   av NAME = FORMULA        an adjust value's default (a shape's own a:avLst replaces it)
//   gd NAME = FORMULA        a guide
//   path[(fill=none|darken|darkenLess|lighten|lightenLess, stroke=0, w=N, h=N)] COMMANDS
// where COMMANDS are M x y, L x y, A wR hR stAng swAng, C x1 y1 x2 y2 x y, Q x1 y1 x y and Z, each value a guide's name
// or a number. geometry.js draws them.

const ARGS = { M: 2, L: 2, A: 4, C: 6, Q: 4, Z: 0 };

/** A definition in the form above, as geometry.js reads one: { av, gd, path }. */
function def(text) {
  const out = { av: [], gd: [], path: [] };
  for (const raw of text.split('\n')) {
    const line = raw.trim();
    if (!line) continue;
    const m = /^(av|gd)\s+(\S+)\s*=\s*(.+)$/.exec(line);
    if (m) { out[m[1]].push([m[2], m[3].trim()]); continue; }
    const p = /^path(?:\(([^)]*)\))?\s+(.*)$/.exec(line);
    if (!p) continue;
    const opts = Object.fromEntries((p[1] || '').split(',').filter(Boolean).map((kv) => kv.split('=').map((s) => s.trim())));
    const toks = p[2].split(/\s+/);
    const cmds = [];
    for (let i = 0; i < toks.length;) {
      const k = toks[i++];
      cmds.push([k, ...toks.slice(i, i + ARGS[k])]);
      i += ARGS[k];
    }
    out.path.push({ fill: opts.fill || 'norm', stroke: opts.stroke !== '0', w: Number(opts.w) || 0, h: Number(opts.h) || 0, cmds });
  }
  return out;
}

const SRC = {
  rect: 'path M l t L r t L r b L l b Z',
  line: 'path(fill=none) M l t L r b',
  ellipse: 'path M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z',
  roundRect: `
    av adj = val 16667
    gd a = pin 0 adj 50000
    gd x1 = */ ss a 100000
    gd x2 = +- r 0 x1
    gd y2 = +- b 0 x1
    path M l x1 A x1 x1 cd2 cd4 L x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4 L x1 b A x1 x1 cd4 cd4 Z`,
  triangle: `
    av adj = val 50000
    gd a = pin 0 adj 100000
    gd x1 = */ w a 100000
    path M l b L x1 t L r b Z`,
  rtTriangle: 'path M l b L l t L r b Z',
  diamond: 'path M l vc L hc t L r vc L hc b Z',
  parallelogram: `
    av adj = val 25000
    gd maxAdj = */ 100000 w ss
    gd a = pin 0 adj maxAdj
    gd x2 = */ ss a 100000
    gd x6 = +- r 0 x2
    path M l b L x2 t L r t L x6 b Z`,
  trapezoid: `
    av adj = val 25000
    gd maxAdj = */ 50000 w ss
    gd a = pin 0 adj maxAdj
    gd x2 = */ ss a 100000
    gd x3 = +- r 0 x2
    path M l b L x2 t L x3 t L r b Z`,
  pentagon: `
    av hf = val 105146
    av vf = val 110557
    gd swd2 = */ wd2 hf 100000
    gd shd2 = */ hd2 vf 100000
    gd svc = */ vc vf 100000
    gd dx1 = cos swd2 1080000
    gd dx2 = cos swd2 18360000
    gd dy1 = sin shd2 1080000
    gd dy2 = sin shd2 18360000
    gd x1 = +- hc 0 dx1
    gd x2 = +- hc 0 dx2
    gd x3 = +- hc dx2 0
    gd x4 = +- hc dx1 0
    gd y1 = +- svc 0 dy1
    gd y2 = +- svc 0 dy2
    path M x1 y1 L hc t L x4 y1 L x3 y2 L x2 y2 Z`,
  hexagon: `
    av adj = val 25000
    av vf = val 115470
    gd maxAdj = */ 50000 w ss
    gd a = pin 0 adj maxAdj
    gd shd2 = */ hd2 vf 100000
    gd x1 = */ ss a 100000
    gd x2 = +- r 0 x1
    gd dy1 = sin shd2 3600000
    gd y1 = +- vc 0 dy1
    gd y2 = +- vc dy1 0
    path M l vc L x1 y1 L x2 y1 L r vc L x2 y2 L x1 y2 Z`,
  heptagon: `
    av hf = val 102572
    av vf = val 105210
    gd swd2 = */ wd2 hf 100000
    gd shd2 = */ hd2 vf 100000
    gd svc = */ vc vf 100000
    gd dx1 = */ swd2 97493 100000
    gd dx2 = */ swd2 78183 100000
    gd dx3 = */ swd2 43388 100000
    gd dy1 = */ shd2 62349 100000
    gd dy2 = */ shd2 22252 100000
    gd dy3 = */ shd2 90097 100000
    gd x1 = +- hc 0 dx1
    gd x2 = +- hc 0 dx2
    gd x3 = +- hc 0 dx3
    gd x4 = +- hc dx3 0
    gd x5 = +- hc dx2 0
    gd x6 = +- hc dx1 0
    gd y1 = +- svc 0 dy1
    gd y2 = +- svc dy2 0
    gd y3 = +- svc dy3 0
    path M x1 y2 L x2 y1 L hc t L x5 y1 L x6 y2 L x4 y3 L x3 y3 Z`,
  octagon: `
    av adj = val 29289
    gd a = pin 0 adj 50000
    gd x1 = */ ss a 100000
    gd x2 = +- r 0 x1
    gd y2 = +- b 0 x1
    path M l x1 L x1 t L x2 t L r x1 L r y2 L x2 b L x1 b L l y2 Z`,
  decagon: `
    av vf = val 105146
    gd shd2 = */ hd2 vf 100000
    gd dx1 = cos wd2 2160000
    gd dx2 = cos wd2 4320000
    gd x1 = +- hc 0 dx1
    gd x2 = +- hc 0 dx2
    gd x3 = +- hc dx2 0
    gd x4 = +- hc dx1 0
    gd dy1 = sin shd2 4320000
    gd dy2 = sin shd2 2160000
    gd y1 = +- vc 0 dy1
    gd y2 = +- vc 0 dy2
    gd y3 = +- vc dy2 0
    gd y4 = +- vc dy1 0
    path M l vc L x1 y2 L x2 y1 L x3 y1 L x4 y2 L r vc L x4 y3 L x3 y4 L x2 y4 L x1 y3 Z`,
  dodecagon: `
    gd x1 = */ w 2894 21600
    gd x2 = */ w 7906 21600
    gd x3 = */ w 13694 21600
    gd x4 = */ w 18706 21600
    gd y1 = */ h 2894 21600
    gd y2 = */ h 7906 21600
    gd y3 = */ h 13694 21600
    gd y4 = */ h 18706 21600
    path M l y2 L x1 y1 L x2 t L x3 t L x4 y1 L r y2 L r y3 L x4 y4 L x3 b L x2 b L x1 y4 L l y3 Z`,
  plus: `
    av adj = val 25000
    gd a = pin 0 adj 50000
    gd x1 = */ ss a 100000
    gd x2 = +- r 0 x1
    gd y2 = +- b 0 x1
    path M l x1 L x1 x1 L x1 t L x2 t L x2 x1 L r x1 L r y2 L x2 y2 L x2 b L x1 b L x1 y2 L l y2 Z`,
  star4: `
    av adj = val 12500
    gd a = pin 0 adj 50000
    gd iwd2 = */ wd2 a 50000
    gd ihd2 = */ hd2 a 50000
    gd sdx = cos iwd2 2700000
    gd sdy = sin ihd2 2700000
    gd sx1 = +- hc 0 sdx
    gd sx2 = +- hc sdx 0
    gd sy1 = +- vc 0 sdy
    gd sy2 = +- vc sdy 0
    path M l vc L sx1 sy1 L hc t L sx2 sy1 L r vc L sx2 sy2 L hc b L sx1 sy2 Z`,
  star5: `
    av adj = val 19098
    av hf = val 105146
    av vf = val 110557
    gd a = pin 0 adj 50000
    gd swd2 = */ wd2 hf 100000
    gd shd2 = */ hd2 vf 100000
    gd svc = */ vc vf 100000
    gd dx1 = cos swd2 1080000
    gd dx2 = cos swd2 18360000
    gd dy1 = sin shd2 1080000
    gd dy2 = sin shd2 18360000
    gd x1 = +- hc 0 dx1
    gd x2 = +- hc 0 dx2
    gd x3 = +- hc dx2 0
    gd x4 = +- hc dx1 0
    gd y1 = +- svc 0 dy1
    gd y2 = +- svc 0 dy2
    gd iwd2 = */ swd2 a 50000
    gd ihd2 = */ shd2 a 50000
    gd sdx1 = cos iwd2 20520000
    gd sdx2 = cos iwd2 3240000
    gd sdy1 = sin ihd2 3240000
    gd sdy2 = sin ihd2 20520000
    gd sx1 = +- hc 0 sdx1
    gd sx2 = +- hc 0 sdx2
    gd sx3 = +- hc sdx2 0
    gd sx4 = +- hc sdx1 0
    gd sy1 = +- svc 0 sdy1
    gd sy2 = +- svc 0 sdy2
    gd sy3 = +- svc ihd2 0
    path M x1 y1 L sx2 sy1 L hc t L sx3 sy1 L x4 y1 L sx4 sy2 L x3 y2 L hc sy3 L x2 y2 L sx1 sy2 Z`,
  star6: `
    av adj = val 28868
    av hf = val 115470
    gd a = pin 0 adj 50000
    gd swd2 = */ wd2 hf 100000
    gd dx1 = cos swd2 1800000
    gd x1 = +- hc 0 dx1
    gd x2 = +- hc dx1 0
    gd y2 = +- vc hd4 0
    gd iwd2 = */ swd2 a 50000
    gd ihd2 = */ hd2 a 50000
    gd sdx2 = */ iwd2 1 2
    gd sx1 = +- hc 0 iwd2
    gd sx2 = +- hc 0 sdx2
    gd sx3 = +- hc sdx2 0
    gd sx4 = +- hc iwd2 0
    gd sdy1 = sin ihd2 3600000
    gd sy1 = +- vc 0 sdy1
    gd sy2 = +- vc sdy1 0
    path M x1 hd4 L sx2 sy1 L hc t L sx3 sy1 L x2 hd4 L sx4 vc L x2 y2 L sx3 sy2 L hc b L sx2 sy2 L x1 y2 L sx1 vc Z`,
  star8: `
    av adj = val 38268
    gd a = pin 0 adj 50000
    gd dx1 = cos wd2 2700000
    gd x1 = +- hc 0 dx1
    gd x2 = +- hc dx1 0
    gd dy1 = sin hd2 2700000
    gd y1 = +- vc 0 dy1
    gd y2 = +- vc dy1 0
    gd iwd2 = */ wd2 a 50000
    gd ihd2 = */ hd2 a 50000
    gd sdx1 = */ iwd2 92388 100000
    gd sdx2 = */ iwd2 38268 100000
    gd sdy1 = */ ihd2 92388 100000
    gd sdy2 = */ ihd2 38268 100000
    gd sx1 = +- hc 0 sdx1
    gd sx2 = +- hc 0 sdx2
    gd sx3 = +- hc sdx2 0
    gd sx4 = +- hc sdx1 0
    gd sy1 = +- vc 0 sdy1
    gd sy2 = +- vc 0 sdy2
    gd sy3 = +- vc sdy2 0
    gd sy4 = +- vc sdy1 0
    path M l vc L sx1 sy2 L x1 y1 L sx2 sy1 L hc t L sx3 sy1 L x2 y1 L sx4 sy2 L r vc L sx4 sy3 L x2 y2 L sx3 sy4 L hc b L sx2 sy4 L x1 y2 L sx1 sy3 Z`,
  rightArrow: `
    av adj1 = val 50000
    av adj2 = val 50000
    gd maxAdj2 = */ 100000 w ss
    gd a1 = pin 0 adj1 100000
    gd a2 = pin 0 adj2 maxAdj2
    gd dx1 = */ ss a2 100000
    gd x1 = +- r 0 dx1
    gd dy1 = */ h a1 200000
    gd y1 = +- vc 0 dy1
    gd y2 = +- vc dy1 0
    path M l y1 L x1 y1 L x1 t L r vc L x1 b L x1 y2 L l y2 Z`,
  leftArrow: `
    av adj1 = val 50000
    av adj2 = val 50000
    gd maxAdj2 = */ 100000 w ss
    gd a1 = pin 0 adj1 100000
    gd a2 = pin 0 adj2 maxAdj2
    gd dx2 = */ ss a2 100000
    gd x2 = +- l dx2 0
    gd dy1 = */ h a1 200000
    gd y1 = +- vc 0 dy1
    gd y2 = +- vc dy1 0
    path M l vc L x2 t L x2 y1 L r y1 L r y2 L x2 y2 L x2 b Z`,
  upArrow: `
    av adj1 = val 50000
    av adj2 = val 50000
    gd maxAdj2 = */ 100000 h ss
    gd a1 = pin 0 adj1 100000
    gd a2 = pin 0 adj2 maxAdj2
    gd dy2 = */ ss a2 100000
    gd y2 = +- t dy2 0
    gd dx1 = */ w a1 200000
    gd x1 = +- hc 0 dx1
    gd x2 = +- hc dx1 0
    path M l y2 L hc t L r y2 L x2 y2 L x2 b L x1 b L x1 y2 Z`,
  downArrow: `
    av adj1 = val 50000
    av adj2 = val 50000
    gd maxAdj2 = */ 100000 h ss
    gd a1 = pin 0 adj1 100000
    gd a2 = pin 0 adj2 maxAdj2
    gd dy1 = */ ss a2 100000
    gd y1 = +- b 0 dy1
    gd dx1 = */ w a1 200000
    gd x1 = +- hc 0 dx1
    gd x2 = +- hc dx1 0
    path M l y1 L x1 y1 L x1 t L x2 t L x2 y1 L r y1 L hc b Z`,
  leftRightArrow: `
    av adj1 = val 50000
    av adj2 = val 50000
    gd maxAdj2 = */ 50000 w ss
    gd a1 = pin 0 adj1 100000
    gd a2 = pin 0 adj2 maxAdj2
    gd x2 = */ ss a2 100000
    gd x3 = +- r 0 x2
    gd dy = */ h a1 200000
    gd y1 = +- vc 0 dy
    gd y2 = +- vc dy 0
    path M l vc L x2 t L x2 y1 L x3 y1 L x3 t L r vc L x3 b L x3 y2 L x2 y2 L x2 b Z`,
  upDownArrow: `
    av adj1 = val 50000
    av adj2 = val 50000
    gd maxAdj2 = */ 50000 h ss
    gd a1 = pin 0 adj1 100000
    gd a2 = pin 0 adj2 maxAdj2
    gd y2 = */ ss a2 100000
    gd y3 = +- b 0 y2
    gd dx1 = */ w a1 200000
    gd x1 = +- hc 0 dx1
    gd x2 = +- hc dx1 0
    path M l y2 L hc t L r y2 L x2 y2 L x2 y3 L r y3 L hc b L l y3 L x1 y3 L x1 y2 Z`,
  notchedRightArrow: `
    av adj1 = val 50000
    av adj2 = val 50000
    gd maxAdj2 = */ 100000 w ss
    gd a1 = pin 0 adj1 100000
    gd a2 = pin 0 adj2 maxAdj2
    gd dx2 = */ ss a2 100000
    gd x2 = +- r 0 dx2
    gd dy1 = */ h a1 200000
    gd y1 = +- vc 0 dy1
    gd y2 = +- vc dy1 0
    gd x1 = */ dy1 dx2 hd2
    path M l y1 L x2 y1 L x2 t L r vc L x2 b L x2 y2 L l y2 L x1 vc Z`,
  stripedRightArrow: `
    av adj1 = val 50000
    av adj2 = val 50000
    gd maxAdj2 = */ 84375 w ss
    gd a1 = pin 0 adj1 100000
    gd a2 = pin 0 adj2 maxAdj2
    gd x4 = */ ss 5 32
    gd dx5 = */ ss a2 100000
    gd x5 = +- r 0 dx5
    gd dy1 = */ h a1 200000
    gd y1 = +- vc 0 dy1
    gd y2 = +- vc dy1 0
    path M l y1 L ssd32 y1 L ssd32 y2 L l y2 Z M ssd16 y1 L ssd8 y1 L ssd8 y2 L ssd16 y2 Z M x4 y1 L x5 y1 L x5 t L r vc L x5 b L x5 y2 L x4 y2 Z`,
  chevron: `
    av adj = val 50000
    gd maxAdj = */ 100000 w ss
    gd a = pin 0 adj maxAdj
    gd x1 = */ ss a 100000
    gd x2 = +- r 0 x1
    path M l t L x2 t L r vc L x2 b L l b L x1 vc Z`,
  homePlate: `
    av adj = val 50000
    gd maxAdj = */ 100000 w ss
    gd a = pin 0 adj maxAdj
    gd dx1 = */ ss a 100000
    gd x1 = +- r 0 dx1
    path M l t L x1 t L r vc L x1 b L l b Z`,
  can: `
    av adj = val 25000
    gd maxAdj = */ 50000 h ss
    gd a = pin 0 adj maxAdj
    gd y1 = */ ss a 200000
    gd y3 = +- b 0 y1
    path(stroke=0) M l y1 A wd2 y1 cd2 -10800000 L r y3 A wd2 y1 0 cd2 Z
    path(fill=lighten, stroke=0) M l y1 A wd2 y1 cd2 cd2 A wd2 y1 0 cd2 Z
    path(fill=none) M r y1 A wd2 y1 0 cd2 A wd2 y1 cd2 cd2 L r y3 A wd2 y1 0 cd2 L l y1`,
  cube: `
    av adj = val 25000
    gd a = pin 0 adj 100000
    gd y1 = */ ss a 100000
    gd y4 = +- b 0 y1
    gd x4 = +- r 0 y1
    path(stroke=0) M l y1 L x4 y1 L x4 b L l b Z
    path(fill=darkenLess, stroke=0) M x4 y1 L r t L r y4 L x4 b Z
    path(fill=lightenLess, stroke=0) M l y1 L y1 t L r t L x4 y1 Z
    path(fill=none) M l y1 L y1 t L r t L r y4 L x4 b L l b Z M l y1 L x4 y1 L r t M x4 y1 L x4 b`,
  frame: `
    av adj1 = val 12500
    gd a1 = pin 0 adj1 50000
    gd x1 = */ ss a1 100000
    gd x4 = +- r 0 x1
    gd y4 = +- b 0 x1
    path M l t L r t L r b L l b Z M x1 x1 L x1 y4 L x4 y4 L x4 x1 Z`,
  snip1Rect: `
    av adj = val 16667
    gd a = pin 0 adj 50000
    gd dx1 = */ ss a 100000
    gd x1 = +- r 0 dx1
    path M l t L x1 t L r dx1 L r b L l b Z`,
  round1Rect: `
    av adj = val 16667
    gd a = pin 0 adj 50000
    gd dx1 = */ ss a 100000
    gd x1 = +- r 0 dx1
    path M l t L x1 t A dx1 dx1 3cd4 cd4 L r b L l b Z`,
  snip2SameRect: `
    av adj1 = val 16667
    av adj2 = val 0
    gd a1 = pin 0 adj1 50000
    gd a2 = pin 0 adj2 50000
    gd tx1 = */ ss a1 100000
    gd tx2 = +- r 0 tx1
    gd bx1 = */ ss a2 100000
    gd bx2 = +- r 0 bx1
    gd by1 = +- b 0 bx1
    path M tx1 t L tx2 t L r tx1 L r by1 L bx2 b L bx1 b L l by1 L l tx1 Z`,
  round2SameRect: `
    av adj1 = val 16667
    av adj2 = val 0
    gd a1 = pin 0 adj1 50000
    gd a2 = pin 0 adj2 50000
    gd tx1 = */ ss a1 100000
    gd tx2 = +- r 0 tx1
    gd bx1 = */ ss a2 100000
    gd bx2 = +- r 0 bx1
    gd by1 = +- b 0 bx1
    path M tx1 t L tx2 t A tx1 tx1 3cd4 cd4 L r by1 A bx1 bx1 0 cd4 L bx1 b A bx1 bx1 cd4 cd4 L l tx1 A tx1 tx1 cd2 cd4 Z`,
  plaque: `
    av adj = val 16667
    gd a = pin 0 adj 50000
    gd x1 = */ ss a 100000
    gd x2 = +- r 0 x1
    gd y2 = +- b 0 x1
    path M l x1 A x1 x1 cd4 -5400000 L x2 t A x1 x1 cd2 -5400000 L r y2 A x1 x1 3cd4 -5400000 L x1 b A x1 x1 0 -5400000 Z`,
  foldedCorner: `
    av adj = val 16667
    gd a = pin 0 adj 50000
    gd dy2 = */ ss a 100000
    gd dy1 = */ dy2 1 5
    gd x1 = +- r 0 dy2
    gd x2 = +- x1 dy1 0
    gd y2 = +- b 0 dy2
    gd y1 = +- y2 dy1 0
    path(stroke=0) M l t L r t L r y2 L x1 b L l b Z
    path(fill=darkenLess, stroke=0) M x1 b L x2 y1 L r y2 Z
    path(fill=none) M x1 b L x2 y1 L r y2 L x1 b L l b L l t L r t L r y2`,
  bevel: `
    av adj = val 12500
    gd a = pin 0 adj 50000
    gd x1 = */ ss a 100000
    gd x2 = +- r 0 x1
    gd y2 = +- b 0 x1
    path(stroke=0) M x1 x1 L x2 x1 L x2 y2 L x1 y2 Z
    path(fill=lightenLess, stroke=0) M l t L r t L x2 x1 L x1 x1 Z
    path(fill=darken, stroke=0) M l b L x1 y2 L x2 y2 L r b Z
    path(fill=lighten, stroke=0) M l t L x1 x1 L x1 y2 L l b Z
    path(fill=darkenLess, stroke=0) M r t L r b L x2 y2 L x2 x1 Z
    path(fill=none) M l t L r t L r b L l b Z M x1 x1 L x2 x1 L x2 y2 L x1 y2 Z M l t L x1 x1 M l b L x1 y2 M r t L x2 x1 M r b L x2 y2`,
  halfFrame: `
    av adj1 = val 33333
    av adj2 = val 33333
    gd maxAdj2 = */ 100000 w ss
    gd a2 = pin 0 adj2 maxAdj2
    gd x1 = */ ss a2 100000
    gd g1 = */ h x1 w
    gd g2 = +- h 0 g1
    gd maxAdj1 = */ 100000 g2 ss
    gd a1 = pin 0 adj1 maxAdj1
    gd y1 = */ ss a1 100000
    gd dx2 = */ y1 w h
    gd x2 = +- r 0 dx2
    gd dy2 = */ x1 h w
    gd y2 = +- b 0 dy2
    path M l t L r t L x2 y1 L x1 y1 L x1 y2 L l b Z`,
  corner: `
    av adj1 = val 50000
    av adj2 = val 50000
    gd maxAdj1 = */ 100000 h ss
    gd maxAdj2 = */ 100000 w ss
    gd a1 = pin 0 adj1 maxAdj1
    gd a2 = pin 0 adj2 maxAdj2
    gd x1 = */ ss a2 100000
    gd dy1 = */ ss a1 100000
    gd y1 = +- b 0 dy1
    path M l t L x1 t L x1 y1 L r y1 L r b L l b Z`,
  diagStripe: `
    av adj = val 50000
    gd a = pin 0 adj 100000
    gd x2 = */ w a 100000
    gd y2 = */ h a 100000
    path M l y2 L x2 t L r t L l b Z`,
  donut: `
    av adj = val 25000
    gd a = pin 0 adj 50000
    gd dr = */ ss a 100000
    gd iwd2 = +- wd2 0 dr
    gd ihd2 = +- hd2 0 dr
    path M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z M dr vc A iwd2 ihd2 cd2 -5400000 A iwd2 ihd2 cd4 -5400000 A iwd2 ihd2 0 -5400000 A iwd2 ihd2 3cd4 -5400000 Z`,
  pie: `
    av adj1 = val 0
    av adj2 = val 16200000
    gd stAng = pin 0 adj1 21599999
    gd enAng = pin 0 adj2 21599999
    gd sw1 = +- enAng 0 stAng
    gd sw2 = +- sw1 21600000 0
    gd swAng = ?: sw1 sw1 sw2
    gd wt1 = sin wd2 stAng
    gd ht1 = cos hd2 stAng
    gd dx1 = cat2 wd2 ht1 wt1
    gd dy1 = sat2 hd2 ht1 wt1
    gd x1 = +- hc dx1 0
    gd y1 = +- vc dy1 0
    path M x1 y1 A wd2 hd2 stAng swAng L hc vc Z`,
  chord: `
    av adj1 = val 2700000
    av adj2 = val 16200000
    gd stAng = pin 0 adj1 21599999
    gd enAng = pin 0 adj2 21599999
    gd sw1 = +- enAng 0 stAng
    gd sw2 = +- sw1 21600000 0
    gd swAng = ?: sw1 sw1 sw2
    gd wt1 = sin wd2 stAng
    gd ht1 = cos hd2 stAng
    gd dx1 = cat2 wd2 ht1 wt1
    gd dy1 = sat2 hd2 ht1 wt1
    gd x1 = +- hc dx1 0
    gd y1 = +- vc dy1 0
    path M x1 y1 A wd2 hd2 stAng swAng Z`,
  arc: `
    av adj1 = val 16200000
    av adj2 = val 0
    gd stAng = pin 0 adj1 21599999
    gd enAng = pin 0 adj2 21599999
    gd sw11 = +- enAng 0 stAng
    gd sw12 = +- sw11 21600000 0
    gd swAng = ?: sw11 sw11 sw12
    gd wt1 = sin wd2 stAng
    gd ht1 = cos hd2 stAng
    gd dx1 = cat2 wd2 ht1 wt1
    gd dy1 = sat2 hd2 ht1 wt1
    gd x1 = +- hc dx1 0
    gd y1 = +- vc dy1 0
    path(stroke=0) M x1 y1 A wd2 hd2 stAng swAng L hc vc Z
    path(fill=none) M x1 y1 A wd2 hd2 stAng swAng`,
  blockArc: `
    av adj1 = val 10800000
    av adj2 = val 0
    av adj3 = val 25000
    gd stAng = pin 0 adj1 21599999
    gd istAng = pin 0 adj2 21599999
    gd a3 = pin 0 adj3 50000
    gd sw11 = +- istAng 0 stAng
    gd sw12 = +- sw11 21600000 0
    gd swAng = ?: sw11 sw11 sw12
    gd iswAng = +- 0 0 swAng
    gd wt1 = sin wd2 stAng
    gd ht1 = cos hd2 stAng
    gd dx1 = cat2 wd2 ht1 wt1
    gd dy1 = sat2 hd2 ht1 wt1
    gd x1 = +- hc dx1 0
    gd y1 = +- vc dy1 0
    gd dr = */ ss a3 100000
    gd iwd2 = +- wd2 0 dr
    gd ihd2 = +- hd2 0 dr
    gd wt2 = sin iwd2 istAng
    gd ht2 = cos ihd2 istAng
    gd dx2 = cat2 iwd2 ht2 wt2
    gd dy2 = sat2 ihd2 ht2 wt2
    gd x2 = +- hc dx2 0
    gd y2 = +- vc dy2 0
    path M x1 y1 A wd2 hd2 stAng swAng L x2 y2 A iwd2 ihd2 istAng iswAng Z`,
  teardrop: `
    av adj = val 100000
    gd a = pin 0 adj 200000
    gd r2 = sqrt 2
    gd tw = */ r2 wd2 1
    gd th = */ r2 hd2 1
    gd sw = */ tw a 100000
    gd sh = */ th a 100000
    gd dx1 = cos sw 2700000
    gd dy1 = sin sh 2700000
    gd x1 = +- hc dx1 0
    gd y1 = +- vc 0 dy1
    gd x2 = +/ hc x1 2
    gd y2 = +/ vc y1 2
    path M l vc A wd2 hd2 cd2 cd4 Q x2 t x1 y1 Q r y2 r vc A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z`,
  heart: `
    gd dx1 = */ w 49 48
    gd dx2 = */ w 10 48
    gd x1 = +- hc 0 dx1
    gd x2 = +- hc 0 dx2
    gd x3 = +- hc dx2 0
    gd x4 = +- hc dx1 0
    gd y1 = +- t 0 hd3
    path M hc hd4 C x3 y1 x4 hd4 hc b C x1 hd4 x2 y1 hc hd4 Z`,
  lightningBolt: 'path(w=21600, h=21600) M 8472 0 L 12860 6080 L 11050 6797 L 16577 12007 L 14767 12877 L 21600 21600 L 10012 14915 L 12222 13987 L 5022 9705 L 7602 8382 L 0 3890 Z',
  irregularSeal1: 'path(w=21600, h=21600) M 10800 5800 L 14522 0 L 14155 5325 L 18380 4457 L 16702 7315 L 21097 8137 L 17607 10475 L 21600 13290 L 16837 12942 L 18145 18095 L 14020 14457 L 13247 19737 L 10532 14935 L 8485 21600 L 7715 15627 L 4762 17617 L 5667 13937 L 135 14587 L 3722 11775 L 0 8615 L 4627 7617 L 370 2295 L 7312 6320 L 8352 2295 Z',
  bracketPair: `
    av adj = val 16667
    gd a = pin 0 adj 50000
    gd x1 = */ ss a 100000
    gd x2 = +- r 0 x1
    gd y2 = +- b 0 x1
    path(stroke=0) M l x1 A x1 x1 cd2 cd4 L x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4 L x1 b A x1 x1 cd4 cd4 Z
    path(fill=none) M x1 b A x1 x1 cd4 cd4 L l x1 A x1 x1 cd2 cd4 M x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4`,
  leftBracket: `
    av adj = val 8333
    gd maxAdj = */ 50000 h ss
    gd a = pin 0 adj maxAdj
    gd y1 = */ ss a 100000
    path(stroke=0) M r b A w y1 cd4 cd4 L l y1 A w y1 cd2 cd4 Z
    path(fill=none) M r b A w y1 cd4 cd4 L l y1 A w y1 cd2 cd4`,
  bracePair: `
    av adj = val 8333
    gd a = pin 0 adj 25000
    gd x1 = */ ss a 100000
    gd x2 = */ ss a 50000
    gd x3 = +- r 0 x2
    gd x4 = +- r 0 x1
    gd y2 = +- vc 0 x1
    gd y3 = +- vc x1 0
    gd y4 = +- b 0 x1
    path(stroke=0) M x2 b A x1 x1 cd4 cd4 L x1 y3 A x1 x1 0 -5400000 A x1 x1 cd4 -5400000 L x1 x1 A x1 x1 cd2 cd4 L x3 t A x1 x1 3cd4 cd4 L x4 y2 A x1 x1 cd2 -5400000 A x1 x1 3cd4 -5400000 L x4 y4 A x1 x1 0 cd4 Z
    path(fill=none) M x2 b A x1 x1 cd4 cd4 L x1 y3 A x1 x1 0 -5400000 A x1 x1 cd4 -5400000 L x1 x1 A x1 x1 cd2 cd4 M x3 t A x1 x1 3cd4 cd4 L x4 y2 A x1 x1 cd2 -5400000 A x1 x1 3cd4 -5400000 L x4 y4 A x1 x1 0 cd4`,
  rightBrace: `
    av adj1 = val 8333
    av adj2 = val 50000
    gd a2 = pin 0 adj2 100000
    gd q1 = +- 100000 0 a2
    gd q2 = min q1 a2
    gd q3 = */ q2 1 2
    gd maxAdj1 = */ q3 h ss
    gd a1 = pin 0 adj1 maxAdj1
    gd y1 = */ ss a1 100000
    gd y3 = */ h a2 100000
    gd y2 = +- y3 0 y1
    gd y4 = +- b 0 y1
    path(stroke=0) M l t A wd2 y1 3cd4 cd4 L hc y2 A wd2 y1 cd2 -5400000 A wd2 y1 3cd4 -5400000 L hc y4 A wd2 y1 0 cd4 Z
    path(fill=none) M l t A wd2 y1 3cd4 cd4 L hc y2 A wd2 y1 cd2 -5400000 A wd2 y1 3cd4 -5400000 L hc y4 A wd2 y1 0 cd4`,
  mathPlus: `
    av adj1 = val 23520
    gd a1 = pin 0 adj1 73490
    gd dx1 = */ w 73490 200000
    gd dy1 = */ h 73490 200000
    gd dx2 = */ ss a1 200000
    gd x1 = +- hc 0 dx1
    gd x2 = +- hc 0 dx2
    gd x3 = +- hc dx2 0
    gd x4 = +- hc dx1 0
    gd y1 = +- vc 0 dy1
    gd y2 = +- vc 0 dx2
    gd y3 = +- vc dx2 0
    gd y4 = +- vc dy1 0
    path M x1 y2 L x2 y2 L x2 y1 L x3 y1 L x3 y2 L x4 y2 L x4 y3 L x3 y3 L x3 y4 L x2 y4 L x2 y3 L x1 y3 Z`,
  mathMinus: `
    av adj1 = val 23520
    gd a1 = pin 0 adj1 100000
    gd dy1 = */ h a1 200000
    gd dx1 = */ w 73490 200000
    gd y1 = +- vc 0 dy1
    gd y2 = +- vc dy1 0
    gd x1 = +- hc 0 dx1
    gd x2 = +- hc dx1 0
    path M x1 y1 L x2 y1 L x2 y2 L x1 y2 Z`,
  mathEqual: `
    av adj1 = val 23520
    av adj2 = val 11760
    gd a1 = pin 0 adj1 36745
    gd 2a1 = */ a1 2 1
    gd mAdj2 = +- 100000 0 2a1
    gd a2 = pin 0 adj2 mAdj2
    gd dy1 = */ h a1 100000
    gd dy2 = */ h a2 200000
    gd dx1 = */ w 73490 200000
    gd y2 = +- vc 0 dy2
    gd y3 = +- vc dy2 0
    gd y1 = +- y2 0 dy1
    gd y4 = +- y3 dy1 0
    gd x1 = +- hc 0 dx1
    gd x2 = +- hc dx1 0
    path M x1 y1 L x2 y1 L x2 y2 L x1 y2 Z M x1 y3 L x2 y3 L x2 y4 L x1 y4 Z`,
  flowChartProcess: 'path(w=1, h=1) M 0 0 L 1 0 L 1 1 L 0 1 Z',
  flowChartDecision: 'path(w=2, h=2) M 0 1 L 1 0 L 2 1 L 1 2 Z',
  flowChartInputOutput: 'path(w=5, h=5) M 0 5 L 1 0 L 5 0 L 4 5 Z',
  flowChartPredefinedProcess: `
    path(stroke=0, w=1, h=1) M 0 0 L 1 0 L 1 1 L 0 1 Z
    path(fill=none, w=8, h=8) M 1 0 L 1 8 M 7 0 L 7 8
    path(fill=none, w=1, h=1) M 0 0 L 1 0 L 1 1 L 0 1 Z`,
  flowChartTerminator: 'path(w=21600, h=21600) M 3475 0 L 18125 0 A 3475 10800 3cd4 cd2 L 3475 21600 A 3475 10800 cd4 cd2 Z',
  flowChartAlternateProcess: `
    gd x2 = +- r 0 ssd6
    gd y2 = +- b 0 ssd6
    path M l ssd6 A ssd6 ssd6 cd2 cd4 L x2 t A ssd6 ssd6 3cd4 cd4 L r y2 A ssd6 ssd6 0 cd4 L ssd6 b A ssd6 ssd6 cd4 cd4 Z`,
  flowChartDocument: 'path(w=21600, h=21600) M 0 0 L 21600 0 L 21600 17322 C 10800 17322 10800 23922 0 20172 Z',
  flowChartManualInput: 'path(w=5, h=5) M 0 1 L 5 0 L 5 5 L 0 5 Z',
  flowChartManualOperation: 'path(w=5, h=5) M 0 0 L 5 0 L 4 5 L 1 5 Z',
  flowChartConnector: 'path M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z',
  flowChartMerge: 'path(w=2, h=2) M 0 0 L 2 0 L 1 2 Z',
  flowChartExtract: 'path(w=2, h=2) M 0 2 L 1 0 L 2 2 Z',
  flowChartSort: `
    path(stroke=0, w=2, h=2) M 0 1 L 1 0 L 2 1 L 1 2 Z
    path(fill=none, w=2, h=2) M 0 1 L 2 1
    path(fill=none, w=2, h=2) M 0 1 L 1 0 L 2 1 L 1 2 Z`,
};

/** The preset definitions by name, read from SRC once. */
export const PRESETS = Object.fromEntries(Object.entries(SRC).map(([k, v]) => [k, def(v)]));
// (Connectors: a straight line, as Word draws one with no bends.)
PRESETS.straightConnector1 = PRESETS.line;

/** Insert > Shapes' gallery, in Word's order: [preset, name, arrow at its end]. Only those defined above are offered. */
export const SHAPE_GALLERY = [
  ['line', 'Line'], ['line', 'Arrow', true],
  ['rect', 'Rectangle'], ['roundRect', 'Rectangle: Rounded Corners'], ['snip1Rect', 'Rectangle: Single Corner Snipped'], ['round1Rect', 'Rectangle: Single Corner Rounded'],
  ['ellipse', 'Oval'], ['triangle', 'Isosceles Triangle'], ['rtTriangle', 'Right Triangle'], ['parallelogram', 'Parallelogram'], ['trapezoid', 'Trapezoid'],
  ['diamond', 'Diamond'], ['pentagon', 'Regular Pentagon'], ['hexagon', 'Hexagon'], ['heptagon', 'Heptagon'], ['octagon', 'Octagon'], ['decagon', 'Decagon'],
  ['dodecagon', 'Dodecagon'], ['pie', 'Partial Circle'], ['chord', 'Chord'], ['teardrop', 'Teardrop'], ['frame', 'Frame'], ['halfFrame', 'Half Frame'],
  ['corner', 'L-Shape'], ['diagStripe', 'Diagonal Stripe'], ['plus', 'Cross'], ['plaque', 'Plaque'], ['can', 'Cylinder'], ['cube', 'Cube'], ['bevel', 'Rectangle: Beveled'],
  ['donut', 'Circle: Hollow'], ['noSmoking', '"No" Symbol'], ['blockArc', 'Block Arc'], ['foldedCorner', 'Rectangle: Folded Corner'], ['smileyFace', 'Smiley Face'],
  ['heart', 'Heart'], ['lightningBolt', 'Lightning Bolt'], ['sun', 'Sun'], ['moon', 'Moon'], ['cloud', 'Cloud'], ['arc', 'Arc'], ['bracketPair', 'Double Bracket'],
  ['bracePair', 'Double Brace'], ['leftBracket', 'Left Bracket'], ['rightBrace', 'Right Brace'],
  ['rightArrow', 'Arrow: Right'], ['leftArrow', 'Arrow: Left'], ['upArrow', 'Arrow: Up'], ['downArrow', 'Arrow: Down'], ['leftRightArrow', 'Arrow: Left-Right'],
  ['upDownArrow', 'Arrow: Up-Down'], ['stripedRightArrow', 'Arrow: Striped Right'], ['notchedRightArrow', 'Arrow: Notched Right'], ['homePlate', 'Arrow: Pentagon'], ['chevron', 'Arrow: Chevron'],
  ['mathPlus', 'Plus Sign'], ['mathMinus', 'Minus Sign'], ['mathMultiply', 'Multiplication Sign'], ['mathEqual', 'Equal'],
  ['flowChartProcess', 'Flowchart: Process'], ['flowChartAlternateProcess', 'Flowchart: Alternate Process'], ['flowChartDecision', 'Flowchart: Decision'],
  ['flowChartInputOutput', 'Flowchart: Data'], ['flowChartPredefinedProcess', 'Flowchart: Predefined Process'], ['flowChartDocument', 'Flowchart: Document'],
  ['flowChartTerminator', 'Flowchart: Terminator'], ['flowChartManualInput', 'Flowchart: Manual Input'], ['flowChartManualOperation', 'Flowchart: Manual Operation'],
  ['flowChartConnector', 'Flowchart: Connector'], ['flowChartExtract', 'Flowchart: Extract'], ['flowChartMerge', 'Flowchart: Merge'], ['flowChartSort', 'Flowchart: Sort'],
  ['irregularSeal1', 'Explosion: 8 Points'], ['star4', 'Star: 4 Points'], ['star5', 'Star: 5 Points'], ['star6', 'Star: 6 Points'], ['star8', 'Star: 8 Points'],
  ['ribbon2', 'Ribbon: Tilted Up'], ['wave', 'Wave'], ['doubleWave', 'Double Wave'],
  ['wedgeRectCallout', 'Speech Bubble: Rectangle'], ['wedgeRoundRectCallout', 'Speech Bubble: Rectangle with Corners Rounded'], ['wedgeEllipseCallout', 'Speech Bubble: Oval'],
];
