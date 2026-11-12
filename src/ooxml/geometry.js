// DrawingML geometry: a shape's outline from its guides and paths (ECMA-376 Part 1, 20.1.9), for preset shapes (by
// their definitions, presets.js) and custom ones (a:custGeom) alike. Guides are formulas over the shape's size and its
// adjust values (a:avLst, which a shape's own a:prstGeom/a:custGeom may override); paths are moveTo, lnTo, arcTo,
// cubicBezTo, quadBezTo and close, each path with its own fill (none, norm, or darkened or lightened, as a can's top)
// and stroke. Angles are in 60000ths of a degree; lengths in the shape's own units (EMU, or px: any unit, consistently).

const RAD = Math.PI / 180 / 60000;

/** The built-in guides for a w x h shape. */
function builtins(w, h) {
  const ss = Math.min(w, h);
  const g = {
    w, h, l: 0, t: 0, r: w, b: h, hc: w / 2, vc: h / 2, ss, ls: Math.max(w, h),
    cd2: 10800000, cd4: 5400000, cd8: 2700000, '3cd4': 16200000, '3cd8': 8100000, '5cd8': 13500000, '7cd8': 18900000,
  };
  for (const n of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) { g[`wd${n}`] = w / n; g[`hd${n}`] = h / n; g[`ssd${n}`] = ss / n; }
  g['3wd4'] = (3 * w) / 4; g['3wd8'] = (3 * w) / 8; g['5wd8'] = (5 * w) / 8; g['7wd8'] = (7 * w) / 8;
  g['3hd4'] = (3 * h) / 4; g['3hd8'] = (3 * h) / 8; g['5hd8'] = (5 * h) / 8; g['7hd8'] = (7 * h) / 8;
  return g;
}

// One guide formula (multiply-divide "w adj 100000", pin, at2...) evaluated against the guides so far.
function formula(f, g) {
  const [op, ...args] = f.trim().split(/\s+/);
  const v = args.map((a) => (a in g ? g[a] : Number(a)));
  const [x, y, z] = v;
  switch (op) {
    case '*/': return (x * y) / z;
    case '+-': return x + y - z;
    case '+/': return (x + y) / z;
    case '?:': return x > 0 ? y : z;
    case 'abs': return Math.abs(x);
    case 'at2': return Math.atan2(y, x) / RAD;
    case 'cat2': return x * Math.cos(Math.atan2(z, y));
    case 'cos': return x * Math.cos(y * RAD);
    case 'max': return Math.max(x, y);
    case 'min': return Math.min(x, y);
    case 'mod': return Math.sqrt(x * x + y * y + z * z);
    case 'pin': return y < x ? x : y > z ? z : y;
    case 'sat2': return x * Math.sin(Math.atan2(z, y));
    case 'sin': return x * Math.sin(y * RAD);
    case 'sqrt': return Math.sqrt(x);
    case 'tan': return x * Math.tan(y * RAD);
    case 'val': return x;
    default: return 0;
  }
}

/**
 * The guides of a definition `def` ({ av: [[name, fmla]], gd: [[name, fmla]] }) for a w x h shape, with `adjust`
 * ({ name: fmla }, the shape's own a:avLst) in place of the defaults.
 */
export function guides(def, w, h, adjust = {}) {
  const g = builtins(w, h);
  for (const [n, f] of def.av || []) g[n] = formula(adjust[n] ?? f, g);
  // (An adjust value a shape names that the definition hasn't, as "adj" for "adj1": Word reads it as the first.)
  if (adjust.adj != null && !(def.av || []).some(([n]) => n === 'adj') && def.av?.length) g[def.av[0][0]] = formula(adjust.adj, g);
  for (const [n, f] of def.gd || []) g[n] = formula(f, g);
  return g;
}

/** A point on the ellipse of radii wR, hR at DrawingML angle `ang` (measured as seen, not parametric), from its centre. */
function onEllipse(wR, hR, ang) {
  const a = ang * RAD;
  const t = Math.atan2(wR * Math.sin(a), hR * Math.cos(a));
  return [wR * Math.cos(t), hR * Math.sin(t)];
}

const num = (v) => +v.toFixed(2);

/**
 * SVG paths for definition `def` drawn in a w x h box: [{ d, fill: 'none'|'norm'|'darken'|'darkenLess'|'lighten'|
 * 'lightenLess', stroke: boolean }]. Commands: ['M', x, y], ['L', x, y], ['A', wR, hR, stAng, swAng], ['C', x1, y1,
 * x2, y2, x, y], ['Q', x1, y1, x, y], ['Z'] (each value a guide name or a number). w and h are in the definition's units
 * (EMU, as its numbers are), drawn at `scale` (px per unit); a path's own w/h are its coordinates' extent. Each path
 * also says where it starts and ends and which way it goes there ({ ends: [[x, y, dx, dy], [x, y, dx, dy]] }), for a
 * line's arrowheads.
 */
export function geometryPaths(def, w, h, adjust = {}, scale = 1) {
  const g = guides(def, w, h, adjust);
  const val = (a) => (typeof a === 'number' ? a : a in g ? g[a] : Number(a) || 0);
  return (def.path || []).map((p) => {
    const sx = p.w ? (w * scale) / p.w : scale;
    const sy = p.h ? (h * scale) / p.h : scale;
    let d = '';
    const pts = []; // [x, y] after each command, for the ends' directions
    let cx = 0, cy = 0, startX = 0, startY = 0;
    for (const c of p.cmds) {
      const [k, ...a] = c;
      const n = a.map(val);
      if (k === 'M') { cx = n[0] * sx; cy = n[1] * sy; startX = cx; startY = cy; d += `M${num(cx)} ${num(cy)}`; }
      else if (k === 'L') { cx = n[0] * sx; cy = n[1] * sy; d += `L${num(cx)} ${num(cy)}`; }
      else if (k === 'C') { d += `C${num(n[0] * sx)} ${num(n[1] * sy)} ${num(n[2] * sx)} ${num(n[3] * sy)} ${num(n[4] * sx)} ${num(n[5] * sy)}`; cx = n[4] * sx; cy = n[5] * sy; }
      else if (k === 'Q') { d += `Q${num(n[0] * sx)} ${num(n[1] * sy)} ${num(n[2] * sx)} ${num(n[3] * sy)}`; cx = n[2] * sx; cy = n[3] * sy; }
      else if (k === 'Z') { d += 'Z'; cx = startX; cy = startY; }
      else if (k === 'A') {
        const wR = n[0] * sx, hR = n[1] * sy, st = n[2], sw = n[3];
        if (!wR || !hR || !sw) continue;
        const [px0, py0] = onEllipse(wR, hR, st);
        const ox = cx - px0, oy = cy - py0; // the ellipse's centre
        // (A whole turn or more in two halves: SVG can't draw an arc back to its own start.)
        const parts = Math.abs(sw) >= 21600000 - 1 ? 2 : 1;
        for (let i = 1; i <= parts; i++) {
          const [ex, ey] = onEllipse(wR, hR, st + (sw * i) / parts);
          const step = sw / parts;
          cx = ox + ex; cy = oy + ey;
          d += `A${num(wR)} ${num(hR)} 0 ${Math.abs(step) > 10800000 ? 1 : 0} ${step > 0 ? 1 : 0} ${num(cx)} ${num(cy)}`;
        }
      }
      pts.push([cx, cy]);
    }
    const dir = (a, b) => { const dx = a[0] - b[0], dy = a[1] - b[1], l = Math.hypot(dx, dy) || 1; return [dx / l, dy / l]; };
    const ends = pts.length > 1 ? [[...pts[0], ...dir(pts[0], pts[1])], [...pts[pts.length - 1], ...dir(pts[pts.length - 1], pts[pts.length - 2])]] : null;
    return { d, fill: p.fill || 'norm', stroke: p.stroke !== false, ends };
  });
}

// --- Adjust handles (a:ahLst): the yellow dots that change a shape's adjust values ------------------------------------

/**
 * Where a w x h shape's adjust handles are (in its own units, as w and h): [{ x, y }], one for each of the definition's
 * handles (def.ah: { kind: 'xy'|'polar', gdRefX, minX, ..., x, y }), with the shape's own adjust values.
 */
export function adjustHandles(def, w, h, adjust = {}) {
  if (!def?.ah?.length) return [];
  const g = guides(def, w, h, adjust);
  const val = (a) => (a in g ? g[a] : Number(a) || 0);
  return def.ah.map((a) => ({ x: val(a.x), y: val(a.y) }));
}

/**
 * The adjust values that put handle `i` of a w x h shape nearest (x, y) (its own units), within the handle's limits:
 * { name: 'val N' } for the values it moves (a:avLst's gd fmla). A handle's position is a guide computed from the
 * values, which can't in general be solved backwards, so each value is searched for (within its limits, the others
 * held), the closest found then narrowed down; twice over, for a handle moving two that interact (a polar handle's
 * angle and radius).
 */
export function dragAdjust(def, w, h, adjust, i, x, y) {
  const a = def?.ah?.[i];
  if (!a) return {};
  const cur = { ...adjust };
  const valueOf = (name) => { const g = guides(def, w, h, cur); return g[name]; };
  const limit = (f, fallback) => { if (f == null) return fallback; const g = guides(def, w, h, cur); return f in g ? g[f] : Number(f); };
  const refs = a.kind === 'xy'
    ? [[a.gdRefX, a.minX, a.maxX], [a.gdRefY, a.minY, a.maxY]]
    : [[a.gdRefAng, a.minAng, a.maxAng], [a.gdRefR, a.minR, a.maxR]];
  const live = refs.filter(([ref]) => ref);
  const miss = (name, v) => {
    cur[name] = `val ${v}`;
    const g = guides(def, w, h, cur);
    const px = a.x in g ? g[a.x] : Number(a.x) || 0, py = a.y in g ? g[a.y] : Number(a.y) || 0;
    // (An xy handle's value moves it one way only: measured that way, so the other doesn't pull it.)
    if (a.kind === 'xy') return name === a.gdRefX ? Math.abs(px - x) : Math.abs(py - y);
    return Math.hypot(px - x, py - y);
  };
  for (let round = 0; round < (live.length > 1 ? 3 : 1); round++) {
    for (const [ref, minF, maxF] of live) {
      const now = valueOf(ref) ?? 0;
      // (Unlimited values, as a callout's pointer, searched within a wide span around where they are.)
      let lo = Math.max(limit(minF, -1e6), now - 2e6), hi = Math.min(limit(maxF, 1e6), now + 2e6);
      if (!(hi >= lo)) [lo, hi] = [hi, lo];
      const N = 96;
      let best = lo, bestMiss = Infinity;
      for (let k = 0; k <= N; k++) {
        const v = lo + ((hi - lo) * k) / N;
        const m = miss(ref, v);
        if (m < bestMiss) { bestMiss = m; best = v; }
      }
      // Narrowed down around the best (golden-section search on the miss).
      let s0 = Math.max(lo, best - (hi - lo) / N), s1 = Math.min(hi, best + (hi - lo) / N);
      for (let k = 0; k < 40 && s1 - s0 > 0.5; k++) {
        const m1 = s0 + (s1 - s0) * 0.382, m2 = s0 + (s1 - s0) * 0.618;
        if (miss(ref, m1) <= miss(ref, m2)) s1 = m2; else s0 = m1;
      }
      const v = Math.round((s0 + s1) / 2);
      cur[ref] = `val ${miss(ref, v) <= bestMiss + 1e-6 ? v : Math.round(best)}`;
    }
  }
  return Object.fromEntries(live.map(([ref]) => [ref, cur[ref]]));
}

// --- a:custGeom / a definition's XML, read into the same form ---------------------------------------------------------

const kidsOf = (el) => (el ? [...el.childNodes].filter((n) => n.nodeType === 1) : []);
const kidOf = (el, name) => kidsOf(el).find((n) => n.localName === name) || null;

/** A geometry element (a:custGeom, or a preset's definition element) as { av, gd, path }. */
export function readGeometry(el) {
  const gds = (list) => kidsOf(list).filter((n) => n.localName === 'gd').map((n) => [n.getAttribute('name'), n.getAttribute('fmla')]);
  const pt = (n) => { const p = kidOf(n, 'pt'); return [p?.getAttribute('x') ?? 0, p?.getAttribute('y') ?? 0]; };
  const path = kidsOf(kidOf(el, 'pathLst')).map((p) => ({
    w: Number(p.getAttribute('w')) || 0,
    h: Number(p.getAttribute('h')) || 0,
    fill: p.getAttribute('fill') || 'norm',
    stroke: p.getAttribute('stroke') !== '0' && p.getAttribute('stroke') !== 'false',
    cmds: kidsOf(p).map((c) => {
      const pts = kidsOf(c).filter((n) => n.localName === 'pt').map((n) => [n.getAttribute('x'), n.getAttribute('y')]).flat();
      switch (c.localName) {
        case 'moveTo': return ['M', ...pt(c)];
        case 'lnTo': return ['L', ...pt(c)];
        case 'cubicBezTo': return ['C', ...pts];
        case 'quadBezTo': return ['Q', ...pts];
        case 'arcTo': return ['A', c.getAttribute('wR'), c.getAttribute('hR'), c.getAttribute('stAng'), c.getAttribute('swAng')];
        case 'close': return ['Z'];
        default: return null;
      }
    }).filter(Boolean),
  }));
  const ah = kidsOf(kidOf(el, 'ahLst')).filter((n) => n.localName === 'ahXY' || n.localName === 'ahPolar').map((n) => {
    const pos = kidOf(n, 'pos');
    return { kind: n.localName === 'ahXY' ? 'xy' : 'polar', ...Object.fromEntries([...n.attributes].map((at) => [at.name, at.value])), x: pos?.getAttribute('x') ?? 0, y: pos?.getAttribute('y') ?? 0 };
  });
  return { av: gds(kidOf(el, 'avLst')), gd: gds(kidOf(el, 'gdLst')), path, ah };
}

/** A shape's own adjust values (its a:prstGeom or a:custGeom's a:avLst): { name: fmla }. */
export function adjustOf(el) {
  return Object.fromEntries(kidsOf(kidOf(el, 'avLst')).filter((n) => n.localName === 'gd').map((n) => [n.getAttribute('name'), n.getAttribute('fmla')]));
}
