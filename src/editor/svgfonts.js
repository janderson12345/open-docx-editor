// Shapes, text boxes, WordArt and SmartArt are drawn as SVG pictures (src/ooxml/shapes.js), and a picture can't load
// the page's web fonts: where the document's font is missing, its text fell back to an installed one (Aptos's text in
// Calibri). So a picture shown on the page is swapped for a copy with the stand-in faces its text needs embedded
// (data: URLs of the files the page's own @font-face rules load: fonts.css, open-metric-fonts, the host's substitutes),
// once they're fetched; until then it shows as drawn. Only what's shown changes: the node keeps the picture as drawn
// (schema.js maps the copy back to it for copy and paste).

const done = new Map(); // a picture as drawn -> the copy shown (itself when it needs nothing)
const pending = new Set();
const files = new Map(); // font file URL -> Promise of its data: URL (null if it can't be had)
const installed = new Map();
const KEEP = 300;

/**
 * The picture to show for `src`: its copy with fonts embedded when that's ready, else `src` itself (the copy then
 * replaces it on the page when done; `seen(copy, src)` is told of each copy made).
 */
export function shownSrc(src, seen) {
  if (typeof document === 'undefined' || typeof src !== 'string' || !src.startsWith('data:image/svg+xml')) return src;
  const got = done.get(src);
  if (got) return got;
  if (!pending.has(src)) {
    pending.add(src);
    withFonts(src).then((out) => {
      remember(src, out || src);
      if (out && out !== src) {
        seen?.(out, src);
        for (const img of document.querySelectorAll('img')) if (img.getAttribute('src') === src) img.setAttribute('src', out);
      }
    }, () => remember(src, src)).finally(() => pending.delete(src));
  }
  return src;
}

function remember(src, out) {
  done.set(src, out);
  // (Resizing a shape draws it again at each step: only the latest are kept.)
  while (done.size > KEEP) done.delete(done.keys().next().value);
}

/** `src` with the faces its text needs embedded, or null when it needs none. */
export async function withFonts(src) {
  const svg = decode(src);
  if (!svg || !/<text\b|<tspan\b/.test(svg)) return null;
  const stacks = new Set();
  for (const m of svg.matchAll(/font-family(?:="([^"]*)"|='([^']*)'|\s*:\s*([^;"'>]+))/g)) stacks.add(unescape(m[1] ?? m[2] ?? m[3]));
  if (!stacks.size) return null;
  const bold = /font-weight(?:=["']|\s*:\s*)(bold|[6-9]00)/.test(svg);
  const italic = /font-style(?:=["']|\s*:\s*)(italic|oblique)/.test(svg);
  const text = unescape([...svg.matchAll(/>([^<]+)</g)].map((m) => m[1]).join(''));
  const rules = faceRules();
  const want = new Set();
  for (const stack of stacks) {
    for (const family of stack.split(',').map((f) => f.trim().replace(/^["']|["']$/g, '')).filter(Boolean)) {
      const key = family.toLowerCase();
      if (/^(serif|sans-serif|monospace|cursive|fantasy|system-ui)$/.test(key)) break;
      if (rules.some((r) => r.family === key)) { want.add(key); break; }
      if (isInstalled(family)) break;
    }
  }
  if (!want.size) return null;
  const faces = [];
  for (const family of want) {
    const own = rules.filter((r) => r.family === family && r.url && covers(r.range, text));
    for (const style of italic ? ['normal', 'italic'] : ['normal']) {
      for (const weight of bold ? [400, 700] : [400]) {
        // The face the browser would pick: its style if it has one, the nearest weight.
        const styled = own.filter((r) => r.italic === (style === 'italic'));
        const pool = styled.length ? styled : own;
        if (!pool.length) continue;
        const best = Math.min(...pool.map((r) => r.distance(weight)));
        for (const r of pool) if (r.distance(weight) === best && !faces.includes(r)) faces.push(r);
      }
    }
  }
  if (!faces.length) return null;
  const css = [];
  for (const r of faces) {
    const data = await file(r.url);
    if (data) css.push(`@font-face{font-family:"${r.family.replace(/["\\]/g, '')}";font-style:${r.italic ? 'italic' : 'normal'};font-weight:${r.weight};src:url(${data})${r.format ? ` format("${r.format}")` : ''}${r.range ? `;unicode-range:${r.range}` : ''}}`);
  }
  if (!css.length) return null;
  const out = svg.replace(/<svg\b[^>]*>/, (open) => `${open}<style>${css.join('')}</style>`);
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(out)}`;
}

function decode(src) {
  const comma = src.indexOf(',');
  if (comma < 0) return null;
  const head = src.slice(0, comma), body = src.slice(comma + 1);
  try {
    if (/;base64$/i.test(head)) {
      const bin = atob(body);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return new TextDecoder().decode(bytes);
    }
    return decodeURIComponent(body);
  } catch {
    return null;
  }
}

const unescape = (s) => s.replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(+d)).replace(/&#x([\da-f]+);/gi, (m, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&amp;/g, '&');

/** The page's @font-face rules: { family (lower case), url, format, weight, italic, range, distance(w) }. */
function faceRules() {
  const out = [];
  const walk = (list, base) => {
    for (const rule of list) {
      if (rule.styleSheet) { walk(rules(rule.styleSheet), rule.styleSheet.href || base); continue; }
      if (rule.cssRules && !(rule instanceof CSSFontFaceRule)) { walk(rule.cssRules, base); continue; }
      if (!(rule instanceof CSSFontFaceRule)) continue;
      const st = rule.style;
      const family = st.getPropertyValue('font-family').trim().replace(/^["']|["']$/g, '').toLowerCase();
      // Its first file (a local() source is the installed font, which the picture can use by name anyway).
      const m = /url\(\s*(?:"([^"]*)"|'([^']*)'|([^)\s]*))\s*\)(?:\s*format\(\s*["']?([^"')]+)["']?\s*\))?/.exec(st.getPropertyValue('src'));
      let url = null;
      try { url = m && new URL(m[1] ?? m[2] ?? m[3], base).href; } catch { /* none */ }
      const wv = (st.getPropertyValue('font-weight') || '400').trim();
      const [lo, hi = lo] = wv.split(/\s+/).map((v) => (v === 'bold' ? 700 : v === 'normal' ? 400 : Number(v) || 400));
      out.push({
        family, url, format: m?.[4] || null, weight: lo === hi ? lo : `${lo} ${hi}`,
        italic: /italic|oblique/.test(st.getPropertyValue('font-style')),
        range: st.getPropertyValue('unicode-range').trim() || null,
        distance: (w) => (w < lo ? lo - w : w > hi ? w - hi : 0),
      });
    }
  };
  for (const sheet of document.styleSheets) walk(rules(sheet), sheet.href || document.baseURI);
  // (The same face twice, e.g. fonts.css linked by the host and by the editor: the later one is the browser's.)
  const last = new Map();
  for (const r of out) last.set(`${r.family}|${r.weight}|${r.italic}|${r.range}`, r);
  return out.filter((r) => last.get(`${r.family}|${r.weight}|${r.italic}|${r.range}`) === r);
}

function rules(sheet) {
  try { return sheet.cssRules || []; } catch { return []; } // (another site's stylesheet, without CORS)
}

/** Whether a unicode-range takes any of the characters in `text` (no range: every one). */
function covers(range, text) {
  if (!range) return true;
  const spans = range.split(',').map((r) => {
    const v = r.trim().replace(/^u\+/i, '');
    if (v.includes('?')) return [parseInt(v.replace(/\?/g, '0'), 16), parseInt(v.replace(/\?/g, 'f'), 16)];
    const [a, b = a] = v.split('-');
    return [parseInt(a, 16), parseInt(b, 16)];
  });
  for (const ch of text) {
    const c = ch.codePointAt(0);
    if (c > 32 && spans.some(([a, b]) => c >= a && c <= b)) return true;
  }
  return false;
}

/** Whether a font of that name is installed (it draws differently from the generic families it would fall back to). */
function isInstalled(family) {
  if (installed.has(family)) return installed.get(family);
  let yes = false;
  try {
    const c = document.createElement('canvas').getContext('2d');
    const probe = 'mmmmmmmmmlliWW@#0O';
    yes = ['monospace', 'serif', 'sans-serif'].some((g) => {
      c.font = `72px ${g}`;
      const w = c.measureText(probe).width;
      c.font = `72px "${family.replace(/["\\]/g, '')}", ${g}`;
      return c.measureText(probe).width !== w;
    });
  } catch { /* none */ }
  installed.set(family, yes);
  return yes;
}

function file(url) {
  if (!files.has(url)) {
    files.set(url, fetch(url, { mode: 'cors', credentials: 'omit' })
      .then((r) => (r.ok ? r.blob() : null))
      .then((blob) => blob && new Promise((resolve) => {
        const fr = new FileReader();
        fr.onload = () => resolve(typeof fr.result === 'string' ? fr.result : null);
        fr.onerror = () => resolve(null);
        fr.readAsDataURL(blob);
      }))
      .catch(() => null));
  }
  return files.get(url);
}
