// Open/save a .docx as a ProseMirror document. A DocxSession keeps the
// original package so every part we don't edit is written back unchanged.
import { Node } from 'prosemirror-model';
import { DocxPackage } from './package.js';
import { REL_TYPES, NS, kids, parseXml, serializeXml } from './xml.js';
import { readDocument, readStory } from './reader.js';
import { writeDocument, writeStory, writeBlocks } from './writer.js';
import { readStyles, readTheme, readNumbering, stylesToCss, styleMenu, listLevel, resolveStyle } from './styles.js';
import { readComments, writeComments } from './comments.js';
import { loadNotes, saveNotes } from './notes.js';
import { chartSvg } from './chart.js';
import { readChart, workbookRel, updateWorkbook } from './chartedit.js';
import { shapesSvg } from './shapes.js';
import { fontAlts } from './fonts.js';
import { setFontAlts } from './props.js';
import { isEmf, emfToSvg } from './emf.js';
import { blankDocxBytes, bulletAbstract, decimalAbstract, abstractFor } from './template.js';
import { schema } from '../editor/schema.js';
import { TWIPS_PER_PX, emuToPx, pxToEmu, twipsToPx } from './units.js';

// A text box's text (its opening tag, then its contents).
const TXBX = /(<w:txbxContent\b[^>]*?(?<!\/)>)([\s\S]*?)<\/w:txbxContent>/g;

export class DocxSession {
  /**
   * @param {ArrayBuffer|Uint8Array} bytes
   * @param {{ imageUrl?: (bytes: Uint8Array, mime: string) => string }} opts
   */
  constructor(bytes, opts = {}) {
    this.pkg = DocxPackage.fromBytes(bytes);
    this.opts = opts;
    setFontAlts(fontAlts(this.pkg)); // (before the styles: their fonts' stacks and line heights use the stand-ins)
    this.theme = readTheme(this.pkg);
    this.styles = readStyles(this.pkg);
    // The document language (styles.xml docDefaults), so screen readers pronounce it correctly.
    this.lang = (/<w:rPrDefault>[\s\S]*?<w:lang\b[^>]*w:val="([^"]+)"/.exec(this.pkg.text("word/styles.xml") || "") || [])[1] || null;
    // Word version the layout follows (15 = Word 2013+); older modes lay some things out differently.
    const settings = this.pkg.partByType(REL_TYPES.settings);
    const mode = settings && /<w:compatSetting\b[^>]*w:name="compatibilityMode"[^>]*w:val="(\d+)"|<w:compatSetting\b[^>]*w:val="(\d+)"[^>]*w:name="compatibilityMode"/.exec(this.pkg.text(settings) || '');
    this.compatMode = mode ? Number(mode[1] || mode[2]) : 12;
    // The compatibility rules it names (w:compat's children, by their names: splitPgBreakAndParaMark...), each a
    // Word version's layout the document asks for; the layout follows those it knows (docs/LAYOUT.md).
    const compat = settings && /<w:compat>([\s\S]*?)<\/w:compat>/.exec(this.pkg.text(settings) || '')?.[1];
    this.compat = new Set(compat ? [...compat.matchAll(/<w:(\w+)\b([^>]*)\/?>/g)].filter(([, n, a]) => n !== 'compatSetting' && !/w:val="(?:0|false|off)"/.test(a)).map(([, n]) => n) : []);
    // Track Changes on when the document was saved (the editor updates this as it's toggled).
    this.settingsPath = settings;
    this.trackRevisions = this._trackRevisions0 = !!settings && /<w:trackRevisions(?![^>]*w:val="(?:0|false|off)")[^>]*\/?>/.test(this.pkg.text(settings) || '');
    const dts = settings && /<w:defaultTabStop\b[^>]*w:val="(\d+)"/.exec(this.pkg.text(settings) || '');
    this.defaultTabStop = dts ? Number(dts[1]) : 720;
    this.numbering = readNumbering(this.pkg);
    this.pictureBullets = this.picBullets(); // (the session's, not numbering's: numbering is read again as lists are added)
    this.originalXml = this.pkg.text(this.pkg.mainPath);
    this._resolvers = {
      resolveImage: (rId) => {
        const m = this.pkg.media(rId);
        return m && opts.imageUrl ? opts.imageUrl(...shown(m.bytes, m.mime)) : null;
      },
      resolveLink: (rId) => {
        const rel = this.pkg.relById(rId);
        return rel && rel.external ? rel.rawTarget : null;
      },
      resolveChart: (rId, w, h) => this.chartUrl(this.pkg.relById(rId), w, h),
      resolveShape: (el, w, h) => this.shapeUrl(el, (rId) => this.pkg.relById(rId), w, h),
    };
    const json = readDocument(parseXml(this.originalXml), this._resolvers);
    this.doc = Node.fromJSON(schema, json);
    this.comments = readComments(this.pkg);
    this.notes = loadNotes(this.pkg);
    this.commentsDirty = false;
  }

  /**
   * Picture bullets (w:numPicBullet): id -> an SVG data URL of the picture at the bullet's size (points), which a list
   * label shows as its content, on the baseline as in Word. (A data URL, so it sizes itself, and prints.)
   */
  picBullets() {
    const out = new Map();
    const path = this.pkg.partByType(REL_TYPES.numbering);
    const doc = path && this.pkg.xml(path);
    if (!doc) return out;
    const rels = this.pkg.readRels(path);
    // (From the parsed part, each bullet's own XML: a pattern over the whole part was slow on a crafted one.)
    for (const el of kids(doc.documentElement, 'numPicBullet')) {
      const xml = serializeXml(el);
      const num = el.getAttributeNS(NS.w, 'numPicBulletId');
      const id = /<v:imagedata\b[^>]*r:id="([^"]+)"/.exec(xml)?.[1] || /<a:blip\b[^>]*r:embed="([^"]+)"/.exec(xml)?.[1];
      const rel = id && rels.find((r) => r.id === id);
      const raw = rel && !rel.external && this.pkg.files[rel.target];
      // (A bullet is a small picture: one over 512 KB is left out, the level's own bullet character shown instead.)
      if (!/^\d+$/.test(num || '') || !raw || raw.length > 512 * 1024) continue;
      // Its size: the VML shape's style (points), or the drawing's extent (EMU); else 9pt, Word's default. At most 72pt.
      const pt = (k) => Number(new RegExp(`${k}:([\\d.]+)pt`).exec(xml)?.[1]);
      const ext = /<wp:extent\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(xml);
      const fit = (v) => Math.min(72, Math.max(1, v || 9));
      const w = fit(pt('width') || (ext ? Number(ext[1]) / 12700 : 9)), h = fit(pt('height') || (ext ? Number(ext[2]) / 12700 : 9));
      const [bytes, mime] = shown(raw, /\.jpe?g$/i.test(rel.target) ? 'image/jpeg' : /\.gif$/i.test(rel.target) ? 'image/gif' : 'image/png');
      let bin = '';
      for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}pt" height="${h}pt" viewBox="0 0 ${w} ${h}"><image href="data:${mime};base64,${btoa(bin)}" width="${w}" height="${h}" preserveAspectRatio="none"/></svg>`;
      out.set(num, 'data:image/svg+xml;base64,' + btoa(svg));
    }
    return out;
  }

  /** The picture bullets as CSS, one rule each, which list paragraphs take by class (pararules.js). */
  picBulletCss(scope) {
    return [...this.pictureBullets].map(([id, url]) => `${scope} p.odx-pb-${id}{--odx-picbullet:url("${url}")}`).join('\n');
  }

  /** A chart part drawn as an SVG data URL (null if it can't be drawn). */
  chartUrl(rel, width, height) {
    const xml = rel && !rel.external && this.pkg.text(rel.target);
    return xml ? this.chartUrlFromXml(xml, width, height) : null;
  }

  /** The chart part a body drawing shows: { path, xml } (an edited chart's XML is the node's own, chartedit.js). */
  chartPart(drawingXml) {
    const id = /<c:chart\b[^>]*r:id="([^"]+)"/.exec(drawingXml || '')?.[1];
    const rel = id && this.pkg.relById(id);
    // (As opened, though a save has since written an edit into it: an undone edit's chart shows the old data.)
    return rel && !rel.external && this.pkg.has(rel.target) ? { path: rel.target, xml: this._chartOrig?.get(rel.target)?.xml ?? this.pkg.text(rel.target) } : null;
  }

  /** Chart XML drawn as an SVG data URL (null if it can't be drawn). */
  chartUrlFromXml(xml, width, height) {
    try {
      const svg = chartSvg(xml, { width, height, font: this.theme.minor, colors: this.theme.colors });
      return svg && 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    } catch {
      return null;
    }
  }

  /** Shapes and groups drawn as an SVG data URL; their pictures are inlined as data URLs. */
  shapeUrl(drawing, relOf, width, height) {
    this._dataUrls ||= new Map();
    const image = (rId) => {
      const rel = rId && relOf(rId);
      if (!rel || rel.external) return null;
      if (!this._dataUrls.has(rel.target)) {
        const bytes = this.pkg.files[rel.target];
        const [b, mime] = bytes ? shown(bytes, /\.jpe?g$/i.test(rel.target) ? 'image/jpeg' : /\.gif$/i.test(rel.target) ? 'image/gif' : 'image/png') : [];
        this._dataUrls.set(rel.target, b ? `data:${mime};base64,${base64(b)}` : null);
      }
      return this._dataUrls.get(rel.target);
    };
    try {
      // Text box paragraphs take their style's (and the document defaults') formatting, as in the body.
      const para = (id) => {
        const sid = id || this.styles.defaultParagraph;
        const st = sid ? resolveStyle(this.styles.styles, sid) : { pPr: {}, rPr: {} };
        const rPr = { ...this.styles.defaults.rPr, ...st.rPr };
        const font = rPr.font || (rPr.themeFont && this.theme[rPr.themeFont.startsWith('major') ? 'major' : 'minor']) || this.theme.minor;
        return { pPr: { ...this.styles.defaults.pPr, ...st.pPr }, rPr, font };
      };
      const char = (id) => (id ? resolveStyle(this.styles.styles, id).rPr : {});
      const svg = shapesSvg(drawing, { image, colors: this.theme.colors, font: this.theme.minor, para, char, theme: this.theme }, width, height);
      return svg && 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    } catch {
      return null;
    }
  }

  static blank(opts) {
    return new DocxSession(blankDocxBytes(), opts);
  }

  /** CSS rules for the document's styles, scoped under `scope`. */
  css(scope) {
    return stylesToCss(scope, this.styles, this.theme);
  }

  styleMenu() {
    return styleMenu(this.styles);
  }

  /** Page geometry in px (96 dpi) from the final section. */
  page() {
    const s = this.doc.attrs.sectPr || '';
    const num = (re, d) => {
      const m = re.exec(s);
      return m ? twipsToPx(m[1]) : d;
    };
    const g = this.gutter(s);
    return {
      width: num(/<w:pgSz[^>]*\sw:w="(\d+)"/, 816),
      height: num(/<w:pgSz[^>]*\sw:h="(\d+)"/, 1056),
      top: num(/<w:pgMar[^>]*\sw:top="(-?\d+)"/, 96) + g.top,
      right: num(/<w:pgMar[^>]*\sw:right="(\d+)"/, 96),
      bottom: num(/<w:pgMar[^>]*\sw:bottom="(-?\d+)"/, 96),
      left: num(/<w:pgMar[^>]*\sw:left="(\d+)"/, 96) + g.left,
      firstTop: this.firstSectionTop(),
    };
  }

  /** A section's gutter (px): added to the left margin, or to the top one when settings say gutterAtTop. */
  gutter(sectPr) {
    const m = /<w:pgMar[^>]*\sw:gutter="(\d+)"/.exec(sectPr || '');
    const px = m ? twipsToPx(m[1]) : 0;
    if (!px) return { top: 0, left: 0 };
    this._gutterTop ??= !!this.settingsPath && /<w:gutterAtTop(?![^>]*w:val="(?:0|false|off)")/.test(this.pkg.text(this.settingsPath) || '');
    return this._gutterTop ? { top: px, left: 0 } : { top: 0, left: px };
  }

  /** Top margin (px) of the first section, where the first page starts. */
  firstSectionTop() {
    let xml = null;
    this.doc.forEach((node) => {
      if (xml == null) xml = (node.attrs.extra || []).find((e) => e[0] === 'sectPr')?.[1] ?? null;
    });
    const sect = xml ?? this.doc.attrs.sectPr ?? '';
    const m = /<w:pgMar[^>]*\sw:top="(-?\d+)"/.exec(sect);
    return m ? Math.abs(Number(m[1])) / TWIPS_PER_PX + this.gutter(sect).top : null;
  }

  /** A header or footer part as an editor doc (display only for now); targets resolve against the part's own rels. */
  story(rId) {
    this._stories ||= new Map();
    if (this._stories.has(rId)) return this._stories.get(rId);
    const rel = this.pkg.relById(rId);
    const xml = rel && this.pkg.xml(rel.target);
    let doc = null;
    if (xml) {
      const rels = this.pkg.readRels(rel.target);
      const relOf = (id) => rels.find((r) => r.id === id);
      doc = Node.fromJSON(schema, readStory(xml, {
        resolveImage: (id) => {
          const r = relOf(id);
          const bytes = r && !r.external && this.pkg.files[r.target];
          return bytes && this.opts.imageUrl ? this.opts.imageUrl(...shown(bytes, /\.jpe?g$/i.test(r.target) ? 'image/jpeg' : 'image/png')) : null;
        },
        resolveLink: (id) => relOf(id)?.rawTarget ?? null,
        resolveChart: (id, w, h) => this.chartUrl(relOf(id), w, h),
        resolveShape: (el, w, h) => this.shapeUrl(el, relOf, w, h),
      }));
    }
    this._stories.set(rId, doc);
    return doc;
  }

  /** Page color (w:background) as 'RRGGBB', or null. */
  get pageColor() {
    const m = /<w:background\b[^>]*w:color="([0-9A-Fa-f]{6})"/.exec(this.originalXml);
    return m ? m[1] : null;
  }

  /** Set or (null) remove the page color; Word shows it when settings has w:displayBackgroundShape. */
  setPageColor(hex) {
    let xml = this.originalXml.replace(/<w:background\b[^>]*\/>|<w:background\b[^>]*>[\s\S]*?<\/w:background>/, '');
    if (hex) xml = xml.replace(/(<w:document\b[^>]*>)/, `$1<w:background w:color="${hex}"/>`);
    this.originalXml = xml;
    if (hex && this.settingsPath) {
      const st = this.pkg.text(this.settingsPath);
      if (!/<w:displayBackgroundShape/.test(st)) this.pkg.setText(this.settingsPath, st.replace(/(<w:settings\b[^>]*>)/, '$1<w:displayBackgroundShape/>'));
    }
  }

  /** Replace a header or footer's content (saved with the document). */
  setStory(rId, doc) {
    this._stories.set(rId, doc);
    this._hfCache?.clear();
    (this._dirtyStories ||= new Set()).add(rId);
  }

  /** A new, empty header or footer part; returns its rId (the caller adds the reference to a w:sectPr). */
  addStory(kind) {
    const tag = kind === 'header' ? 'hdr' : 'ftr';
    let n = 1;
    while (this.pkg.files[`${this.pkg.mainDir}${kind}${n}.xml`]) n++;
    const name = `${kind}${n}.xml`;
    this.pkg.setText(`${this.pkg.mainDir}${name}`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${tag} xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:p><w:pPr><w:pStyle w:val="${kind === 'header' ? 'Header' : 'Footer'}"/></w:pPr></w:p></w:${tag}>`);
    this.pkg.addOverride(`${this.pkg.mainDir}${name}`, `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`);
    const rId = this.pkg.addRel(`http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}`, name);
    this.onAddStory?.({ rId, kind, name, xml: this.pkg.text(`${this.pkg.mainDir}${name}`) }); // (shared: src/collab.js)
    return rId;
  }

  /** A header or footer part another editor of a shared document added (src/collab.js), under the same rId. */
  addSharedStory({ rId, kind, name, xml }) {
    if (this.pkg.relById(rId)) return;
    let file = name;
    // (The name may be taken here by a part this editor added at the same time: another one, the rId is what counts.)
    for (let n = 1; this.pkg.files[`${this.pkg.mainDir}${file}`]; n++) file = `${kind}-shared${n}.xml`;
    this.pkg.setText(`${this.pkg.mainDir}${file}`, xml);
    this.pkg.addOverride(`${this.pkg.mainDir}${file}`, `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`);
    this.pkg.addRel(`http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}`, file, false, this.pkg.mainPath, rId);
    this._hfCache?.clear();
  }

  /** True when Word protects the document so only its form fields can be filled in. */
  get formsProtected() {
    const set = (this.settingsPath && this.pkg.text(this.settingsPath)) || '';
    const p = /<w:documentProtection\b[^>]*>/.exec(set)?.[0] || '';
    return /w:edit="forms"/.test(p) && /w:enforcement="(1|true|on)"/.test(p);
  }

  /** Word's grey shading of form fields on screen is on unless w:doNotShadeFormData. */
  get shadeFormData() {
    return !/<w:doNotShadeFormData(?![^>]*w:val="(?:0|false|off)")/.test((this.settingsPath && this.pkg.text(this.settingsPath)) || '');
  }

  /** "Different Odd & Even Pages" (w:evenAndOddHeaders in settings.xml), for every section. */
  get evenAndOddHeaders() {
    return !!this.settingsPath && /<w:evenAndOddHeaders(?![^>]*w:val="(?:0|false|off)")/.test(this.pkg.text(this.settingsPath) || '');
  }

  set evenAndOddHeaders(on) {
    if (!this.settingsPath || on === this.evenAndOddHeaders) return;
    let st = this.pkg.text(this.settingsPath).replace(/<w:evenAndOddHeaders\b[^>]*\/>/g, '');
    if (on) {
      // Its place in w:settings: before the first of these that is present, else at the end.
      const after = /<w:(?:bookFoldRevPrinting|bookFoldPrinting|bookFoldPrintingSheets|drawingGrid\w*|display\w*DrawingGridEvery|doNotUseMarginsForDrawingGridOrigin|doNotShadeFormData|noPunctuationKerning|characterSpacingControl|printTwoOnOne|strictFirstAndLastChars|noLineBreaks\w*|savePreviewPicture|doNotValidateAgainstSchema|saveInvalidXml|ignoreMixedContent|alwaysShowPlaceholderText|doNotDemarcateInvalidXml|saveXmlDataOnly|useXSLTWhenSaving|saveThroughXslt|showXMLTags|alwaysMergeEmptyNamespace|updateFields|hdrShapeDefaults|footnotePr|endnotePr|compat|docVars|rsids|mathPr|attachedSchema|themeFontLang|clrSchemeMapping|doNotIncludeSubdocsInStats|doNotAutoCompressPictures|forceUpgrade|captions|readModeInkLockDown|smartTagType|schemaLibrary|shapeDefaults|doNotEmbedSmartTags|decimalSymbol|listSeparator)\b/.exec(st);
      st = after ? st.slice(0, after.index) + '<w:evenAndOddHeaders/>' + st.slice(after.index) : st.replace(/<\/w:settings>/, '<w:evenAndOddHeaders/></w:settings>');
    }
    this.pkg.setText(this.settingsPath, st);
    this._hfCache = null;
  }

  /**
   * Headers and footers of each section, given each section's w:sectPr in order.
   * A section without a reference of some type inherits the previous section's,
   * as Word does. Equal results are the same object, so callers can compare them.
   */
  headerFooters(sectPrs) {
    this._hfCache ||= new Map();
    const even = this.evenAndOddHeaders;
    let refs = {};
    return sectPrs.map((s = '') => {
      refs = { ...refs };
      for (const m of s.matchAll(/<w:(header|footer)Reference\b[^>]*>/g)) {
        const type = /w:type="(\w+)"/.exec(m[0])?.[1] || 'default';
        const id = /r:id="([^"]+)"/.exec(m[0]);
        if (id) refs[`${m[1]}.${type}`] = id[1];
      }
      const dist = (a) => {
        const m = new RegExp(`<w:pgMar[^>]*\\sw:${a}="(\\d+)"`).exec(s);
        return m ? twipsToPx(m[1]) : 48;
      };
      const titlePg = /<w:titlePg(?![^>]*w:val="(?:0|false|off)")/.test(s);
      const key = JSON.stringify([refs, dist('header'), dist('footer'), titlePg, even]);
      if (!this._hfCache.has(key)) {
        const out = { headerDist: dist('header'), footerDist: dist('footer'), titlePg, even };
        for (const kind of ['header', 'footer']) {
          out[kind] = {};
          for (const type of ['default', 'first', 'even']) {
            const id = refs[`${kind}.${type}`];
            if (id) out[kind][type] = this.story(id);
          }
        }
        this._hfCache.set(key, out);
      }
      return this._hfCache.get(key);
    });
  }

  /** Header and footer of the final section (default type). */
  headerFooter() {
    const hf = this.headerFooters([this.doc.attrs.sectPr])[0];
    return { header: hf.header.default, footer: hf.footer.default, headerDist: hf.headerDist, footerDist: hf.footerDist };
  }

  /** A paragraph-style property after basedOn inheritance (default paragraph style when styleId is null). */
  styleProp(styleId, prop) {
    const id = styleId || this.styles.defaultParagraph;
    if (!id) return undefined;
    this._styleCache ||= new Map();
    if (!this._styleCache.has(id)) this._styleCache.set(id, resolveStyle(this.styles.styles, id));
    return this._styleCache.get(id).pPr[prop];
  }

  /** A run property (rPr) of a paragraph style, with the document defaults under it. */
  /** A character style's run properties (w:rStyle), with its basedOn chain. */
  charStyle(id) {
    return id ? resolveStyle(this.styles.styles, id).rPr : {};
  }

  styleRunProp(styleId, prop) {
    const id = styleId || this.styles.defaultParagraph;
    this._styleCache ||= new Map();
    if (id && !this._styleCache.has(id)) this._styleCache.set(id, resolveStyle(this.styles.styles, id));
    return (id && this._styleCache.get(id).rPr?.[prop]) ?? this.styles.defaults.rPr[prop];
  }

  /** numbering level definition for (numId, ilvl), or null */
  listLevel(numId, ilvl) {
    return listLevel(this.numbering, numId, ilvl);
  }

  /**
   * A numId for a bullet or numbered list, reusing an existing definition
   * when there is one and adding one to numbering.xml when not.
   * @param {'bullet'|'decimal'} kind
   */
  listId(kind) {
    for (const [numId, n] of this.numbering.nums) {
      const levels = this.numbering.abstract.get(n.abstractId);
      const fmt = levels && levels[0] && levels[0].fmt;
      if (kind === 'bullet' ? fmt === 'bullet' : fmt === 'decimal') return numId;
    }
    return this._addList((absId) => (kind === 'bullet' ? bulletAbstract : decimalAbstract)(absId));
  }

  /** A list (numId) with these levels ([numFmt, lvlText, font, left, hanging] x 9): an existing one when it matches, else a new one. */
  listFor(levels) {
    const depth = levels.some((l, i) => i && l[1].includes('%1')) ? 3 : 1;
    for (const [numId, n] of this.numbering.nums) {
      const have = this.numbering.abstract.get(n.abstractId);
      if (have && levels.slice(0, depth).every((l, i) => have[i] && have[i].fmt === l[0] && have[i].text === l[1])) return numId;
    }
    return this._addList((absId) => abstractFor(absId, levels));
  }

  _addList(abstractXml) {
    let path = this.pkg.partByType(REL_TYPES.numbering);
    if (!path) {
      path = this.pkg.mainDir + 'numbering.xml';
      this.pkg.setText(path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="${NS.w}"/>`);
      this.pkg.addRel(REL_TYPES.numbering, 'numbering.xml');
      this.pkg.addOverride(path, 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
    }
    const doc = this.pkg.xml(path);
    const root = doc.documentElement;
    const maxId = (els, a) => els.reduce((m, e) => Math.max(m, Number(e.getAttributeNS(NS.w, a)) || 0), 0);
    const absId = maxId(kids(root, 'abstractNum'), 'abstractNumId') + 1;
    const numId = String(maxId(kids(root, 'num'), 'numId') + 1);
    const frag = parseXml(`<x xmlns:w="${NS.w}">${abstractXml(absId)}<w:num w:numId="${numId}"><w:abstractNumId w:val="${absId}"/></w:num></x>`);
    const [abs, num] = kids(frag.documentElement);
    // All abstractNum elements must precede the num elements.
    root.insertBefore(doc.importNode(abs, true), kids(root, 'num')[0] || null);
    root.appendChild(doc.importNode(num, true));
    this.pkg.setText(path, serializeXml(doc));
    this.numbering = readNumbering(this.pkg);
    return numId;
  }

  /**
   * Add a picture to the package. Returns the w:drawing run XML for an inline
   * picture of width x height px (as Word writes it), ready for an image node.
   */
  addImage(bytes, mime, width, height) {
    const ext = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/svg+xml': 'svg' }[mime] || 'png';
    let n = 1;
    while (this.pkg.files[`${this.pkg.mainDir}media/odx-image${n}.${ext}`]) n++;
    const name = `odx-image${n}.${ext}`;
    this.pkg.files[`${this.pkg.mainDir}media/${name}`] = bytes;
    this.pkg.addOverride(`${this.pkg.mainDir}media/${name}`, mime);
    const rId = this.pkg.addRel(REL_TYPES.image, `media/${name}`);
    this.onAddImage?.({ rId, name, mime, bytes }); // (shared: src/collab.js)
    this._picId = (this._picId || 1000) + 1;
    const id = this._picId;
    const xml = '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
      + `<wp:extent cx="${pxToEmu(width)}" cy="${pxToEmu(height)}"/><wp:docPr id="${id}" name="Picture ${id}"/>`
      + '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
      + '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
      + `<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>`
      + `<pic:blipFill><a:blip r:embed="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
      + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${pxToEmu(width)}" cy="${pxToEmu(height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
      + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    return { xml, src: this.opts.imageUrl ? this.opts.imageUrl(bytes, mime) : null };
  }

  /** A picture another editor of a shared document added (src/collab.js), under the same rId. */
  addSharedImage({ rId, name, mime, bytes }) {
    if (this.pkg.relById(rId)) return;
    let file = name;
    for (let n = 1; this.pkg.files[`${this.pkg.mainDir}media/${file}`]; n++) file = name.replace(/(\.\w+)$/, `-shared${n}$1`);
    this.pkg.files[`${this.pkg.mainDir}media/${file}`] = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    this.pkg.addOverride(`${this.pkg.mainDir}media/${file}`, mime);
    this.pkg.addRel(REL_TYPES.image, `media/${file}`, false, this.pkg.mainPath, rId);
  }

  /** Hyperlink targets for the main document part, one relationship per address. */
  _linkRel(href) {
    const links = (this._links ||= new Map());
    if (!links.has(href)) links.set(href, this.pkg.addRel(REL_TYPES.hyperlink, href, true));
    return links.get(href);
  }

  /** Main-document XML read as a small story (text box contents, a lone run), with the document's namespaces. */
  _fragment(inner, body = false) {
    const open = /<w:document\b[^>]*?>/.exec(this.originalXml)[0].replace(/^<w:document/, '<w:hdr').replace(/\s*\/>$/, '>');
    // (On the body's page when it's for the body: a box as wide as the column is a band there, not a float.)
    const s0 = body ? this.doc.attrs.sectPr : null;
    const sect = s0 ? (/^<w:sectPr\b[^>]*xmlns:w=/.test(s0) ? s0 : s0.replace(/^<w:sectPr\b/, `<w:sectPr xmlns:w="${NS.w}"`)) : null;
    return Node.fromJSON(schema, readStory(parseXml(`${open}${inner}</w:hdr>`), this._resolvers, sect));
  }

  /**
   * The text boxes in a drawing (an image node's run XML), in order: where each sits in the drawing
   * (fractions of its width and height), its inner margins in px, and which copies of the text are its
   * own (the shape's, then the matching one in Word's VML fallback). Null when the copies don't match up.
   */
  textBoxes(runXml) {
    const all = [...runXml.matchAll(TXBX)].map((m) => m[2]);
    if (!all.length) return null;
    const open = /<w:document\b[^>]*?>/.exec(this.originalXml)[0].replace(/\s*\/>$/, '>');
    const els = [...parseXml(`${open}<w:body><w:p>${runXml}</w:p></w:body></w:document>`).getElementsByTagNameNS(NS.w, 'txbxContent')];
    if (els.length !== all.length) return null;
    const up = (el, name) => { for (let n = el.parentNode; n && n.nodeType === 1; n = n.parentNode) if (n.localName === name) return n; return null; };
    const main = [], fallback = [];
    els.forEach((el, i) => (up(el, 'Fallback') ? fallback : main).push(i));
    // Word's fallback repeats each box's text (with its own bookmark ids), so copies are compared by their text.
    const text = (x) => x.replace(/<w:(tab|br|cr)\b[^>]*\/>/g, '\t').replace(/<w:t\b[^>]*>([^<]*)<\/w:t>|<[^>]+>/g, '$1');
    if (fallback.length && (fallback.length !== main.length || fallback.some((f, k) => text(all[f]) !== text(all[main[k]])))) return null;
    const kid = (el, name) => el && [...el.childNodes].find((n) => n.localName === name);
    const xy = (el, name) => { const e = kid(el, name); return e ? [Number(e.getAttribute(name === 'off' || name === 'chOff' ? 'x' : 'cx')), Number(e.getAttribute(name === 'off' || name === 'chOff' ? 'y' : 'cy'))] : null; };
    return main.map((i, k) => {
      const wsp = up(els[i], 'wsp');
      const bodyPr = kid(wsp, 'bodyPr');
      const ins = (a, d) => Math.round(emuToPx(bodyPr?.getAttribute(a) ?? d));
      // The box's frame in its group's coordinates, mapped out through any nested groups to the drawing.
      let rect = [0, 0, 1, 1];
      const xfrm = kid(kid(wsp, 'spPr'), 'xfrm');
      if (wsp && xfrm && up(wsp, 'wgp')) {
        let [x, y] = xy(xfrm, 'off') || [0, 0];
        let [w, h] = xy(xfrm, 'ext') || [0, 0];
        for (let g = wsp.parentNode; g && g.nodeType === 1; g = g.parentNode) {
          if (g.localName !== 'grpSp' && g.localName !== 'wgp') continue;
          const gx = kid(kid(g, 'grpSpPr'), 'xfrm');
          const [ox, oy] = (g.localName === 'wgp' ? [0, 0] : xy(gx, 'off')) || [0, 0];
          const [cx, cy] = xy(gx, 'chOff') || [0, 0];
          const [ex, ey] = xy(gx, 'ext') || [1, 1];
          const [sx, sy] = xy(gx, 'chExt') || [ex, ey];
          if (g.localName === 'wgp') { rect = [(x - cx) / sx, (y - cy) / sy, w / sx, h / sy]; break; }
          [x, y, w, h] = [ox + (x - cx) * ex / sx, oy + (y - cy) * ey / sy, w * ex / sx, h * ey / sy];
        }
      }
      return { rect, ins: [ins('tIns', 45720), ins('rIns', 91440), ins('bIns', 45720), ins('lIns', 91440)], copies: fallback.length ? [i, fallback[k]] : [i] };
    });
  }

  /** The text of text box `k` in a drawing as an editor doc, or null. */
  textBox(runXml, k = 0) {
    const box = this.textBoxes(runXml)?.[k];
    return box ? this._fragment([...runXml.matchAll(TXBX)][box.copies[0]][2]) : null;
  }

  /** The run XML with text box `k`'s text replaced by `doc` (in the shape and its VML fallback), and its new picture. */
  setTextBox(runXml, doc, k = 0) {
    const box = this.textBoxes(runXml)?.[k];
    if (!box) throw new Error('no such text box');
    let body = writeBlocks(doc, { linkRel: (href) => this._linkRel(href) });
    if (doc.lastChild && doc.lastChild.type.name === 'table') body += '<w:p/>';
    // The shape's copy is the one Word reads; the fallback copy gets the text without bookmarks, so ids stay unique.
    const plain = body.replace(/<w:bookmark(Start|End)\b[^>]*\/>/g, '');
    let i = 0;
    const xml = runXml.replace(TXBX, (m, open) => {
      const n = i++;
      return n === box.copies[0] ? `${open}${body}</w:txbxContent>` : n === box.copies[1] ? `${open}${plain}</w:txbxContent>` : m;
    });
    let src = null;
    this._fragment(`<w:p>${xml}</w:p>`).descendants((n) => { if (src == null && n.type.name === 'image') src = n.attrs.src; });
    return { xml, src };
  }

  /**
   * A shape or group (a text box) at a new size, its text wrapped again: the frame (wp:extent), the shape's own size
   * (or the group's, whose children follow it through chOff/chExt), and the fallback copy older Word versions read.
   * Returns { xml, src, width, height }.
   */
  resizeShape(runXml, width, height) {
    const ext = /<wp:extent\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(runXml);
    if (!ext) return null;
    const [cx, cy] = [pxToEmu(width), pxToEmu(height)];
    const kx = cx / Number(ext[1]), ky = cy / Number(ext[2]);
    let xml = runXml.replace(/(<wp:extent\b[^>]*\bcx=")\d+("[^>]*\bcy=")\d+"/, `$1${cx}$2${cy}"`);
    // The shape takes the size; a group's frame (wpg:grpSpPr) is scaled with it, its children following through
    // chOff/chExt.
    const group = /<wpg:grpSpPr\b/.test(xml);
    xml = xml.replace(/(<(?:wps:spPr|wpg:grpSpPr)\b[^>]*>\s*<a:xfrm\b[^>]*>[\s\S]*?<a:ext\b[^>]*\bcx=")(\d+)("[^>]*\bcy=")(\d+)"/, (m, a, x, b, y) => (group
      ? `${a}${Math.round(Number(x) * kx)}${b}${Math.round(Number(y) * ky)}"` : `${a}${cx}${b}${cy}"`));
    // The VML copy: its style's width and height (points).
    const pt = (px) => +(px * 0.75).toFixed(2);
    xml = xml.replace(/(<v:(?:shape|group|rect|roundrect)\b[^>]*\bstyle="[^"]*?)width:([\d.]+)pt;height:([\d.]+)pt/, (m, a, w, h) => (group
      ? `${a}width:${+(Number(w) * kx).toFixed(2)}pt;height:${+(Number(h) * ky).toFixed(2)}pt` : `${a}width:${pt(width)}pt;height:${pt(height)}pt`));
    return this._shapeNode(xml);
  }

  /**
   * A floating shape moved by (dx, dy) px: its offsets from what it's placed against (wp:positionH/V posOffset), and its
   * VML copy's margins. Null when it isn't placed by offsets (an alignment, "centred" or the like).
   */
  moveShape(runXml, dx, dy) {
    let moved = 0;
    let xml = runXml.replace(/(<wp:position([HV])\b[^>]*>\s*<wp:posOffset>)(-?\d+)(<\/wp:posOffset>)/g, (m, a, axis, v, b) => { moved++; return `${a}${Number(v) + pxToEmu(axis === 'H' ? dx : dy)}${b}`; });
    if (moved < 2) return null;
    const pt = (px) => px * 0.75;
    xml = xml.replace(/(<v:(?:shape|group|rect|roundrect)\b[^>]*\bstyle="[^"]*)/, (style) => style
      .replace(/margin-left:(-?[\d.]+)pt/, (m, v) => `margin-left:${+(Number(v) + pt(dx)).toFixed(2)}pt`)
      .replace(/margin-top:(-?[\d.]+)pt/, (m, v) => `margin-top:${+(Number(v) + pt(dy)).toFixed(2)}pt`));
    return this._shapeNode(xml);
  }

  /** A shape's run XML read again: its drawing, size and placement as the reader makes them. */
  _shapeNode(xml) {
    let out = null;
    // (The XML as given, not as read back: the rest of it stays byte for byte.)
    this._fragment(`<w:p>${xml}</w:p>`, true).descendants((n) => { if (!out && n.type.name === 'image') out = { xml, src: n.attrs.src, width: n.attrs.width, height: n.attrs.height, place: n.attrs.place }; });
    return out;
  }

  /**
   * A new, empty text box in line with the text (Word's Insert > Text Box: white fill, thin black border,
   * 0.1in by 0.05in inner margins), with the VML copy older Word versions read. Returns { xml, src }.
   */
  addTextBox(width, height) {
    const pt = (px) => +(px * 0.75).toFixed(2);
    this._picId = (this._picId || 1000) + 1;
    const id = this._picId;
    const ns = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
      + ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
      + ' xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"';
    const text = '<w:txbxContent><w:p/></w:txbxContent>';
    const xml = `<w:r><mc:AlternateContent ${ns}><mc:Choice Requires="wps"><w:drawing>`
      + `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${pxToEmu(width)}" cy="${pxToEmu(height)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
      + `<wp:docPr id="${id}" name="Text Box ${id}"/><wp:cNvGraphicFramePr/>`
      + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>'
      + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${pxToEmu(width)}" cy="${pxToEmu(height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
      + '<a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln w="6350"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></wps:spPr>'
      + `<wps:txbx>${text}</wps:txbx>`
      + '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
      + '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>'
      + '<mc:Fallback><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
      + `<v:shape id="Text Box ${id}" o:spid="_x0000_s${id}" type="#_x0000_t202" style="width:${pt(width)}pt;height:${pt(height)}pt" fillcolor="white [3201]" strokeweight=".5pt">`
      + `<v:textbox>${text}</v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
    let src = null;
    this._fragment(`<w:p>${xml}</w:p>`).descendants((n) => { if (src == null && n.type.name === 'image') src = n.attrs.src; });
    return { xml, src };
  }

  /** Serialize `doc` (default: the loaded one) to .docx bytes. */
  save(doc = this.doc) {
    const xml = writeDocument(doc, this.originalXml, { linkRel: (href) => this._linkRel(href) });
    // Edited charts: their parts, and the workbooks Word edits them in (chartedit.js). Each save starts from the parts
    // as they were opened, so an edit undone since an earlier save is undone in the file too.
    this._chartOrig ||= new Map(); // chart part path -> { xml, wb: { target, bytes } | null }, as opened
    for (const [path, o] of this._chartOrig) {
      this.pkg.setText(path, o.xml);
      if (o.wb) this.pkg.files[o.wb.target] = o.wb.bytes;
    }
    doc.descendants((n) => {
      if (n.type.name !== 'image' || !n.attrs.chart) return;
      const part = this.chartPart(n.attrs.xml);
      if (!part) return;
      const data = readChart(n.attrs.chart);
      const wb = data && workbookRel(n.attrs.chart, this.pkg.readRels(part.path));
      if (!this._chartOrig.has(part.path)) this._chartOrig.set(part.path, { xml: part.xml, wb: wb && this.pkg.files[wb.target] ? { target: wb.target, bytes: this.pkg.files[wb.target] } : null });
      this.pkg.setText(part.path, n.attrs.chart);
      const bytes = wb && this.pkg.files[wb.target] && updateWorkbook(this.pkg.files[wb.target], data, data.sheet);
      if (bytes) this.pkg.files[wb.target] = bytes;
    });
    this.pkg.setText(this.pkg.mainPath, xml);
    if (this.trackRevisions !== this._trackRevisions0 && this.settingsPath) {
      this.pkg.setText(this.settingsPath, withTrackRevisions(this.pkg.text(this.settingsPath), this.trackRevisions));
      this._trackRevisions0 = this.trackRevisions;
    }
    for (const rId of this._dirtyStories || []) {
      const path = this.pkg.relById(rId).target;
      this.pkg.setText(path, writeStory(this._stories.get(rId), this.pkg.text(path), {
        linkRel: (href) => this.pkg.addRel(REL_TYPES.hyperlink, href, true, path),
      }));
    }
    this._dirtyStories?.clear();
    const used = { footnote: new Set(), endnote: new Set() };
    doc.descendants((n) => {
      for (const m of (n.attrs.xml || '').matchAll(/<w:(footnote|endnote)Reference\b[^>]*w:id="([^"]*)"/g)) used[m[1]].add(m[2]);
    });
    saveNotes(this.pkg, this.notes, used);
    if (this.commentsDirty) {
      writeComments(this.pkg, this.comments);
      this.commentsDirty = false;
    }
    return this.pkg.toBytes();
  }
}

// w:settings children that come after w:trackRevisions (ECMA-376 Part 1, §17.15.1.78).
const AFTER_TRACK = ['doNotTrackMoves', 'doNotTrackFormatting', 'documentProtection', 'autoFormatOverride', 'styleLockTheme', 'styleLockQFSet', 'defaultTabStop', 'autoHyphenation', 'consecutiveHyphenLimit', 'hyphenationZone', 'doNotHyphenateCaps', 'showEnvelope', 'summaryLength', 'clickAndTypeStyle', 'defaultTableStyle', 'evenAndOddHeaders', 'bookFoldRevPrinting', 'bookFoldPrinting', 'bookFoldPrintingSheets', 'drawingGridHorizontalSpacing', 'drawingGridVerticalSpacing', 'displayHorizontalDrawingGridEvery', 'displayVerticalDrawingGridEvery', 'doNotUseMarginsForDrawingGridOrigin', 'drawingGridHorizontalOrigin', 'drawingGridVerticalOrigin', 'doNotShadeFormData', 'noPunctuationKerning', 'characterSpacingControl', 'printTwoOnOne', 'strictFirstAndLastChars', 'noLineBreaksAfter', 'noLineBreaksBefore', 'savePreviewPicture', 'doNotValidateAgainstSchema', 'saveInvalidXml', 'ignoreMixedContent', 'alwaysShowPlaceholderText', 'doNotDemarcateInvalidXml', 'saveXmlDataOnly', 'useXSLTWhenSaving', 'saveThroughXslt', 'showXMLTags', 'alwaysMergeEmptyNamespace', 'updateFields', 'hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'docVars', 'rsids', 'mathPr', 'attachedSchema', 'themeFontLang', 'clrSchemeMapping', 'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures', 'forceUpgrade', 'captions', 'readModeInkLockDown', 'smartTagType', 'schemaLibrary', 'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol', 'listSeparator'];

/** settings.xml with w:trackRevisions added or removed. */
function withTrackRevisions(xml, on) {
  const out = xml.replace(/<w:trackRevisions\b[^>]*\/>|<w:trackRevisions\b[^>]*>[\s\S]*?<\/w:trackRevisions>/g, '');
  if (!on) return out;
  let at = -1;
  for (const name of AFTER_TRACK) {
    const m = new RegExp(`<w:${name}[\\s/>]`).exec(out);
    if (m && (at < 0 || m.index < at)) at = m.index;
  }
  if (at < 0) at = out.lastIndexOf('</w:settings>');
  return at < 0 ? out : out.slice(0, at) + '<w:trackRevisions/>' + out.slice(at);
}

/** A picture's bytes and type as the browser can show them: an EMF drawn as SVG (the package keeps the EMF). */
function shown(bytes, mime) {
  if (!isEmf(bytes)) return [bytes, mime];
  try {
    const svg = emfToSvg(bytes);
    if (svg) return [new TextEncoder().encode(svg), 'image/svg+xml'];
  } catch { /* left as it was */ }
  return [bytes, mime];
}

function base64(bytes) {
  if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64');
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
