// Comments: word/comments.xml (author, date, text) plus Word 2013+'s
// commentsExtended.xml (reply threads and "resolved"). Comments we don't touch
// are written back byte-for-byte.
import { parseXml, kids, attr, NS, esc } from './xml.js';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
export const COMMENT_RELS = {
  comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  extended: 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
  ids: 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds',
  extensible: 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible',
};
const CT = {
  comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
  extended: 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml',
};

/**
 * The document's comments, in file order:
 * [{ id, author, initials, date, text, paraId, parentId, done, xml }]
 * (`xml` is the original w:comment element; null once edited).
 */
export function readComments(pkg) {
  const path = pkg.partByType(COMMENT_RELS.comments);
  const src = path && pkg.text(path);
  if (!src) return [];
  const raw = src.match(/<w:comment\b[\s\S]*?<\/w:comment>|<w:comment\b[^>]*\/>/g) || [];
  const rootTag = (/<w:comments\b[^>]*>/.exec(src) || [''])[0];
  const list = raw.map((xml) => {
    const el = kids(parseXml(`${rootTag}${xml}</w:comments>`).documentElement, 'comment')[0];
    const ps = kids(el, 'p');
    const text = ps.map((p) => [...p.getElementsByTagNameNS(NS.w, '*')]
      .map((n) => (n.localName === 't' ? n.textContent : n.localName === 'tab' ? '\t' : n.localName === 'br' ? '\n' : '')).join('')).join('\n');
    const last = ps[ps.length - 1];
    return {
      id: attr(el, 'id'), author: attr(el, 'author') || '', initials: attr(el, 'initials') || '', date: attr(el, 'date') || '',
      text, paraId: last ? last.getAttributeNS(W14, 'paraId') || null : null, parentId: null, done: false, xml,
    };
  });
  const ext = pkg.partByType(COMMENT_RELS.extended);
  const extSrc = ext && pkg.text(ext);
  if (extSrc) {
    const byPara = new Map(list.filter((c) => c.paraId).map((c) => [c.paraId, c]));
    for (const m of extSrc.matchAll(/<w15:commentEx\b[^>]*>/g)) {
      const a = (n) => (new RegExp(`w15:${n}="([^"]*)"`).exec(m[0]) || [])[1];
      const c = byPara.get(a('paraId'));
      if (!c) continue;
      c.done = a('done') === '1';
      const parent = byPara.get(a('paraIdParent'));
      if (parent) c.parentId = parent.id;
    }
  }
  return list;
}

const newParaId = () => (Math.floor(Math.random() * 0x7ffffffe) + 1).toString(16).toUpperCase().padStart(8, '0');

function commentXml(c) {
  c.paraId ||= newParaId();
  const lines = String(c.text).split('\n');
  const ps = lines.map((line, i) => {
    const runs = line.split('\t').map((t, j) => (j ? '<w:r><w:tab/></w:r>' : '') + (t ? `<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r>` : '')).join('');
    const ref = i === 0 ? '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>' : '';
    const pid = i === lines.length - 1 ? ` w14:paraId="${c.paraId}" w14:textId="77777777"` : '';
    return `<w:p${pid}><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>${ref}${runs}</w:p>`;
  }).join('');
  return `<w:comment w:id="${c.id}" w:author="${esc(c.author)}" w:date="${esc(c.date)}" w:initials="${esc(c.initials)}">${ps}</w:comment>`;
}

/** Write `list` back into the package (adding the parts on first use). */
export function writeComments(pkg, list) {
  let path = pkg.partByType(COMMENT_RELS.comments);
  const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  let src = path && pkg.text(path);
  if (!src) {
    if (!list.length) return;
    path = `${pkg.mainDir}comments.xml`;
    pkg.addRel(COMMENT_RELS.comments, 'comments.xml');
    pkg.addOverride(path, CT.comments);
    src = `${decl}<w:comments xmlns:w="${NS.w}" xmlns:w14="${W14}" xmlns:mc="${MC}" mc:Ignorable="w14"></w:comments>`;
  }
  for (const c of list) if (!c.paraId) c.xml = null; // needs a paraId to carry thread/resolved state
  // (An empty part may close its root at once, <w:comments .../>: open it, as the comments go inside.)
  let rootTag = (/<w:comments\b[^>]*>/.exec(src) || [''])[0].replace(/\s*\/>$/, '>');
  if (!/xmlns:w14=/.test(rootTag)) rootTag = rootTag.replace(/>$/, ` xmlns:w14="${W14}">`);
  const start = src.search(/<w:comments\b/);
  pkg.setText(path, `${src.slice(0, start)}${rootTag}${list.map((c) => c.xml || commentXml(c)).join('')}</w:comments>`);

  // Threads and resolved state.
  const byId = new Map(list.map((c) => [c.id, c]));
  let ext = pkg.partByType(COMMENT_RELS.extended);
  let extSrc = ext && pkg.text(ext);
  if (!extSrc) {
    ext = `${pkg.mainDir}commentsExtended.xml`;
    pkg.addRel(COMMENT_RELS.extended, 'commentsExtended.xml');
    pkg.addOverride(ext, CT.extended);
    extSrc = `${decl}<w15:commentsEx xmlns:w15="${W15}" xmlns:mc="${MC}" mc:Ignorable="w15"></w15:commentsEx>`;
  }
  const extRaw = (/<w15:commentsEx\b[^>]*>/.exec(extSrc) || [''])[0];
  const extRoot = extRaw.replace(/\s*\/>$/, '>');
  const entries = list.map((c) => {
    const parent = c.parentId != null && byId.get(c.parentId);
    return `<w15:commentEx w15:paraId="${c.paraId}"${parent ? ` w15:paraIdParent="${parent.paraId}"` : ''} w15:done="${c.done ? 1 : 0}"/>`;
  }).join('');
  pkg.setText(ext, `${extSrc.slice(0, extSrc.indexOf(extRaw))}${extRoot}${entries}</w15:commentsEx>`);

  // Durable ids of comments that no longer exist (Word rebuilds missing ones).
  const live = new Set(list.map((c) => c.paraId));
  const idsPath = pkg.partByType(COMMENT_RELS.ids);
  const idsSrc = idsPath && pkg.text(idsPath);
  if (idsSrc) {
    const gone = new Set();
    const out = idsSrc.replace(/<w16cid:commentId\b[^>]*\/>/g, (m) => {
      const p = (/w16cid:paraId="([^"]*)"/.exec(m) || [])[1];
      if (live.has(p)) return m;
      gone.add((/w16cid:durableId="([^"]*)"/.exec(m) || [])[1]);
      return '';
    });
    pkg.setText(idsPath, out);
    const xPath = pkg.partByType(COMMENT_RELS.extensible);
    const xSrc = xPath && pkg.text(xPath);
    if (xSrc && gone.size) {
      pkg.setText(xPath, xSrc.replace(/<w16cex:commentExtensible\b[^>]*?(?:\/>|>[\s\S]*?<\/w16cex:commentExtensible>)/g,
        (m) => (gone.has((/w16cex:durableId="([^"]*)"/.exec(m) || [])[1]) ? '' : m)));
    }
  }
}
