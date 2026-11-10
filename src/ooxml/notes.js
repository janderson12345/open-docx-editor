// Footnotes and endnotes (word/footnotes.xml, word/endnotes.xml). Notes we
// don't edit are written back as they were; separators are always kept.
import { Node } from 'prosemirror-model';
import { NS } from './xml.js';
import { readNotes } from './reader.js';
import { writeBlocks, rootPrefixes } from './writer.js';
import { schema } from '../editor/schema.js';

export const NOTE_KINDS = {
  footnote: { rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes', file: 'footnotes.xml', root: 'footnotes', ct: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml', textStyle: 'FootnoteText', refStyle: 'FootnoteReference' },
  endnote: { rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes', file: 'endnotes.xml', root: 'endnotes', ct: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml', textStyle: 'EndnoteText', refStyle: 'EndnoteReference' },
};

/** { footnote: Map(id -> { id, type, xml, doc, dirty }), endnote: Map } */
export function loadNotes(pkg) {
  const out = {};
  for (const [kind, k] of Object.entries(NOTE_KINDS)) {
    out[kind] = new Map();
    const path = pkg.partByType(k.rel);
    const xml = path && pkg.xml(path);
    if (!xml) continue;
    for (const n of readNotes(xml)) {
      const doc = Node.fromJSON(schema, { type: 'doc', content: n.content.length ? n.content : [{ type: 'paragraph' }] });
      out[kind].set(n.id, { id: n.id, type: n.type, xml: n.xml, doc, dirty: false });
    }
  }
  return out;
}

const SEPARATORS = (tag) => `<w:${tag} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:${tag}>`
  + `<w:${tag} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:${tag}>`;

/** Write notes that changed, dropping ones no longer referenced from `used` (Set of ids per kind). */
export function saveNotes(pkg, notes, used) {
  for (const [kind, k] of Object.entries(NOTE_KINDS)) {
    const map = notes[kind];
    const changed = [...map.values()].some((n) => n.dirty) || [...map.values()].some((n) => n.type === 'normal' && !used[kind].has(n.id));
    if (!changed) continue;
    let path = pkg.partByType(k.rel);
    let src = path && pkg.text(path);
    const tag = kind;
    if (!src) {
      path = `${pkg.mainDir}${k.file}`;
      pkg.addRel(k.rel, k.file);
      pkg.addOverride(path, k.ct);
      src = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${k.root} xmlns:w="${NS.w}" xmlns:r="${NS.r}">${SEPARATORS(tag)}</w:${k.root}>`;
    }
    const open = new RegExp(`<w:${k.root}\\b[^>]*>`).exec(src);
    const keepSeps = [...src.matchAll(new RegExp(`<w:${tag}\\b[^>]*w:type="(?:separator|continuationSeparator|continuationNotice)"[\\s\\S]*?</w:${tag}>`, 'g'))].map((m) => m[0]).join('') || SEPARATORS(tag);
    const wctx = { prefixes: rootPrefixes(src) };
    const body = [...map.values()].filter((n) => n.type === 'normal' && used[kind].has(n.id)).map((n) => {
      if (!n.dirty && n.xml) return n.xml;
      return `<w:${tag} w:id="${n.id}">${writeBlocks(n.doc, wctx)}</w:${tag}>`;
    }).join('');
    pkg.setText(path, `${src.slice(0, open.index)}${open[0]}${keepSeps}${body}</w:${k.root}>`);
    for (const n of map.values()) n.dirty = false;
  }
}
