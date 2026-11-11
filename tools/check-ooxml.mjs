// Structural check of saved .docx files, the things that make Word call a file corrupt: every XML part well-formed,
// every part with a content type, every internal relationship's target present, every r:id the main document uses
// defined in its relationships, every namespace prefix declared. (Not a full XSD validation: that needs the ECMA-376
// schemas.)
// usage: node tools/check-ooxml.mjs <file.docx | dir>...
import fs from 'node:fs';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';

const files = process.argv.slice(2).flatMap((a) => (fs.statSync(a).isDirectory() ? fs.readdirSync(a).filter((f) => f.endsWith('.docx')).map((f) => path.join(a, f)) : [a]));
if (!files.length) { console.error('usage: node tools/check-ooxml.mjs <file.docx | dir>...'); process.exit(2); }

const parse = (text, where, problems) => {
  try {
    return new DOMParser({ onError: (level, msg) => { if (level !== 'warning') throw new Error(msg); } }).parseFromString(text, 'text/xml');
  } catch (e) {
    problems.push(`${where}: not well-formed: ${e.message.split('\n')[0]}`);
    return null;
  }
};
const resolve = (base, target) => path.posix.normalize(path.posix.join(base ? path.posix.dirname(base) : '', target)).replace(/^\//, '');

let bad = 0;
for (const file of files) {
  const problems = [];
  const zip = unzipSync(new Uint8Array(fs.readFileSync(file)));
  const names = Object.keys(zip);
  const xml = new Map();
  for (const n of names) if (/\.(xml|rels)$/i.test(n)) { const d = parse(strFromU8(zip[n]), n, problems); if (d) xml.set(n, d); }
  const types = xml.get('[Content_Types].xml');
  if (!types) problems.push('[Content_Types].xml missing');
  else {
    const defaults = new Set([...types.getElementsByTagName('Default')].map((e) => e.getAttribute('Extension').toLowerCase()));
    const overrides = new Set([...types.getElementsByTagName('Override')].map((e) => e.getAttribute('PartName').replace(/^\//, '')));
    for (const n of names) {
      if (n.endsWith('/') || n === '[Content_Types].xml') continue;
      if (!overrides.has(n) && !defaults.has(n.split('.').pop().toLowerCase())) problems.push(`${n}: no content type`);
    }
    for (const o of overrides) if (!names.includes(o)) problems.push(`[Content_Types].xml: override for missing part ${o}`);
  }
  // Relationships: internal targets must exist.
  const relsOf = new Map(); // part -> Map(id -> target)
  for (const [n, d] of xml) {
    if (!n.endsWith('.rels')) continue;
    const at = n.lastIndexOf('_rels/');
    const part = n.slice(0, at) + n.slice(at + 6).replace(/\.rels$/, ''); // ('' for the package's own)
    const ids = new Map();
    for (const r of d.getElementsByTagName('Relationship')) {
      const id = r.getAttribute('Id');
      if (ids.has(id)) problems.push(`${n}: relationship ${id} defined twice`);
      ids.set(id, r.getAttribute('Target'));
      if (r.getAttribute('TargetMode') === 'External') continue;
      const t = r.getAttribute('Target');
      const full = t.startsWith('/') ? t.slice(1) : resolve(part, t);
      if (!names.includes(full)) problems.push(`${n}: ${id} points to missing ${full}`);
    }
    relsOf.set(part, ids);
  }
  // The main document and its headers, footers, notes and comments: every r:id they use is defined.
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  for (const [n, d] of xml) {
    if (!/^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/.test(n)) continue;
    const ids = relsOf.get(n) || new Map();
    const all = d.getElementsByTagName('*');
    for (let i = 0; i < all.length; i++) {
      const attrs = all[i].attributes;
      for (let j = 0; j < attrs.length; j++) {
        const a = attrs[j];
        if (a.namespaceURI === R && !ids.has(a.value)) problems.push(`${n}: ${all[i].tagName} uses ${a.name}="${a.value}", not in its relationships`);
      }
    }
  }
  // Namespaces: every prefixed element and attribute has its prefix declared (Word refuses a part that uses one it
  // can't resolve), and every prefix mc:Ignorable names is declared too.
  const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  for (const [n, d] of xml) {
    const all = d.getElementsByTagName('*');
    const undeclared = new Set();
    for (let i = 0; i < all.length; i++) {
      const el = all[i];
      if (el.prefix && !el.namespaceURI) undeclared.add(el.prefix);
      for (let j = 0; j < el.attributes.length; j++) {
        const a = el.attributes[j];
        if (a.prefix && a.prefix !== 'xmlns' && a.prefix !== 'xml' && !a.namespaceURI) undeclared.add(a.prefix);
      }
      const ign = el.getAttributeNS(MC, 'Ignorable');
      for (const p of (ign || '').split(/\s+/).filter(Boolean)) if (!el.lookupNamespaceURI(p)) undeclared.add(`${p} (mc:Ignorable)`);
    }
    for (const p of undeclared) problems.push(`${n}: prefix ${p} used but not declared`);
  }
  if (!names.includes('word/document.xml')) problems.push('word/document.xml missing');
  if (problems.length) { bad++; console.log(`FAIL ${file}`); for (const p of problems.slice(0, 20)) console.log(`  ${p}`); }
  else console.log(`ok   ${file}`);
}
console.log(bad ? `${bad} of ${files.length} files with problems` : `all ${files.length} files ok`);
process.exit(bad ? 1 : 0);
