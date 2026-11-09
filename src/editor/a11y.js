// Review > Check Accessibility: the problems Word's checker reports most often.
import { schema } from './schema.js';
import { outlineLevel } from './commands.js';

/** [{ kind: 'error'|'warning'|'tip', title, detail, pos }] */
export function checkAccessibility(view) {
  const { session } = view.odx;
  const doc = view.state.doc;
  const out = [];
  const descr = (xml) => (/<wp:docPr\b[^>]*\sdescr="([^"]*)"/.exec(xml || '') || [])[1];
  let lastLevel = -1;
  doc.descendants((node, pos) => {
    if (node.type === schema.nodes.image || (node.type === schema.nodes.opaque_inline && node.attrs.kind === 'drawing')) {
      const alt = node.attrs.alt || descr(node.attrs.xml);
      const decorative = /a16:decorative\b[^>]*val="1"|adec:decorative\b[^>]*val="1"/.test(node.attrs.xml || '');
      if (!alt && !decorative) out.push({ kind: 'error', title: 'Missing alternative text', detail: 'Describe this picture for people using screen readers.', pos });
    }
    if (node.type === schema.nodes.table) {
      const first = node.firstChild;
      if (node.childCount > 1 && !/<w:tblHeader(?![^>]*w:val="(?:0|false|off)")/.test(first?.attrs.trPr || '')) out.push({ kind: 'warning', title: 'No header row', detail: 'Mark the first row as a header row so it repeats and is read as headings.', pos });
      return false;
    }
    if (node.type === schema.nodes.paragraph) {
      const lvl = outlineLevel(session, node);
      if (lvl != null) {
        if (!node.textContent.trim()) out.push({ kind: 'warning', title: 'Empty heading', detail: 'Remove it or add text.', pos });
        else if (lvl > lastLevel + 1) out.push({ kind: 'tip', title: 'Skipped heading level', detail: `Heading ${lvl + 1} follows ${lastLevel < 0 ? 'no heading' : `Heading ${lastLevel + 1}`}.`, pos });
        lastLevel = lvl;
      }
      node.forEach((c) => {
        const link = c.marks.find((m) => m.type === schema.marks.link);
        if (link && c.isText && /^(https?:\/\/|www\.)\S+$/.test(c.text.trim())) out.push({ kind: 'tip', title: 'Unclear link text', detail: `"${c.text.trim().slice(0, 40)}" reads as an address; use words that say where it goes.`, pos });
      });
      return false;
    }
    return true;
  });
  const core = session.pkg.text('docProps/core.xml') || '';
  if (!/<dc:title>[^<]+<\/dc:title>/.test(core)) out.unshift({ kind: 'tip', title: 'No document title', detail: 'Add a title in the file properties.', pos: null });
  return out;
}
