// Paragraphs joined by a tracked deletion of their paragraph mark. In Word's final view (simple markup, or none) a
// paragraph whose mark is deleted runs on into the next paragraph, in that paragraph's formatting (the formatting lives
// in the mark that's left). Here the joined paragraph's own box is hidden then, and its content is drawn again at the
// start of the paragraph it joins; with all markup shown, both stay as they are (see editor.css, odx-joined).
// The Original view (the document before its changes) joins the other way: a paragraph whose mark was inserted runs on
// into the next one (odx-ojoined), and those whose marks are deleted stay apart.
import { wrapText } from './textspans.js';
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { DOMSerializer } from 'prosemirror-model';
import { schema } from './schema.js';
import { paragraphMarkRev } from './track.js';
import { insertedWhole } from './pararules.js';

const key = new PluginKey('odx-joins');

/** True when nothing of the paragraph survives the deletion (it disappears instead: odx-all-del). */
function allDeleted(node) {
  let all = true;
  node.forEach((c) => { if (!c.marks.some((m) => m.type.name === 'del') && (c.isText ? c.text.trim() : c.type.name !== 'opaque_inline' || c.attrs.width != null)) all = false; });
  return all;
}

function compute(doc) {
  const decos = [];
  const serializer = DOMSerializer.fromSchema(schema);
  const visit = (parent, offset) => {
    let pending = []; // paragraphs (with text) whose marks are deleted, waiting for the paragraph they run into
    let added = []; // paragraphs whose marks were inserted (the Original view joins them to the next: odx-ojoined)
    parent.forEach((node, rel) => {
      const pos = offset + rel;
      if (node.type !== schema.nodes.paragraph) {
        pending = [];
        added = [];
        if (!node.isLeaf) visit(node, pos + 1);
        return;
      }
      const rev = paragraphMarkRev(node.attrs);
      if (rev === 'ins' && !insertedWhole(node)) added.push({ node, pos });
      else if (rev !== 'ins' && added.length) {
        const split = added;
        added = [];
        for (const j of split) decos.push(Decoration.node(j.pos, j.pos + j.node.nodeSize, { class: 'odx-ojoined-away' }));
        const sig = split.map((j) => j.node.textContent).join('\u0001');
        decos.push(Decoration.widget(pos + 1, () => {
          const span = document.createElement('span');
          span.className = 'odx-ojoined';
          span.contentEditable = 'false';
          for (const j of split) span.append(serializer.serializeFragment(j.node.content));
          wrapText(span);
          return span;
        }, { side: -2, key: `ojoin:${sig.length}:${sig.slice(0, 40)}`, ignoreSelection: true }));
      }
      const del = rev === 'del';
      if (del && allDeleted(node)) return; // hidden whole; the paragraphs around it still join
      if (del) { pending.push({ node, pos }); return; }
      if (!pending.length) return;
      const joined = pending;
      pending = [];
      for (const j of joined) decos.push(Decoration.node(j.pos, j.pos + j.node.nodeSize, { class: 'odx-joined-away' }));
      const sig = joined.map((j) => j.node.textContent).join('\u0001');
      decos.push(Decoration.widget(pos + 1, () => {
        const span = document.createElement('span');
        span.className = 'odx-joined';
        span.contentEditable = 'false';
        for (const j of joined) span.append(serializer.serializeFragment(j.node.content));
        wrapText(span);
        return span;
      }, { side: -1, key: `join:${sig.length}:${sig.slice(0, 40)}`, ignoreSelection: true }));
    });
  };
  visit(doc, 0);
  separators(doc, serializer, decos);
  return DecorationSet.create(doc, decos);
}

/** A hidden paragraph mark (w:vanish with w:specVanish: Word's style separator, for run-in headings). */
const hiddenMark = (node) => {
  const mark = (node.attrs.extra || []).find((e) => e[0] === 'rPr');
  return !!mark && /<w:vanish(?![^>]*w:val="(?:0|false|off)")/.test(mark[1]);
};

/**
 * Style separators: a paragraph whose mark is hidden runs on into the next one, always (it's how Word puts a run-in
 * heading and its text on one line). The pair keeps the first paragraph's formatting, as Word lays it out, so the
 * next paragraph's content is drawn at the end of the first and its own box is hidden.
 */
function separators(doc, serializer, decos) {
  const visit = (parent, offset) => {
    let host = null; // { pos, node, take: [] }
    const flush = () => {
      if (!host || !host.take.length) return;
      for (const t of host.take) decos.push(Decoration.node(t.pos, t.pos + t.node.nodeSize, { class: 'odx-sep-away' }));
      const sig = host.take.map((t) => t.node.textContent).join('\u0001');
      const take = host.take;
      decos.push(Decoration.widget(host.pos + host.node.nodeSize - 1, () => {
        const span = document.createElement('span');
        span.className = 'odx-sep';
        span.contentEditable = 'false';
        for (const t of take) span.append(serializer.serializeFragment(t.node.content));
        wrapText(span);
        return span;
      }, { side: 1, key: `sep:${sig.length}:${sig.slice(0, 40)}`, ignoreSelection: true }));
    };
    parent.forEach((node, rel) => {
      const pos = offset + rel;
      if (node.type !== schema.nodes.paragraph) {
        flush(); host = null;
        if (!node.isLeaf) visit(node, pos + 1);
        return;
      }
      if (host) host.take.push({ pos, node });
      else if (hiddenMark(node)) host = { pos, node, take: [] };
      if (!hiddenMark(node)) { flush(); host = null; }
    });
    flush();
  };
  visit(doc, 0);
}

export const joinsPlugin = new Plugin({
  key,
  state: {
    init: (_, state) => compute(state.doc),
    apply: (tr, set) => (tr.docChanged ? compute(tr.doc) : set),
  },
  props: { decorations: (state) => key.getState(state) },
});
