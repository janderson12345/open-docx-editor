// Paragraphs joined by a tracked deletion of their paragraph mark. In Word's final view (simple markup, or none) a
// paragraph whose mark is deleted runs on into the next paragraph, in that paragraph's formatting (the formatting lives
// in the mark that's left). Here the joined paragraph's own box is hidden then, and its content is drawn again at the
// start of the paragraph it joins; with all markup shown, both stay as they are (see editor.css, odx-joined).
// The Original view (the document before its changes) joins the other way: a paragraph whose mark was inserted runs on
// into the next one (odx-ojoined), and those whose marks are deleted stay apart.
import { wrapText } from './textspans.js';
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration } from 'prosemirror-view';
import { decoSet } from './decoset.js';
import { DOMSerializer } from 'prosemirror-model';
import { schema } from './schema.js';
import { paragraphMarkRev } from './track.js';
import { insertedWhole } from './pararules.js';
import { paragraphList } from './commands.js';

const key = new PluginKey('odx-joins');

/** True when nothing of the paragraph survives the deletion (it disappears instead: odx-all-del). */
function allDeleted(node) {
  let all = true;
  node.forEach((c) => { if (!c.marks.some((m) => m.type.name === 'del') && (c.isText ? c.text.trim() : c.type.name !== 'opaque_inline' || c.attrs.width != null)) all = false; });
  return all;
}

// The paragraph a paragraph is run into through hidden marks takes its page keeping (w:keepLines) from the first of
// them with text, as Word paginates it (Word's PDFs of built documents: keepLines on the paragraph run into, the pair
// split at the page foot; on the hidden one, moved whole). pages.js asks; set afresh at each compute.
let leads = new WeakMap();
/** The paragraph whose page keeping `node` takes, when paragraphs run on into it through hidden marks. */
export const runOnLead = (node) => leads.get(node) || null;

function compute(doc, session) {
  leads = new WeakMap();
  const modern = session?.compatMode >= 15;
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
      // A hidden paragraph mark that isn't a style separator runs on into the next paragraph in every view, in that
      // paragraph's formatting, as a deleted mark does in the final view (Word's PDFs of built documents: a first
      // paragraph indented 2in with its mark hidden, inserted or not, the pair at the second's indent; and the other
      // way round). In compatibility mode 15 only: in 14 and before (or with none set) the pair keeps the first's
      // formatting, as a style separator does (the same document in modes 11, 12 and 14: at 2in).
      if (modern && hiddenMark(node) && !specMark(node)) { pending.push({ node, pos, hidden: true }); return; }
      const del = rev === 'del';
      if (del && allDeleted(node)) return; // hidden whole; the paragraphs around it still join
      if (del) { pending.push({ node, pos }); return; }
      if (!pending.length) return;
      const joined = pending;
      pending = [];
      for (const j of joined) decos.push(Decoration.node(j.pos, j.pos + j.node.nodeSize, { class: j.hidden ? 'odx-hjoined-away' : 'odx-joined-away' }));
      const widget = (list, cls) => {
        const sig = list.map((j) => j.node.textContent).join('\u0001');
        decos.push(Decoration.widget(pos + 1, () => {
          const span = document.createElement('span');
          span.className = cls;
          span.contentEditable = 'false';
          for (const j of list) span.append(serializer.serializeFragment(j.node.content));
          wrapText(span);
          return span;
        }, { side: -1, key: `${cls}:${sig.length}:${sig.slice(0, 40)}`, ignoreSelection: true }));
      };
      // All of them in the final view (odx-joined); with all markup shown, or in the Original view, where paragraphs
      // whose marks are deleted stay apart, only those whose marks are hidden (odx-hjoined).
      widget(joined, 'odx-joined');
      const hidden = joined.filter((j) => j.hidden);
      if (hidden.length) widget(hidden, 'odx-hjoined');
      // The paragraph run into keeps its number only when the first of them with text is numbered too; otherwise the
      // text starts where the number would be (Word's PDFs of built documents: a plain paragraph with its mark hidden
      // before a numbered one, no number, the text at the number's place; both numbered, the number; the hidden one
      // empty, the number; a sample's body paragraph run into a numbered heading, no number).
      const lead = hidden.find((j) => j.node.textContent.trim());
      if (lead) leads.set(node, lead.node);
      if (lead && !paragraphList(session?.styles ? session : null, lead.node.attrs)) decos.push(Decoration.node(pos, pos + node.nodeSize, { class: 'odx-hjoin-nolabel' }));
    });
  };
  visit(doc, 0);
  separators(doc, serializer, decos, modern);
  return decoSet(doc, decos);
}

/** A hidden paragraph mark (w:vanish). */
const hiddenMark = (node) => {
  const mark = (node.attrs.extra || []).find((e) => e[0] === 'rPr');
  return !!mark && /<w:vanish(?=[\s/>])(?![^>]*w:val="(?:0|false|off)")/.test(mark[1]);
};
/**
 * True when the paragraph runs on into the next one in that one's formatting: a hidden mark, not a style separator, in
 * compatibility mode 15 (editor.js numbers the pair: the first's number, the next one's not counted).
 */
export const runsOnHidden = (node, session) => session?.compatMode >= 15 && hiddenMark(node) && !specMark(node);
/** A hidden paragraph mark that is also w:specVanish: Word's style separator, for run-in headings. */
const specMark = (node) => {
  const mark = (node.attrs.extra || []).find((e) => e[0] === 'rPr');
  return !!mark && /<w:specVanish(?=[\s/>])(?![^>]*w:val="(?:0|false|off)")/.test(mark[1]);
};

/**
 * Style separators: a paragraph whose mark is hidden and w:specVanish (or any hidden mark before compatibility mode 15)
 * runs on into the next one, always (it's how Word
 * puts a run-in heading and its text on one line). The pair keeps the first paragraph's formatting, as Word lays it out
 * (Word's PDFs of built documents: the first paragraph's indent, either way round), so the
 * next paragraph's content is drawn at the end of the first and its own box is hidden.
 */
function separators(doc, serializer, decos, modern) {
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
      else if (hiddenMark(node) && (specMark(node) || !modern)) host = { pos, node, take: [] };
      if (!hiddenMark(node)) { flush(); host = null; }
    });
    flush();
  };
  visit(doc, 0);
}

/** The joins of a document in `session`'s compatibility mode (a hidden mark's paragraph runs on differently from 15). */
export const joinsPluginFor = (session) => new Plugin({
  key,
  state: {
    init: (_, state) => compute(state.doc, session),
    apply: (tr, set) => (tr.docChanged ? compute(tr.doc, session) : set),
  },
  props: { decorations: (state) => key.getState(state) },
});
export const joinsPlugin = joinsPluginFor({ compatMode: 15 });
