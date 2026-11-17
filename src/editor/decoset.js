// DecorationSet.create without its quadratic time. create gives each child node the decorations inside it by looking
// through all that are left, and checks each node decoration by finding its node from the parent's first child, so a
// document of many blocks with a decoration or more in each took time growing with the square of their number (4
// times the paragraphs, 8 to 20 times as long: comments, empty paragraphs' marks, merge fields, joined paragraphs, soft
// hyphens). Here the decorations are sorted once and each node takes the run of them starting inside it, a node
// decoration checked against the child it starts at: the same set as create's (test/stress.test.js compares them).
import { DecorationSet } from 'prosemirror-view';

const byPos = (a, b) => a.from - b.from || a.to - b.to;
const isNode = (d) => !d.inline && !d.widget;

/**
 * The set of `decos` (positions in `node`, its content starting at `offset`), sorted by start; as create's buildTree
 * makes it: each child's decorations, those strictly inside it, in its own set; the rest this node's own, those that
 * don't fit it dropped.
 */
function build(decos, node, offset) {
  const children = [];
  const locals = [];
  const fits = new Map(); // node decoration -> whether it covers the child it starts at
  let i = 0;
  node.forEach((child, rel) => {
    const start = offset + rel;
    const end = start + child.nodeSize;
    // (Starting before the child: not inside it, nor inside any after it.)
    while (i < decos.length && decos[i].from < start) locals.push(decos[i++]);
    for (; i < decos.length && decos[i].from === start; i++) {
      const d = decos[i];
      if (isNode(d)) fits.set(d, d.to === end && !child.isText);
      locals.push(d);
    }
    if (child.isLeaf) return;
    let found = null;
    for (; i < decos.length && decos[i].from < end; i++) {
      if (decos[i].to < end) (found ||= []).push(decos[i]);
      else locals.push(decos[i]);
    }
    if (!found) return;
    const sub = build(found, child, start + 1);
    if (sub !== DecorationSet.empty) children.push(rel, rel + child.nodeSize, sub);
  });
  while (i < decos.length) locals.push(decos[i++]);
  const own = [];
  for (const d of locals) {
    // (A node decoration starting inside a child, not at one, fits nothing; one at the very end is left to valid, which
    // throws as create does.)
    const ok = fits.has(d) ? fits.get(d) : isNode(d) && d.from - offset > 0 && d.from - offset < node.content.size ? false : null;
    const local = offset ? d.copy(d.from - offset, d.to - offset) : d;
    if (ok ?? local.type.valid(node, local)) own.push(local);
  }
  own.sort(byPos);
  return own.length || children.length ? new DecorationSet(own, children) : DecorationSet.empty;
}

/** DecorationSet.create(doc, decos), in time growing with the number of decorations and nodes, not their product. */
export function decoSet(doc, decos) {
  if (!decos.length) return DecorationSet.empty;
  // (A stable sort: decorations at the same place keep their order, as create's do.)
  return build(decos.slice().sort((a, b) => a.from - b.from), doc, 0);
}
