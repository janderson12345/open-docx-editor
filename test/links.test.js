import { test } from 'node:test';
import assert from 'node:assert/strict';
import { schema, safeHref } from '../src/editor/schema.js';

const hrefOf = (attrs) => schema.marks.link.spec.toDOM(schema.marks.link.create(attrs))[1].href;

test('a link is followed only to web, mail and phone addresses', () => {
  assert.equal(hrefOf({ href: 'https://example.com/a' }), 'https://example.com/a');
  assert.equal(hrefOf({ href: 'mailto:someone@example.com' }), 'mailto:someone@example.com');
  for (const bad of ['javascript:alert(1)', ' JavaScript:alert(1)', 'data:text/html,<script>alert(1)</script>', 'vbscript:x', 'file:///c:/x']) {
    assert.equal(hrefOf({ href: bad }), '#', bad);
  }
  assert.equal(hrefOf({ href: 'javascript:alert(1)', anchor: 'intro' }), '#intro');
  assert.equal(safeHref(null), null);
});

test('the document keeps its own address, whatever it is (only the page is guarded)', () => {
  const m = schema.marks.link.create({ href: 'javascript:alert(1)' });
  assert.equal(m.attrs.href, 'javascript:alert(1)');
});

test('a picture that is a link (a:hlinkClick) keeps its address, shown on the picture for Ctrl+click', async () => {
  await import('./setup.js');
  const { readDocument } = await import('../src/ooxml/reader.js');
  const { parseXml } = await import('../src/ooxml/xml.js');
  const body = '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="1" name="Logo"><a:hlinkClick r:id="rIdL"/></wp:docPr>'
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdP"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const json = readDocument(parseXml(`<w:document ${ns}><w:body>${body}</w:body></w:document>`), { resolveImage: () => 'blob:p', resolveLink: (id) => (id === 'rIdL' ? 'https://example.com/' : null) });
  const img = schema.node('image', json.content[0].content[0].attrs);
  assert.equal(img.attrs.href, 'https://example.com/');
  assert.equal(schema.nodes.image.spec.toDOM(img)[1]['data-href'], 'https://example.com/');
  const bad = schema.node('image', { ...img.attrs, href: 'javascript:alert(1)' });
  assert.equal(schema.nodes.image.spec.toDOM(bad)[1]['data-href'], null);
});

test('a link made in the editor has the Hyperlink character style, and is saved with it', async () => {
  const { setLink } = await import('../src/editor/commands.js');
  const { EditorState, TextSelection } = await import('prosemirror-state');
  const { DocxSession } = await import('../src/ooxml/docx.js');
  const session = DocxSession.blank();
  session.addBuiltInStyles(['Hyperlink']);
  const id = session.builtInId('Hyperlink');
  assert.ok(session.styles.styles.has(id), 'the document has the style');
  // Over a selection, and typed at the caret.
  const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, [schema.text('our site here')])]);
  let s = EditorState.create({ doc, selection: TextSelection.create(doc, 5, 9) });
  setLink('https://example.com/', id)(s, (tr) => { s = s.apply(tr); });
  const linked = [];
  s.doc.firstChild.forEach((n) => linked.push([n.text, n.marks.map((m) => m.type.name).sort().join('+')]));
  assert.deepEqual(linked, [['our ', ''], ['site', 'cstyle+link'], [' here', '']]);
  let c = EditorState.create({ doc, selection: TextSelection.create(doc, 1) });
  setLink('https://example.org/', id)(c, (tr) => { c = c.apply(tr); });
  assert.deepEqual(c.doc.firstChild.firstChild.marks.map((m) => m.type.name).sort(), ['cstyle', 'link']);
  // Saved: the link's run names the style, and styles.xml defines it as Word's own does.
  const saved = new DocxSession(session.save(s.doc));
  assert.match(saved.originalXml, new RegExp('<w:hyperlink [^>]*>(?:<w:r>|<w:r [^>]*>)<w:rPr><w:rStyle w:val="' + id + '"/>'));
  assert.match(saved.pkg.text('word/styles.xml'), /<w:style w:type="character" w:styleId="Hyperlink">[^]*?<w:color w:val="0563C1"[^>]*\/><w:u w:val="single"\/>/);
  // Taking the link off takes the style with it.
  setLink(null, id)(s, (tr) => { s = s.apply(tr); });
  assert.equal(s.doc.firstChild.childCount, 1);
  assert.deepEqual(s.doc.firstChild.firstChild.marks.map((m) => m.type.name), []);
});

test('removing a link takes its Hyperlink character style too', async () => {
  const { setLink } = await import('../src/editor/commands.js');
  const { EditorState, TextSelection } = await import('prosemirror-state');
  const t = schema.text('site', [schema.marks.cstyle.create({ val: 'Hyperlink' }), schema.marks.link.create({ href: 'https://example.com/' })]);
  const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, [t])]);
  let s = EditorState.create({ doc, selection: TextSelection.create(doc, 1, 5) });
  setLink(null)(s, (tr) => { s = s.apply(tr); });
  assert.deepEqual(s.doc.firstChild.firstChild.marks.map((m) => m.type.name), []);
});
