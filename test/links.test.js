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

test('removing a link takes its Hyperlink character style too', async () => {
  const { setLink } = await import('../src/editor/commands.js');
  const { EditorState, TextSelection } = await import('prosemirror-state');
  const t = schema.text('site', [schema.marks.cstyle.create({ val: 'Hyperlink' }), schema.marks.link.create({ href: 'https://example.com/' })]);
  const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, [t])]);
  let s = EditorState.create({ doc, selection: TextSelection.create(doc, 1, 5) });
  setLink(null)(s, (tr) => { s = s.apply(tr); });
  assert.deepEqual(s.doc.firstChild.firstChild.marks.map((m) => m.type.name), []);
});
