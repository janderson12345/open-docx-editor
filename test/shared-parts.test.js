// Parts another editor in a co-editing room shares (src/collab.js): a picture or a header added under the name they
// give is kept to a plain file name of its own in the package's folders, whatever name arrives.
import './setup.js';
import test from 'node:test';
import assert from 'node:assert/strict';
import { DocxSession } from '../src/ooxml/docx.js';

const png = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

test('a shared picture: a name with "../" or no extension stays a plain file in media/, and two alike both go in', () => {
  const s = DocxSession.blank();
  s.addSharedImage({ rId: 'rIdA', name: '../../evil.png', mime: 'image/png', bytes: png });
  s.addSharedImage({ rId: 'rIdB', name: 'noext', mime: 'image/png', bytes: png });
  s.addSharedImage({ rId: 'rIdC', name: 'noext', mime: 'image/png', bytes: png });
  s.addSharedImage({ rId: 'rIdD', name: 'x.png', mime: 'text/html', bytes: png });
  const media = Object.keys(s.pkg.files).filter((p) => /media\//.test(p));
  assert.ok(media.every((p) => /^word\/media\/[A-Za-z0-9_-]+\.[A-Za-z0-9]+$/.test(p)), media.join(', '));
  assert.equal(media.length, 3, media.join(', '));
  assert.ok(!s.pkg.relById('rIdD'));
});

test('a shared header: only a header or footer, under a name of its kind', () => {
  const s = DocxSession.blank();
  const xml = '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>';
  s.addSharedStory({ rId: 'rIdH', kind: 'header', name: '../../etc/x.xml', xml });
  s.addSharedStory({ rId: 'rIdS', kind: 'styles', name: 'styles.xml', xml });
  const parts = Object.keys(s.pkg.files).filter((p) => /header|footer|etc/.test(p));
  assert.ok(parts.every((p) => /^word\/(header|footer)[-a-z0-9]*\.xml$/.test(p)), parts.join(', '));
  assert.ok(!s.pkg.relById('rIdS'));
});

test('a picture node from a co-editor: no other site\'s address or CSS of its own reaches the page', async () => {
  const { schema } = await import('../src/editor/schema.js');
  const dom = (attrs) => schema.nodes.image.spec.toDOM(schema.nodes.image.create(attrs))[1];
  const bad = dom({ src: 'javascript:alert(1)', width: '1px;background:url(https://x.test/a)', height: 10, place: 'float:left;background:url(https://x.test/b)' });
  assert.equal(bad.src, '');
  assert.ok(!/url\(|x\.test/.test(bad.style), bad.style);
  const good = dom({ src: 'data:image/png;base64,AAAA', width: 20, height: 10, place: 'float:left;margin:0px 12px 0px 0px' });
  assert.equal(good.src, 'data:image/png;base64,AAAA');
  assert.match(good.style, /width:20px;height:10px;float:left/);
});
