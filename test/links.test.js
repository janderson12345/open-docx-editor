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
