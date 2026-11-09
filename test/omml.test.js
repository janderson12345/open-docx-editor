import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { DocxSession } from '../src/ooxml/docx.js';
import { ommlToMathML } from '../src/ooxml/omml.js';

const fixture = new URL('./fixtures/math.docx', import.meta.url);
const R = (t) => `<m:r><m:t>${t}</m:t></m:r>`;

test('an equation kept without its prefixes declared (as the reader keeps it) is still read', () => {
  assert.equal(ommlToMathML(`<m:oMath>${R('x')}</m:oMath>`), '<math><mi>x</mi></math>');
  assert.equal(ommlToMathML('<w:p/>'), null);
});

test('a fraction has one numerator and one denominator, however many tokens each holds', () => {
  const m = ommlToMathML(`<m:oMath><m:f><m:num>${R('1')}</m:num><m:den>${R('2a')}</m:den></m:f></m:oMath>`);
  assert.equal(m, '<math><mfrac><mn>1</mn><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>');
});

test('a hyphen is a minus sign, an infix one between its operands in the same row', () => {
  const m = ommlToMathML(`<m:oMath><m:sSup><m:e>${R('b')}</m:e><m:sup>${R('2')}</m:sup></m:sSup>${R('-4ac')}</m:oMath>`);
  assert.match(m, /^<math><mrow><msup><mi>b<\/mi><mn>2<\/mn><\/msup><mo>−<\/mo><mn>4<\/mn>/);
});

test('n-ary operators: a sum\'s limits over and under, an integral\'s beside it; delimiters and roots', () => {
  const sum = ommlToMathML(`<m:oMath><m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>${R('i=1')}</m:sub><m:sup>${R('n')}</m:sup><m:e>${R('i')}</m:e></m:nary></m:oMath>`);
  assert.match(sum, /<munderover><mo[^>]*>∑<\/mo>/);
  const int = ommlToMathML(`<m:oMath><m:nary><m:sub>${R('0')}</m:sub><m:sup>${R('1')}</m:sup><m:e>${R('x')}</m:e></m:nary></m:oMath>`);
  assert.match(int, /<msubsup><mo[^>]*>∫<\/mo>/);
  const d = ommlToMathML(`<m:oMath><m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e>${R('x')}</m:e></m:d></m:oMath>`);
  assert.match(d, /<mo fence="true"[^>]*>\[<\/mo><mi>x<\/mi><mo fence="true"[^>]*>\]<\/mo>/);
  assert.match(ommlToMathML(`<m:oMath><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${R('2')}</m:e></m:rad></m:oMath>`), /<msqrt><mn>2<\/mn><\/msqrt>/);
  assert.match(ommlToMathML(`<m:oMath><m:rad><m:deg>${R('3')}</m:deg><m:e>${R('x')}</m:e></m:rad></m:oMath>`), /<mroot><mi>x<\/mi><mn>3<\/mn><\/mroot>/);
});

test('a display equation (m:oMathPara) is a block; every equation of the fixture is drawn', () => {
  const s = new DocxSession(fs.readFileSync(fixture));
  const out = [];
  s.doc.descendants((n) => { if (n.type.name === 'opaque_inline' && /^oMath/.test(n.attrs.kind)) out.push([n.attrs.kind, ommlToMathML(n.attrs.xml)]); });
  assert.equal(out.length, 8);
  for (const [kind, m] of out) {
    assert.ok(m, kind);
    assert.equal(m.startsWith('<math display="block">'), kind === 'oMathPara');
  }
});

test('equations are saved as they were read', () => {
  const s = new DocxSession(fs.readFileSync(fixture));
  const before = strFromU8(unzipSync(fs.readFileSync(fixture))['word/document.xml']).match(/<m:oMath(Para)?>[\s\S]*?<\/m:oMath(Para)?>/g);
  const after = strFromU8(unzipSync(s.save())[s.pkg.mainPath]).match(/<m:oMath(Para)?>[\s\S]*?<\/m:oMath(Para)?>/g);
  assert.deepEqual(after, before);
});
