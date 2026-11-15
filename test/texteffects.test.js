import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxSession } from '../src/ooxml/docx.js';
import { w14Color, textEffectsCss, setThemeColors } from '../src/ooxml/texteffects.js';
import { WORDART, wordArtItems } from '../src/ooxml/wordart.js';

// Word 2010's text effects as Word writes them (odx-test/mkeffects14.ps1: Word's own XML for each effect set through
// its Font object).
const items = (...xs) => xs.map((x) => [/^<w14:(\w+)/.exec(x)[1], x]);

test('w14 colours: a transparency (w14:alpha is Word\'s transparency, not an opacity), theme colours and their changes', () => {
  setThemeColors({ accent1: '156082', dk1: '000000', lt1: 'FFFFFF', accent5: 'A02B93' });
  assert.equal(w14Color('<w14:srgbClr w14:val="FF0000"><w14:alpha w14:val="50000"/></w14:srgbClr>'), 'rgba(255,0,0,0.5)');
  assert.equal(w14Color('<w14:schemeClr w14:val="bg1"/>'), '#FFFFFF');
  assert.equal(w14Color('<w14:schemeClr w14:val="accent1"/>'), '#156082');
  assert.notEqual(w14Color('<w14:schemeClr w14:val="accent5"><w14:lumMod w14:val="60000"/><w14:lumOff w14:val="40000"/></w14:schemeClr>'), '#A02B93');
});

test('fill, outline, shadow and glow as Word draws them, and the room a glow or reflection takes', () => {
  // Word's 1.5pt blue outline round white letters, its 4pt-down-and-right grey shadow blurred 3pt, a 10pt glow.
  const css = textEffectsCss(items(
    '<w14:textOutline w14:w="19050" w14:cap="rnd" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill><w14:srgbClr w14:val="0070C0"/></w14:solidFill><w14:prstDash w14:val="solid"/><w14:bevel/></w14:textOutline>',
    '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="FFFFFF"/></w14:solidFill></w14:textFill>',
  )).join(';');
  assert.match(css, /-webkit-text-stroke:1.5pt #0070C0/);
  assert.match(css, /-webkit-text-fill-color:#FFFFFF/);
  const sh = textEffectsCss(items('<w14:shadow w14:blurRad="38100" w14:dist="71843" w14:dir="2700113" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="b"><w14:srgbClr w14:val="808080"><w14:alpha w14:val="30000"/></w14:srgbClr></w14:shadow>')).join(';');
  assert.match(sh, /text-shadow:4pt 4pt 3pt rgba\(128,128,128,0.7\)/);
  assert.doesNotMatch(sh, /--odx-grow/);
  const glow = textEffectsCss(items('<w14:glow w14:rad="127000"><w14:srgbClr w14:val="00B050"><w14:alpha w14:val="60000"/></w14:srgbClr></w14:glow>')).join(';');
  assert.match(glow, /0 0 10pt rgba\(0,176,80,0.4\)/);
  assert.match(glow, /--odx-grow-up:0.1278em;--odx-grow-down:0.1722em/);
  const refl = textEffectsCss(items('<w14:reflection w14:blurRad="6350" w14:stA="55000" w14:stPos="0" w14:endA="300" w14:endPos="45500" w14:dist="0" w14:dir="5400000" w14:fadeDir="5400000" w14:sx="100000" w14:sy="-100000" w14:kx="0" w14:ky="0" w14:algn="bl"/>')).join(';');
  assert.match(refl, /--odx-grow-down:0.257em/);
  assert.equal(textEffectsCss(items('<w14:textFill><w14:noFill/></w14:textFill>')).join(';'), '-webkit-text-fill-color:transparent');
});

test('WordArt: a floating text box as wide as its text, its runs in the style\'s effects, read back with them', () => {
  const session = DocxSession.blank();
  assert.equal(WORDART.length, 15);
  for (let i = 0; i < WORDART.length; i++) assert.equal(wordArtItems(i).map(([, x]) => x).join(''), WORDART[i][2], WORDART[i][0]);
  const r = session.addWordArt('Your text here', 3, 295, 84, { x: 86, y: 0 });
  assert.match(r.xml, /<wp:anchor\b[^>]*behindDoc="0"[\s\S]*<wp:wrapNone\/>/);
  assert.match(r.xml, /<wps:cNvSpPr txBox="1"\/>[\s\S]*<a:noFill\/><a:ln><a:noFill\/><\/a:ln>/);
  assert.match(r.xml, /<wps:bodyPr\b[^>]*wrap="none"[\s\S]*<a:spAutoFit\/>/);
  assert.match(r.xml, /<w:rPr xmlns:w14="[^"]+"><w:color w:val="FFFFFF" w:themeColor="background1"\/><w:sz w:val="72"\/>[\s\S]*<w14:glow w14:rad="63500">/);
  assert.equal(r.width, 295);
  assert.match(r.place, /position:absolute/);
  // Its text box's text as the editor edits it, its effects kept.
  const box = session.textBox(r.xml, 0);
  assert.equal(box.textContent, 'Your text here');
  const back = session.setTextBox(r.xml, box, 0);
  assert.match(back.xml, /<w14:glow\b[\s\S]*<w14:textOutline\b[\s\S]*<w14:textFill\b/);
});
