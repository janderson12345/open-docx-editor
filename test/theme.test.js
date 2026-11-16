// A document's theme colours (styles.js readTheme) are read only as six hex digits: they go into styles and XML as they
// are, so a crafted theme's CSS or markup is dropped and the Office default stands in.
import './setup.js';
import test from 'node:test';
import assert from 'node:assert/strict';
import { readTheme } from '../src/ooxml/styles.js';
import { parseXml } from '../src/ooxml/xml.js';

const pkgWith = (scheme) => {
  const xml = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x">${scheme}</a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;
  return { partByType: () => 'word/theme/theme1.xml', xml: () => parseXml(xml) };
};

test('theme colours: six hex digits kept, CSS or markup in one dropped', () => {
  const theme = readTheme(pkgWith('<a:dk1><a:sysClr val="windowText" lastClr="000000&quot;/&gt;&lt;x"/></a:dk1>'
    + '<a:accent1><a:srgbClr val="1a2B3c"/></a:accent1><a:accent2><a:srgbClr val="FF0000;background:url(https://x.test/)"/></a:accent2>'
    + '<a:accent3><a:srgbClr val="ABCDE"/></a:accent3>'));
  assert.equal(theme.colors.accent1, '1a2B3c');
  assert.equal(theme.colors.accent2, undefined);
  assert.equal(theme.colors.accent3, undefined);
  assert.equal(theme.colors.dk1, undefined);
});
