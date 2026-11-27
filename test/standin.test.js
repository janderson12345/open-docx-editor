import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { lineFactor, setFontAlts } from '../src/ooxml/props.js';
import { fontMetrics } from '../src/fonts/wordmetrics.js';

// Futura Lt BT, missing, prints in Century Gothic with Century Gothic's own line (1.2261 em), also where the font
// table's altName names Century Gothic (58 samples): not the 1.19 an installed Century Gothic measures.
test("Futura in Century Gothic keeps Word's Century Gothic line where the font table names Century Gothic too", () => {
  setFontAlts(new Map([['Futura Lt BT', 'Century Gothic']]));
  assert.equal(lineFactor('Futura Lt BT'), fontMetrics('century gothic').line);
});
