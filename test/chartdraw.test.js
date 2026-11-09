import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { chartSvg } from '../src/ooxml/chart.js';
import { writeChart, readChart } from '../src/ooxml/chartedit.js';

const part = strFromU8(unzipSync(fs.readFileSync(new URL('./fixtures/chart.docx', import.meta.url)))['word/charts/chart1.xml']);

test('a line leaves a gap at a blank value, as Word does, rather than dropping to zero', () => {
  const d = readChart(part);
  const xml = writeChart(part, { ...d, type: 'line', title: null, series: [{ name: 'A', vals: [1, null, 3] }] });
  const svg = chartSvg(xml, { width: 400, height: 240 });
  const path = /<path d="([^"]+)" fill="none"/.exec(svg)[1];
  assert.equal((path.match(/M/g) || []).length, 2, path);
});
