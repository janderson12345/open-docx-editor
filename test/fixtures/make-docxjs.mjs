// Builds test fixtures with the `docx` library (MIT). Run: node test/fixtures/make-docxjs.mjs
import fs from 'node:fs';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, ExternalHyperlink, ImageRun, AlignmentType, PageBreak, InsertedTextRun, DeletedTextRun, VerticalMergeType, FootnoteReferenceRun, LevelFormat, WidthType } from 'docx';
import { deflateSync } from 'node:zlib';

function png(w, h, rgb) {
  const crc = (b) => { let c, t = []; for (let n = 0; n < 256; n++) { c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } let x = 0xffffffff; for (const v of b) x = t[(x ^ v) & 0xff] ^ (x >>> 8); return (x ^ 0xffffffff) >>> 0; };
  const chunk = (type, data) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  const raw = Buffer.alloc((w * 3 + 1) * h); for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) raw.set(rgb, y * (w * 3 + 1) + 1 + x * 3);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

const cell = (text, opts = {}) => new TableCell({ children: [new Paragraph(text)], ...opts });
const doc = new Document({
  creator: 'open-docx-editor tests',
  footnotes: { 1: { children: [new Paragraph('A footnote.')] } },
  numbering: { config: [{ reference: 'nums', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.START }, { level: 1, format: LevelFormat.LOWER_LETTER, text: '%2)', alignment: AlignmentType.START }] }] },
  sections: [{
    children: [
      new Paragraph({ text: 'Quarterly Report', heading: HeadingLevel.TITLE }),
      new Paragraph({ text: 'Introduction', heading: HeadingLevel.HEADING_1 }),
      new Paragraph({ children: [new TextRun('Plain, '), new TextRun({ text: 'bold, ', bold: true }), new TextRun({ text: 'italic, ', italics: true }), new TextRun({ text: 'underlined, ', underline: {} }), new TextRun({ text: 'red 14pt', color: 'FF0000', size: 28 }), new TextRun({ text: ' and Courier.', font: 'Courier New' }), new TextRun({ text: '2', superScript: true }), new FootnoteReferenceRun(1)] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun('Centered with a '), new ExternalHyperlink({ link: 'https://example.com/', children: [new TextRun({ text: 'hyperlink', style: 'Hyperlink' })] }), new TextRun('.')] }),
      new Paragraph({ text: 'Lists', heading: HeadingLevel.HEADING_2 }),
      new Paragraph({ text: 'Bullet one', bullet: { level: 0 } }),
      new Paragraph({ text: 'Bullet nested', bullet: { level: 1 } }),
      new Paragraph({ text: 'Step one', numbering: { reference: 'nums', level: 0 } }),
      new Paragraph({ text: 'Sub step', numbering: { reference: 'nums', level: 1 } }),
      new Paragraph({ text: 'Step two', numbering: { reference: 'nums', level: 0 } }),
      new Paragraph({ children: [new TextRun('Tracked: '), new InsertedTextRun({ text: 'added', id: 1, author: 'Ann', date: '2026-01-01T00:00:00Z' }), new DeletedTextRun({ text: 'removed', id: 2, author: 'Bob', date: '2026-01-01T00:00:00Z' }), new TextRun('\tafter tab')] }),
      new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [3000, 3000, 3000], rows: [
        new TableRow({ children: [cell('A1'), cell('B1 spans', { columnSpan: 2 })] }),
        new TableRow({ children: [cell('A2 tall', { verticalMerge: VerticalMergeType.RESTART }), cell('B2'), cell('C2')] }),
        new TableRow({ children: [cell('', { verticalMerge: VerticalMergeType.CONTINUE }), cell('B3'), cell('C3')] }),
      ] }),
      new Paragraph({ children: [new ImageRun({ type: 'png', data: png(40, 20, [30, 120, 220]), transformation: { width: 120, height: 60 }, altText: { name: 'blue box', description: 'A blue box', title: 'box' } })] }),
      new Paragraph({ children: [new PageBreak()] }),
      new Paragraph({ children: [new TextRun({ text: 'Second page with <special> & "chars" é ü 中文', highlight: 'yellow' })] }),
    ],
  }],
});
fs.writeFileSync(new URL('./docxjs-rich.docx', import.meta.url), await Packer.toBuffer(doc));
console.log('wrote docxjs-rich.docx');
