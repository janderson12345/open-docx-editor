// Browser check of the smaller things Word draws (test/fixtures/mkfeatures.mjs, mkrtl.mjs, mkstrict.mjs), each laid
// out as Word's own PDF of it showed: picture bullets, line numbers, watermarks, a section's vertical alignment,
// emphasis marks and run borders making their lines taller, right-to-left paragraphs, a Strict document.
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';

const root = path.resolve('demo');
const server = http.createServer((req, res) => {
  const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' }[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const launch = process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {};
const browser = await chromium.launch(launch);
const page = await browser.newPage({ viewport: { width: 1300, height: 950 } });
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
let failed = 0;
const check = (ok, what) => { console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}`); if (!ok) failed++; };
await page.goto(`http://localhost:${server.address().port}/`);
await page.waitForSelector('.odx-content p');
const open = async (f) => {
  await page.setInputFiles('input[type=file]', `test/fixtures/${f}`);
  await page.waitForFunction((n) => window.odx?.fileName === n, f);
  await page.waitForTimeout(1200);
};

await open('numbering-legal.docx');
const labels = await page.evaluate(() => [...document.querySelectorAll('.odx-content p.odx-li')].map((p) => p.dataset.label));
check(labels.slice(0, 4).join(' ') === 'I. 1.1. 1.1.1 I.1.a.1', `legal numbering: ${labels.slice(0, 4).join(' ')}`);
const pic = await page.evaluate(() => { const p = document.querySelector('.odx-content p.odx-picbullet'); return p && getComputedStyle(p, '::before').content; });
check(/^url\("data:image\/svg\+xml/.test(pic || ''), 'a picture bullet shows its picture');

// Bold typed at the end of a list item carries on into the next item made with Enter, as in Word.
await page.click('.odx-content p.odx-li');
await page.keyboard.press('End');
await page.waitForTimeout(100); // (the caret move End makes reaches the editor as a selection change, after the key)
await page.keyboard.press('Control+b');
await page.keyboard.type(' more');
await page.keyboard.press('Enter');
await page.keyboard.type('carried');
await page.waitForTimeout(300);
const carried = await page.evaluate(() => { const w = document.createTreeWalker(document.querySelector('.odx-content'), NodeFilter.SHOW_TEXT); let n; while ((n = w.nextNode())) if (n.textContent === 'carried') return getComputedStyle(n.parentElement).fontWeight; return 'missing'; });
check(carried === '700', `bold carries into the list item Enter makes (${carried})`);

await open('line-numbers.docx');
const nums = await page.evaluate(() => [...document.querySelectorAll('.odx-lnnum')].map((s) => s.textContent));
// (Word's PDF: 32 on the first page, 10 on the second, then every second line (12 to 22) and 5 to 10 from w:start.)
check(nums.slice(0, 5).join(',') === '1,2,3,4,5' && nums.length === 54 && nums.slice(42, 48).join(',') === '12,14,16,18,20,22', `line numbers: ${nums.length} drawn, as Word's`);
const lnBox = await page.evaluate(() => { const s = document.querySelector('.odx-lnnum').getBoundingClientRect(); const c = document.querySelector('.odx-content').getBoundingClientRect(); return [s.right, c.left + parseFloat(getComputedStyle(document.querySelector('.odx-content')).paddingLeft)]; });
check(Math.abs(lnBox[1] - lnBox[0] - 24) < 1, `a line number ends 0.25in left of the text (${(lnBox[1] - lnBox[0]).toFixed(1)}px)`);

await open('watermark-text.docx');
const wm = await page.evaluate(() => [...document.querySelectorAll('.odx-sheet > img')].map((i) => ({ t: getComputedStyle(i).transform, w: i.width })));
// (550px: the shape's 412.4pt width.)
const sheets = await page.evaluate(() => document.querySelectorAll('.odx-sheets > .odx-sheet').length);
check(wm.length === sheets && wm.every((x) => x.t !== 'none' && x.w === 550), `a text watermark on each of the ${sheets} pages, turned (${wm.length})`);

await open('valign.docx');
const vtop = await page.evaluate(() => { const p = document.querySelector('.odx-content > p'); const s = document.querySelector('.odx-sheet'); return p.getBoundingClientRect().top - s.getBoundingClientRect().top; });
check(vtop > 350 && vtop < 500, `a centred section's text sits in the middle of the page (${vtop.toFixed(0)}px down)`);

await open('effects.docx');
// (Two lines of 11pt Calibri at 1.08 are 38.7px; the marks add 0.31em, 4.5px, above the first or below it.)
const grow = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].slice(6, 10).map((p) => p.getBoundingClientRect().height));
check(grow.every((h) => h > 42.5 && h < 45), `emphasis marks make their line taller (${grow.map((h) => h.toFixed(1)).join(', ')}px)`);

await open('notes-long.docx');
// (Four notes of 45 paragraphs each, from one paragraph: they continue at the foot of the pages after it, as in Word,
// each page's notes inside its own margins, never over its body text.)
const notes = await page.evaluate(() => [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => {
  const n = s.querySelector('.odx-notes');
  return n ? n.getBoundingClientRect().top - s.getBoundingClientRect().top : null;
}));
check(notes.length >= 5 && notes.every((t) => t != null && t >= 96 + 20), `long footnotes continue over ${notes.length} pages, below each page's text (${notes.map((t) => Math.round(t)).join(',')})`);

// An input method (Chinese pinyin) composing with Track Changes on: the composed text is one tracked insertion, and
// nothing of the letters typed while composing is left behind.
await page.evaluate(() => { const s = document.querySelector('.odx-tracksel'); s.value = 'on'; s.dispatchEvent(new Event('change', { bubbles: true })); });
await page.click('.odx-content > p');
await page.evaluate(() => { const r = document.createRange(); r.selectNodeContents(document.querySelector('.odx-content > p')); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); });
const cdp = await page.context().newCDPSession(page);
for (const s of ['z', 'zh', 'zhong']) { await cdp.send('Input.imeSetComposition', { text: s, selectionStart: s.length, selectionEnd: s.length }); await page.waitForTimeout(60); }
await cdp.send('Input.insertText', { text: '中文' });
await page.waitForTimeout(400);
const ime = await page.evaluate(() => ({ text: document.querySelector('.odx-content > p').textContent, ins: [...document.querySelectorAll('.odx-content ins.odx-ins')].map((e) => e.textContent).join('|') }));
check(ime.text.endsWith('中文') && !/z/.test(ime.text.slice(-8)) && ime.ins === '中文', `IME input with Track Changes on is one tracked insertion (${JSON.stringify(ime)})`);
await page.evaluate(() => { const s = document.querySelector('.odx-tracksel'); s.value = 'off'; s.dispatchEvent(new Event('change', { bubbles: true })); });

await open('rtl-hebrew.docx');
const rtl = await page.evaluate(() => {
  const p = document.querySelectorAll('.odx-content > p')[1];
  const r = document.createRange(); r.selectNodeContents(p);
  const c = document.querySelector('.odx-content');
  return [getComputedStyle(p).direction, c.getBoundingClientRect().right - parseFloat(getComputedStyle(c).paddingRight) - Math.max(...[...r.getClientRects()].map((q) => q.right))];
});
check(rtl[0] === 'rtl' && rtl[1] < 2, `a right-to-left paragraph starts at the right margin (${rtl[1].toFixed(1)}px)`);

await open('fields.docx');
// A running head (STYLEREF) worked out for each page, as Word does; the body's fields keep their saved results until
// they're updated (F9), then match Word's (its PDF of the document with its fields updated).
const heads = await page.evaluate(() => [...document.querySelectorAll('.odx-sheet .odx-header')].map((h) => h.textContent.replace(/\s+/g, ' ').trim()));
check(heads.join(' | ') === 'Report — Introduction — page i | Report — Introduction — page 1 | Report — Results — page 2 | Report — Results — page 3', `running heads per page: ${heads.join(' | ')}`);
const bodyText = () => page.evaluate(() => window.odx.view.state.doc.textContent);
check((await bodyText()).includes('chapter 9, "Old title", on page 99'), 'the body keeps its saved field results until they are updated');
const updated = await page.evaluate(() => window.odx.updateFields());
const after = await bodyText();
check(updated >= 5 && after.includes('chapter 2, "Methods of measurement", on page 1') && after.includes('results are on page 2, on page 2, and the introduction on page 1') && after.includes('Error! Reference source not found.'), `Update Fields gives Word's values (${updated} updated)`);
await page.click('.odx-content p >> nth=1');
await page.keyboard.press('Control+z');
check((await bodyText()).includes('on page 99'), 'one Ctrl+Z puts the saved results back');

// An agent working in the editor: its edit appears as a tracked change under its name; the person accepts it.
const agentRun = await page.evaluate(() => {
  const a = window.odx.agent({ author: 'Assistant' });
  const head = a.outline()[2];
  const n = a.replace('Results', 'Findings', { id: head.id });
  const del = document.querySelector('.odx-content del.odx-del[data-author="Assistant"]')?.textContent;
  const ins = document.querySelector('.odx-content ins.odx-ins[data-author="Assistant"]')?.textContent;
  a.resolveChanges({ accept: true });
  return { n, del, ins, after: a.outline()[2].text, left: a.changes().length };
});
check(agentRun.n === 1 && agentRun.del === 'Results' && agentRun.ins === 'Findings' && agentRun.after === 'Findings' && agentRun.left === 0, `an agent's edit shows as a tracked change under its name, then accepted (${JSON.stringify(agentRun)})`);

await open('tiff.docx');
const tiffs = await page.evaluate(() => [...document.querySelectorAll('.odx-content img:not(.ProseMirror-separator)')].map((i) => i.complete && i.naturalWidth > 0));
check(tiffs.length === 3 && tiffs.every(Boolean), `TIFF pictures (uncompressed, LZW, fax) are shown (${tiffs.length})`);

// Restrict Editing, as Word enforces it.
const typeAt = async (text, keys) => { await page.click(`.odx-content p:has-text("${text}")`); await page.keyboard.press('End'); await page.keyboard.type(keys); await page.waitForTimeout(150); };
const note = () => page.evaluate(() => document.querySelector('.odx-protectnote span')?.textContent || '');
await open('protect-readonly.docx');
check(/can only be read/.test(await note()), 'a document protected read-only says so above the ribbon');
await typeAt('This paragraph is locked.', 'XX');
check(!(await bodyText()).includes('XX'), 'read-only: typing in a locked paragraph does nothing');
const perms = await page.evaluate(() => [...document.querySelectorAll('.odx-content .odx-perm')].map((e) => e.textContent).join('|'));
check(perms === 'Everyone may edit this.', `the part everyone may edit is highlighted, not the tester's (${perms})`);
// No keyboard trap: in a locked part, Tab moves the focus on instead of typing (WCAG 2.1.2).
await page.click('.odx-content p:has-text("This paragraph is locked.")');
await page.keyboard.press('Tab');
check(await page.evaluate(() => !document.activeElement.classList.contains('odx-content')), 'read-only: Tab leaves the document');
await page.click('.odx-content .odx-perm');
await page.keyboard.type('YY');
await page.waitForTimeout(150);
check((await bodyText()).includes('YY'), 'read-only: the part everyone may edit can be typed in');
check(await page.evaluate(() => window.odx.view.odx.session.comments.length === 0 && (document.querySelector('.odx-content p').click(), true)), 'no comments yet');
await page.keyboard.press('Control+Alt+m');
check(await page.evaluate(() => window.odx.view.odx.session.comments.length === 0), 'read-only: no comment can be added');
// Comments only, with a password; then stopped with it.
check(await page.evaluate(async () => { await window.odx.unprotect(); await window.odx.protect({ edit: 'comments', password: 'pw' }); return window.odx.protection.edit === 'comments' && window.odx.protection.password; }), 'Restrict Editing: protected for comments with a password');
await typeAt('This paragraph is locked too.', 'ZZ');
await page.keyboard.press('Control+Alt+m');
await page.waitForTimeout(200);
check(!(await bodyText()).includes('ZZ') && (await page.evaluate(() => window.odx.view.odx.session.comments.length)) === 1, 'comments only: no typing, but a comment is added');
check(await page.evaluate(async () => !(await window.odx.unprotect('wrong')) && window.odx.protection.enforced && (await window.odx.unprotect('pw')) && !window.odx.protection.enforced), 'Stop Protection needs the password');
// Tracked changes only: every edit tracked, Track Changes can't be turned off, nothing accepted.
await page.evaluate(() => window.odx.protect({ edit: 'trackedChanges' }));
await typeAt('This paragraph is locked.', 'TT');
const tracked = await page.evaluate(() => [...document.querySelectorAll('.odx-content ins.odx-ins')].map((e) => e.textContent).join(''));
await page.evaluate(() => { const t = document.querySelector('.odx-tracksel'); t.value = 'off'; t.dispatchEvent(new Event('change')); });
await page.waitForTimeout(150);
const stillOn = await page.evaluate(() => document.querySelector('.odx-tracksel').value);
await page.evaluate(() => window.odx.view.focus());
check(tracked.includes('TT') && stillOn === 'on', `tracked changes only: typing is tracked and Track Changes stays on (${tracked}, ${stillOn})`);
await page.evaluate(() => window.odx.unprotect());
// A protected form with an unprotected section.
await open('protect-sections.docx');
await typeAt('can be edited freely.', ' Yes.');
await typeAt('This text of the form is locked.', 'QQ');
// (A click on the form's own text puts the caret in its next field, as in Word: the typing goes there.)
const formText = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].map((p) => p.textContent));
check(formText[0].endsWith('freely. Yes.') && formText[3] === 'This text of the form is locked.', `a protected form: its unprotected section is edited, the form's own text is not (${formText.join(' / ')})`);
// Formatting limited to the allowed styles.
await open('protect-format.docx');
const styleNames = await page.evaluate(() => [...document.querySelectorAll('.odx-stile')].map((t) => t.title));
check(styleNames.includes('Heading 1') && !styleNames.includes('Heading 2'), 'limited formatting: a locked style is not offered');
await page.click('.odx-content p:has-text("Some text")');
await page.keyboard.press('Control+a');
await page.keyboard.press('Control+b');
await page.waitForTimeout(150);
await page.keyboard.press('Control+Home');
check(await page.evaluate(() => !document.querySelector('.odx-content strong, .odx-content b')), 'limited formatting: no direct formatting (Ctrl+B does nothing)');
await typeAt('Some text', ' More.');
check((await bodyText()).includes('formatting. More.'), 'limited formatting: the text can still be edited');
// Mark as Final: read-only until Edit Anyway, which takes the mark away.
await open('protect-final.docx');
await typeAt('marked as final', 'FF');
const finalNote = await note();
await page.click('.odx-protectnote button');
await typeAt('marked as final', 'GG');
check(/marked this document as final/.test(finalNote) && !(await bodyText()).includes('FF') && (await bodyText()).includes('GG'), 'Mark as Final: read-only, then Edit Anyway');

// Table of contents lines (mktoc.mjs): the page number sits right-aligned at the margin on its tab's line, also when
// the title leaves no room for it, where the tab starts the next line at the indent, as in Word's PDF of the file.
await open('toc-wrap.docx');
const toc = await page.evaluate(() => [...document.querySelectorAll('.odx-content p')].filter((p) => p.querySelectorAll('.odx-tab').length === 2).map((p) => {
  const tab = p.querySelectorAll('.odx-tab')[1].getBoundingClientRect(), pr = p.getBoundingClientRect();
  const num = [...p.querySelectorAll('.odx-t')].pop().getBoundingClientRect();
  return { sameLine: num.bottom > tab.top && num.top < tab.bottom, atRight: Math.abs(num.right - pr.right) < 2, startsLine: Math.abs(tab.left - pr.left) < 1, at: [p.textContent.slice(0, 24), Math.round(num.top), Math.round(tab.top), Math.round(num.right), Math.round(pr.right), getComputedStyle(p).fontFamily.slice(0, 30), p.querySelectorAll('.odx-tab')[1].style.cssText, [...p.querySelectorAll('.odx-t')].map((t) => `${t.textContent.slice(0, 12)}@${Math.round(t.getBoundingClientRect().left)}-${Math.round(t.getBoundingClientRect().right)}`).join(' '), (() => { const r = document.createRange(); r.setStartAfter(p.querySelectorAll('.odx-tab')[1]); r.setEnd(p, p.childNodes.length); return [...r.getClientRects()].map((x) => `${Math.round(x.left)}+${x.width.toFixed(1)}@${Math.round(x.top)}`).join(' '); })(), [...p.childNodes].slice(-4).map((n) => n.nodeName + (n.className ? `.${n.className}` : '')).join(' '),
    // (What the tab's width would be from the page as it is now: right if it was measured too early.)
    (() => { const cs = getComputedStyle(p), t = p.querySelectorAll('.odx-tab')[1]; const edge = pr.left - parseFloat(cs.marginLeft); const x = t.getBoundingClientRect().left - edge; const right = pr.right - parseFloat(cs.paddingRight) - edge; return { x: +x.toFixed(2), right: +right.toFixed(2), stop: +(9350 / 15).toFixed(2), rw: +num.width.toFixed(2), now: +(Math.min(9350 / 15, right) - x - num.width - 0.5).toFixed(2), dpr: devicePixelRatio, zoom: getComputedStyle(document.querySelector('.odx-sheet') || p).transform }; })()] };
}));
const tocBad = toc.find((x) => !(x.sameLine && x.atRight));
check(toc.length === 30 && !tocBad, `TOC: each page number right-aligned on its tab's line (${toc.filter((x) => x.sameLine && x.atRight).length}/${toc.length}${tocBad ? `; first not: ${JSON.stringify(tocBad)}` : ''})`);
check(toc.some((x) => x.startsLine), 'TOC: a title with no room left takes its tab and number to the next line');

// Find goes through every part of the document as Word's does (mkfindall.mjs): the main text, a text box, the header
// and footer, a footnote and an endnote, opening each part with its match selected; Replace All changes them all.
await open('find-all.docx');
await page.click('.odx-content p');
await page.keyboard.press('Control+Home');
await page.keyboard.press('Control+f');
await page.fill('.odx-findbar input[aria-label="Find"]', 'apple');
const visits = [];
const findSaid = []; // (what the find bar says at each, for screen readers: which match, and where)
for (let i = 0; i < 9; i++) {
  await page.press('.odx-findbar input[aria-label="Find"]', 'Enter');
  await page.waitForTimeout(250);
  findSaid.push(await page.textContent('.odx-findinfo'));
  findSaid[findSaid.length - 1] += await page.evaluate(() => (document.activeElement?.getAttribute('aria-label') === 'Find' ? '' : ' (focus left the find box)'));
  visits.push(await page.evaluate(() => {
    const ta = document.querySelector('.odx-markup textarea[data-odx-box^="e"]');
    if (ta) return `comment:${ta.value.slice(ta.selectionStart, ta.selectionEnd)}`;
    const v = window.odx.view.odx.hfView || window.odx.view;
    const { from, to } = v.state.selection;
    const part = v.dom.closest('.odx-txbxedit') ? 'box' : v.dom.closest('.odx-hfedit-header') ? 'header' : v.dom.closest('.odx-hfedit-footer') ? 'footer'
      : v.dom.closest('.odx-noteedit') ? (v.state.doc.textContent.includes('endnote') ? 'endnote' : 'footnote') : 'body';
    return `${part}:${v.state.doc.textBetween(from, to)}`;
  }));
}

const info = await page.textContent('.odx-findinfo');
check(findSaid[3] === '4 of 8, in the header' && findSaid[5] === '6 of 8, in a footnote' && findSaid[7] === '8 of 8, in a comment', `the find bar says which match and where (${findSaid.join(' | ')})`);
check(visits.join(' ') === 'body:apple body:apple box:apple header:apple footer:apple footnote:apple endnote:apple comment:apple body:apple' && info === '1 of 8',
  `Find goes through the body, text box, header, footer, notes and comments (${visits.join(' ')}; ${info})`);
await page.keyboard.press('Control+h');
await page.fill('.odx-findbar input[aria-label="Replace with"]', 'pear');
await page.click('.odx-findbar button:has-text("Replace all")');
await page.waitForTimeout(500);
const replaced = await page.textContent('.odx-findinfo');
const saved = unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))));
const partText = (f) => strFromU8(saved[f]).replace(/<w:t\b[^>]*>([^<]*)<\/w:t>|<[^>]+>/g, '$1');
const parts = ['word/document.xml', 'word/header1.xml', 'word/footer1.xml', 'word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml'];
const pears = parts.map((f) => (partText(f).match(/pear/g) || []).length);
check(replaced === '8 results' && parts.every((f) => !partText(f).includes('apple')) && pears.join() === '4,1,1,1,1,1',
  `Replace All in every part, saved (${replaced}; pears per part ${pears.join()}, the text box's twice: its shape and fallback copy)`);
await page.keyboard.press('Escape');

// Columns of their own widths (mkcolumns.mjs: 1.5in, 3in and 1.25in), laid out as Word's PDF of the file has them: each
// column at its place and width, a paragraph going on in a wider column laid out again at that width (P9's last
// lines: two at 3in), P24 starting the third column and P30 ending the first page.
await open('columns-unequal.docx');
const ucols = await page.evaluate(() => {
  const sheet = document.querySelector('.odx-sheet').getBoundingClientRect();
  const at = (n) => { const p = [...document.querySelectorAll('.odx-content > p')].find((x) => x.textContent.startsWith(`P${n} `)); const r = p.getBoundingClientRect(); return { x: Math.round(r.left - sheet.left), w: Math.round(r.width), top: Math.round(r.top - sheet.top) }; };
  const cont = document.querySelector('.odx-repeat .odx-contlines');
  return { p1: at(1), p10: at(10), p24: at(24), p30: at(30), p31: at(31), cont: cont && { w: Math.round(cont.getBoundingClientRect().width), lines: Math.round(cont.getBoundingClientRect().height / 19.33), text: cont.textContent.slice(0, 15) }, pages: window.odx.view.odx.pages.count };
});
check(ucols.p1.x === 96 && ucols.p1.w === 144 && ucols.p10.x === 264 && ucols.p10.w === 288 && ucols.p24.x === 600 && ucols.p24.w === 120,
  `columns of their own widths: each at its place and width (${JSON.stringify([ucols.p1, ucols.p10, ucols.p24])})`);
check(ucols.cont?.w === 288 && ucols.cont.lines === 2 && ucols.cont.text.startsWith('procedures'), `a paragraph going on in a wider column is laid out again at its width (${JSON.stringify(ucols.cont)})`);
check(ucols.pages === 3 && ucols.p30.top < 1056 && ucols.p31.top > 1056, `columns of their own widths: P30 ends the first page, as in Word (${ucols.pages} pages)`);

// Paragraphs in frames (mkframes.mjs) where Word's PDF of the file has them: frame A's text 2in down the page and 4.5in
// across, its two paragraphs one frame; sidebar B at the right of the text, Body 6 beside it; note C in the left margin.
await open('frames.docx');
const fr = await page.evaluate(() => {
  const sheet = document.querySelector('.odx-sheet').getBoundingClientRect();
  const para = (t) => [...document.querySelectorAll('.odx-content > p')].find((x) => x.textContent.startsWith(t));
  const text = (t) => { const r = document.createRange(); r.selectNodeContents(para(t)); const b = r.getClientRects()[0]; return { x: Math.round(b.left - sheet.left), y: Math.round(b.top - sheet.top) }; };
  const box = (t) => { const b = para(t).getBoundingClientRect(); return { left: Math.round(b.left - sheet.left), right: Math.round(b.right - sheet.left), top: Math.round(b.top - sheet.top) }; };
  const r = document.createRange(); r.selectNodeContents(para('Body 6.')); const body6 = Math.max(...[...r.getClientRects()].filter((x) => x.top < para('Frame B').getBoundingClientRect().bottom).map((x) => x.right)) - sheet.left;
  return { a1: text('Frame A, first'), a2: box('Frame A, second'), a1box: box('Frame A, first'), b: box('Frame B'), c: box('Note C'), body6: Math.round(body6) };
});
check(Math.abs(fr.a1.x - 432) <= 2 && fr.a1box.top === 192 && fr.a2.left === fr.a1box.left && fr.a2.top > fr.a1box.top,
  `a frame set down on the page: its text 4.5in across, the frame 2in down, its second paragraph below (${JSON.stringify([fr.a1, fr.a1box, fr.a2])})`);
check(fr.b.right === 720 && fr.body6 < fr.b.left && fr.c.left === 12 && fr.c.right <= 96,
  `a frame at the right of the text with the text beside it, and one in the margin (${JSON.stringify([fr.b, fr.body6, fr.c])})`);

// The Original view (mkoriginal.mjs): the document before its tracked changes, as Word's Review > Original shows it.
await open('original.docx');
const shown = () => page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].filter((p) => getComputedStyle(p).display !== 'none').map((p) => p.innerText.replace(/\s+/g, ' ').trim()).filter(Boolean));
await page.click('.odx-rtab:text-is("Review")');
await page.selectOption('.odx-markupsel', 'original');
await page.waitForTimeout(400);
const orig = await shown();
await page.selectOption('.odx-markupsel', 'none');
await page.waitForTimeout(400);
const final = await shown();
await page.selectOption('.odx-markupsel', 'simple');
await page.click('.odx-rtab:text-is("Home")');
check(orig.join(' | ') === 'Alpha words old end. | A paragraph deleted whole. | Split here and joined again. | First of two | Second of two.',
  `Original: insertions out, deletions in, paragraphs as they were (${orig.join(' | ')})`);
check(final.join(' | ') === 'Alpha new words end. | A paragraph inserted whole. | Split here by Enter | and joined again. | First of twoSecond of two.',
  `No Markup still shows the changed document (${final.join(' | ')})`);

// Wrap Text (Layout > Wrap Text): the selected picture floats with the text square round it, goes behind the text, and
// back in line; each drawn so and saved as Word writes it.
await open('rich.docx');
await page.click('img[alt="A blue box"]'); // (selected; Wrap Text keeps it selected)
const wrapTo = async (name) => {
  await page.click('.odx-rtab:text-is("Layout")');
  await page.click('button[title^="Wrap Text"]');
  await page.click(`.odx-item:has-text("${name}")`);
  await page.waitForTimeout(400);
  const look = await page.evaluate(() => { const cs = getComputedStyle(document.querySelector('img[alt="A blue box"]')); return `${cs.float}/${cs.position}/${cs.zIndex}`; });
  const saved = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']);
  const at = saved.indexOf('name="blue box"');
  const open = Math.max(saved.lastIndexOf('<wp:inline', at), saved.lastIndexOf('<wp:anchor', at));
  const head = open < 0 ? '' : saved.slice(open, at); // (the drawing's start up to its name: its wrap element comes before)
  const wrapEl = /<wp:(wrap(?!Polygon)\w+)/.exec(head)?.[1];
  const saidSo = [head.startsWith('<wp:anchor') ? 'anchor' : 'inline', wrapEl, /behindDoc="1"/.test(head) && 'behind'].filter(Boolean).join(' ');
  return { look, saved: head ? saidSo : 'missing' };
};
const sq = await wrapTo('Square');
const behind = await wrapTo('Behind Text');
const inl = await wrapTo('In Line with Text');
await page.click('.odx-rtab:text-is("Home")');
check(sq.look.startsWith('left/') && sq.saved === 'anchor wrapSquare', `Wrap Text Square: floated, saved square (${JSON.stringify(sq)})`);
check(behind.look === 'none/absolute/-1' && behind.saved === 'anchor wrapNone behind', `Wrap Text Behind Text: behind the text, saved so (${JSON.stringify(behind)})`);
check(inl.look.startsWith('none/static') && inl.saved === 'inline', `Wrap Text In Line with Text: back in line (${JSON.stringify(inl)})`);

// Layout > Page Setup (mkpagesetup.mjs, two sections): margins, orientation and paper for the section the cursor is in,
// then the whole document, saved in each section's w:sectPr.
await open('page-setup.docx');
const sectsSaved = async () => {
  const xml = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']);
  return [...xml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((m) => {
    const a = (tag, k) => new RegExp(`<w:${tag}\\b[^>]*\\sw:${k}="(\\d+)"`).exec(m[0])?.[1];
    return `${a('pgSz', 'w')}x${a('pgSz', 'h')}${/w:orient="landscape"/.test(m[0]) ? 'L' : ''} t${a('pgMar', 'top')} l${a('pgMar', 'left')}`;
  });
};
const setup = async (vals, all) => {
  await page.click('.odx-content p');
  await page.keyboard.press('Control+Home');
  await page.click('.odx-rtab:text-is("Layout")');
  await page.click('button[title^="Page Setup"]');
  for (const [label, v] of Object.entries(vals)) {
    const sel = `.odx-pagesetup [aria-label="${label}"]`;
    if ((await page.$eval(sel, (e) => e.tagName)) === 'SELECT') await page.selectOption(sel, v); else await page.fill(sel, v);
  }
  await page.selectOption('.odx-pagesetup [aria-label="Apply to"]', all ? 'all' : 'section');
  await page.click('.odx-pagesetup button:has-text("OK")');
  await page.waitForFunction(() => window.odx?.fileName === 'page-setup.docx');
  await page.waitForTimeout(800);
};
await setup({ 'Left (inches)': '1.5', Orientation: 'landscape', 'Paper size': 'A4' }, false);
const one = await sectsSaved();
await setup({ 'Top (inches)': '0.5' }, true);
const both = await sectsSaved();
await page.click('.odx-rtab:text-is("Home")');
check(one.join(' / ') === '16838x11906L t1440 l2160 / 12240x15840 t1440 l1440', `Page Setup, this section: A4 landscape, left 1.5in, the other section as it was (${one.join(' / ')})`);
check(both.join(' / ') === '16838x11906L t720 l2160 / 16838x11906L t720 l2160', `Page Setup, whole document: the dialog's page in both sections (${both.join(' / ')})`);

// Imported content (w:altChunk, mkaltchunk.mjs): drawn as the paragraphs Word makes of it, in the cells and the body,
// with no placeholder left; Find finds its text.
await open('altchunk.docx');
const chunk = await page.evaluate(() => ({
  opaque: document.querySelectorAll('.odx-content .odx-opaque').length,
  cell: [...document.querySelectorAll('.odx-content td')].map((td) => td.textContent).find((t) => t.startsWith('Basis')),
  bullets: [...document.querySelectorAll('.odx-content p.odx-li')].map((p) => p.dataset.label + p.textContent).join('|'),
}));
check(chunk.opaque === 0 && chunk.cell === 'Basis of Estimate: hours from the last three task orders.', `imported content drawn as text, none as a placeholder (${JSON.stringify(chunk.cell)}, ${chunk.opaque} left)`);
check(chunk.bullets.includes('First bullet') && chunk.bullets.includes('1.First step'), `its lists are lists (${chunk.bullets})`);
await page.click('.odx-content p');
await page.keyboard.press('Control+f');
await page.fill('.odx-findbar input[aria-label="Find"]', 'task orders');
await page.press('.odx-findbar input[aria-label="Find"]', 'Enter');
await page.waitForTimeout(300);
check((await page.textContent('.odx-findinfo')) === '1 of 1', 'Find finds imported text');
await page.keyboard.press('Escape');

// List numbers in their level's bold, italic and colour, else the paragraph mark's (mklists2.mjs; as in Word's PDF).
await open('lists-more.docx');
const lbls = await page.evaluate(() => [...document.querySelectorAll('.odx-content p.odx-li')].slice(0, 6).map((p) => { const b = getComputedStyle(p, '::before'); return `${b.fontWeight >= 700 ? 'b' : ''}${b.fontStyle === 'italic' ? 'i' : ''}${b.color === 'rgb(192, 0, 0)' ? 'red' : ''}`; }));
check(lbls.join(',') === 'bred,bred,i,i,b,', `list numbers take their level's (or the paragraph mark's) bold, italic and colour (${lbls.join(',')})`);

// Pasted from a web page or another editor: colour, size and font kept as marks; black (set on every run by some
// editors) and generic families left out.
await page.evaluate(() => window.odx.newDocument({ share: false }));
await page.click('.odx-content p');
const pasted = await page.evaluate(() => {
  const dt = new DataTransfer();
  dt.setData('text/html', '<p><span style="color:rgb(192, 0, 0); font-size:16pt; font-family:Georgia, serif">Red big Georgia</span> <span style="color:#000000; font-family:sans-serif; font-size:1.2em">plain</span></p>');
  dt.setData('text/plain', 'Red big Georgia plain');
  window.odx.view.dom.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  const marks = (word) => { let out = ''; window.odx.view.state.doc.descendants((n) => { if (n.isText && n.text.includes(word)) out = n.marks.map((m) => `${m.type.name}=${m.attrs.val}`).sort().join(' '); }); return out; };
  return [marks('Red'), marks('plain')];
});
check(pasted[0] === 'color=C00000 font=Georgia size=32' && pasted[1] === '', `pasted colour, size and font kept (${pasted.join(' | ')})`);

// Fit text (mkfittext.py): each fitted run, or group of runs, as wide as it says, the text after it starting there, as in
// Word's PDF (96, 192, 96, 72 and 144px on from the paragraph's left).
await open('fittext.docx');
const fits = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].slice(0, 5).map((p) => {
  const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  for (let t = w.nextNode(); t; t = w.nextNode()) {
    const i = t.data.indexOf('|');
    if (i < 0) continue;
    const r = document.createRange(); r.setStart(t, i); r.setEnd(t, i + 1);
    return Math.round(r.getBoundingClientRect().left - p.getBoundingClientRect().left);
  }
  return null;
}));
check(fits.join() === '96,192,96,72,144', `fitted text as wide as it says, as in Word (${fits.join()})`);

// Typing through an input method (Chinese, Japanese, Korean: composed, then committed) in each part of a document: the
// body, a text box, the header, the footer, a footnote and an endnote (find-all.docx, each reached with Find).
const typeIme = async (text) => {
  await cdp.send('Input.imeSetComposition', { text: text.slice(0, 1), selectionStart: 1, selectionEnd: 1 });
  await cdp.send('Input.imeSetComposition', { text, selectionStart: text.length, selectionEnd: text.length });
  await cdp.send('Input.insertText', { text });
  await page.waitForTimeout(150);
};
await open('find-all.docx');
await page.click('.odx-content p');
await page.keyboard.press('Control+Home');
await page.keyboard.press('Control+f');
await page.fill('.odx-findbar input[aria-label="Find"]', 'apple');
const imeParts = [];
for (let i = 0; i < 7; i++) {
  await page.press('.odx-findbar input[aria-label="Find"]', 'Enter');
  await page.waitForTimeout(250);
  if (i === 1) continue; // (the body's second apple)
  const v = await page.evaluate(() => { const v = window.odx.view.odx.hfView || window.odx.view; v.focus(); const { to } = v.state.selection; v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.near(v.state.doc.resolve(to)))); return true; });
  void v;
  await typeIme(`漢字${i}`);
  imeParts.push(await page.evaluate((i) => { const v = window.odx.view.odx.hfView || window.odx.view; return v.state.doc.textContent.includes(`漢字${i}`); }, i));
}
await page.keyboard.press('Escape');
const imeSaved = unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))));
const imeIn = ['word/document.xml', 'word/header1.xml', 'word/footer1.xml', 'word/footnotes.xml', 'word/endnotes.xml'].map((f) => (strFromU8(imeSaved[f]).match(/漢字\d/g) || []).length);
check(imeParts.every(Boolean) && imeIn.join() === '3,1,1,1,1', `input-method typing in the body, a text box, the header, the footer and notes, saved (${imeParts.join()}; ${imeIn.join()}: the text box's twice, its shape and fallback copy)`);

// Track Changes edge cases: a word deleted letter by letter, then typed over, reads in order (the new text before the
// deleted, where Word's caret is once it has gone back over it); undoing tracked typing
// takes its change away; Reject All in a table cell leaves no insertion behind.
await open('rich.docx');
await page.evaluate(() => window.odx.view.focus());
await page.click('.odx-content p >> nth=0');
await page.keyboard.press('End');
await page.keyboard.type(' alpha');
await page.evaluate(() => { const t = document.querySelector('.odx-tracksel'); t.value = 'on'; t.dispatchEvent(new Event('change')); });
await page.evaluate(() => window.odx.view.focus());
for (let i = 0; i < 5; i++) await page.keyboard.press('Backspace');
await page.keyboard.type('beta');
await page.waitForTimeout(150);
const order = await page.evaluate(() => { const p = window.odx.view.state.doc.firstChild; let out = ''; p.forEach((n) => { if (n.isText) out += n.marks.some((m) => m.type.name === 'del') ? `[-${n.text}]` : n.marks.some((m) => m.type.name === 'ins') ? `[+${n.text}]` : ''; }); return out; });
check(/\[\+beta\]\[-alpha\]$/.test(order), `a word deleted letter by letter and typed over reads in order (${order})`);
const insBefore = await page.evaluate(() => document.querySelectorAll('.odx-content ins.odx-ins').length);
await page.waitForTimeout(700); // (a new undo step, not the typing before)
await page.keyboard.type('X');
await page.keyboard.press('Control+z');
await page.waitForTimeout(150);
check(await page.evaluate((n) => document.querySelectorAll('.odx-content ins.odx-ins').length === n && !window.odx.view.state.doc.textContent.includes('betaX'), insBefore), 'undoing tracked typing takes its insertion away');
const cellOk = await page.evaluate(async () => {
  const v = window.odx.view;
  let at = null;
  v.state.doc.descendants((n, pos) => { if (at == null && n.type.name === 'table_cell') at = pos + 2; return at == null; });
  if (at == null) return 'no table';
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.near(v.state.doc.resolve(at))));
  v.focus();
  return 'ok';
});
await page.keyboard.type('cell text');
await page.evaluate(() => { const b = [...document.querySelectorAll('.odx-tracksel')][0]; void b; });
await page.evaluate(() => window.odx.agent().resolveChanges({ accept: false }));
const strayIns = /<w:ins\b/.test(strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']));
check(cellOk === 'ok' && !strayIns && !(await bodyText()).includes('cell text'), `Reject All with a change in a table cell leaves no insertion behind (${cellOk}, ${strayIns})`);

// Inside a protected document's editable part: Enter, Tab and paste all work there.
await open('protect-readonly.docx');
// (The caret after "Everyone may edit", inside the part.)
await page.evaluate(() => {
  const v = window.odx.view;
  let at = null;
  v.state.doc.descendants((n, pos) => { if (at == null && n.isText && n.text.includes('may edit')) at = pos + n.text.indexOf('may edit') + 8; return at == null; });
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.near(v.state.doc.resolve(at))));
  v.focus();
});
await page.keyboard.press('Tab');
await page.keyboard.type('T');
await page.keyboard.press('Enter');
await page.keyboard.type('N');
await page.evaluate(() => { const dt = new DataTransfer(); dt.setData('text/plain', 'P'); window.odx.view.dom.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); });
await page.waitForTimeout(150);
const permText = await page.evaluate(() => { const d = window.odx.view.state.doc; return d.textBetween(0, d.content.size, '\n', (n) => (n.type.name === 'tab' ? '\t' : '')); });
check(/edit\tT\nNP this/.test(permText), `Tab, Enter and paste inside an editable part of a protected document (${JSON.stringify(permText.slice(0, 80))})`);

// Zoom (View > Zoom, zoom.js): only what's shown changes. The pages break the same at 150%, a click puts the caret where
// it lands, a double-click on the header edits it there, Ctrl+wheel zooms, and a PDF export is the same as at 100%.
await open('fields.docx');
const zoomSnap = () => page.evaluate(() => { const v = window.odx.view; const sh = v.dom.parentElement; const z = parseFloat(sh.style.zoom) || 1; const top = sh.getBoundingClientRect().top; return { pages: v.odx.pages.count, tops: [...v.dom.querySelectorAll(':scope > p')].filter((_, i) => i % 8 === 0).map((p) => Math.round((p.getBoundingClientRect().top - top) / z)) }; });
const z100 = await zoomSnap();
const pdf100 = await page.evaluate(async () => (await window.odx.exportPdf()).length);
await page.evaluate(() => window.odx.setZoom(1.5));
await page.waitForTimeout(1500);
const z150 = await zoomSnap();
check(z150.pages === z100.pages && z150.tops.every((t, i) => Math.abs(t - z100.tops[i]) <= 1), `at 150% the pages break as at 100% (${z100.pages} pages; ${z150.tops.join(',')})`);
await page.click('[role=tab]:text-is("View")');
check(await page.evaluate(() => document.querySelector('.odx-zoomsel').value === '1.5'), 'View > Zoom shows 150%');
await page.click('.odx-content p:has-text("Front matter paragraph 2.")', { position: { x: 30, y: 8 } });
await page.keyboard.type('@@');
check(await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].find((p) => p.textContent.includes('@@'))?.textContent.replace('@@', '').startsWith('Front matter paragraph 2.')), 'at 150% a click puts the caret where it lands');
await page.keyboard.press('Control+z');
const hdr = await page.evaluate(() => { const r = document.querySelector('.odx-sheet .odx-hf')?.getBoundingClientRect(); return r && { x: r.left + r.width / 2, y: r.top + r.height / 2 }; });
await page.mouse.dblclick(hdr.x, hdr.y);
await page.waitForTimeout(400);
const hfAt = await page.evaluate(() => { const e = document.querySelector('.odx-hfedit-header'); const h = document.querySelector('.odx-sheet .odx-hf'); return e && h ? Math.round(Math.abs(e.getBoundingClientRect().top - h.getBoundingClientRect().top)) : null; });
check(hfAt != null && hfAt < 30, `at 150% a double-click on the header edits it in place (${hfAt}px off)`);
await page.keyboard.press('Escape');
await page.mouse.move(640, 600);
await page.keyboard.down('Control');
await page.mouse.wheel(0, -100);
await page.keyboard.up('Control');
await page.waitForTimeout(300);
check(await page.evaluate(() => Math.abs(window.odx.zoom - 1.6) < 1e-6), `Ctrl+wheel zooms in 10% (${await page.evaluate(() => window.odx.zoom)})`);
const pdf160 = await page.evaluate(async () => (await window.odx.exportPdf()).length);
check(Math.abs(pdf160 - pdf100) < pdf100 * 0.01 && await page.evaluate(() => window.odx.view.dom.parentElement.style.zoom === '1.6'), `a PDF export at 160% is the 100% one (${pdf100} and ${pdf160} bytes), the zoom kept on screen`);
await page.evaluate(() => window.odx.setZoom('pageWidth'));
await page.waitForTimeout(300);
const fit = await page.evaluate(() => { const sh = window.odx.view.dom.parentElement; const c = sh.parentElement; const cs = getComputedStyle(c); return Math.round(c.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight) - sh.getBoundingClientRect().width); });
check(Math.abs(fit) <= 1, `Page Width fits the page to the window (${fit}px left over)`);
await page.evaluate(() => window.odx.setZoom(1));
await page.click('[role=tab]:text-is("Home")');

await open('strict.docx');
const strictText = await page.evaluate(() => document.querySelector('.odx-content td')?.textContent);
check(strictText === 'Half the page wide', 'a Strict document opens');

check(!errors.length, `no page errors ${errors.join(' | ')}`);
await browser.close();
server.close();
console.log(failed ? `features: ${failed} failed` : 'features: all passed');
process.exit(failed ? 1 : 0);
