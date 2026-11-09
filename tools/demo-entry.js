import { mountDocxEditor } from '../src/index.js';
import { exportPdf } from '../src/pdf.js';
document.documentElement.lang = 'en';
const app = document.getElementById('app');
const status = document.getElementById('status');
const say = (t) => { status.textContent = t; };
// Where does the wait go when opening from the file dialog? Mark when the
// dialog closes (window regains focus) and when the browser delivers the file.
const pick = { closed: 0, delivered: 0 };
window.addEventListener('focus', () => { pick.closed = performance.now(); });
document.addEventListener('change', (e) => { if (e.target.type === 'file') pick.delivered = performance.now(); }, true);
const editor = window.odx = mountDocxEditor(app, {
  onOpened(t, name) {
    const fm = [...document.querySelectorAll('label')].find((l) => l.textContent.trim() === 'Forms mode')?.querySelector('input'); if (fm) fm.checked = editor.forms;
    const ms = (v) => Math.round(v);
    const wait = pick.delivered && pick.closed && pick.delivered >= pick.closed ? `Browser took ${ms(pick.delivered - pick.closed)} ms to hand over the file after the dialog closed. ` : '';
    const opened = performance.now();
    requestAnimationFrame(() => requestAnimationFrame(() => {
      say(`${wait}Opened ${name} in ${ms(t.total)} ms (read ${ms(t.read)} · parse ${ms(t.parse)} · editor ${ms(t.editor)} · pages ${ms(t.layout)}), on screen ${ms(performance.now() - opened)} ms later.`);
      pick.closed = pick.delivered = 0;
    }));
  },
  onError(e) { say(`Couldn't open that file: ${e.message}`); },
  pdf: exportPdf,
  onSavePdf: (bytes, name) => offer(bytes, name, 'application/pdf'),
  onSave: (bytes, name) => offer(bytes, name, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
});
// In the claude.ai viewer a page can't download a file itself: it offers it through the viewer's downloads
// capability (the viewer confirms). Elsewhere, an ordinary browser download.
async function offer(bytes, name, type) {
  const dl = window.claude?.use ? await window.claude.use('downloads').catch(() => null) : null;
  const kb = `${(bytes.length / 1024).toFixed(1)} KB`;
  if (dl) {
    try {
      await dl.save({ filename: name, data: new Blob([bytes]) });
      say(`Saved ${name} (${kb}).`);
    } catch (e) {
      say(e && e.code === 'declined' ? 'Save cancelled.' : `Couldn't save here (${e && e.code || 'unavailable'}).`);
    }
    return;
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([bytes], { type }));
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  say(`Saved ${name} (${kb}).`);
}
const FORM_DOCX = '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';
const sample = Uint8Array.from(atob(window.SAMPLE_DOCX), (c) => c.charCodeAt(0));
editor.open(sample, 'Quarterly Report.docx').then(() => say(`Ready in ${Math.round(performance.now())} ms. Open your own .docx, or drop one on the page.`));
app.addEventListener('dragover', (e) => e.preventDefault());
app.addEventListener('drop', (e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) editor.open(f).catch((err) => say(`Couldn't open that file: ${err.message}`)); });
// Demo of the embedding options: read-only and turning features off.
{
  const head = document.querySelector('header.head');
  const box = document.createElement('div');
  box.className = 'demo-opts';
  box.style.cssText = 'display:flex;gap:12px;font:12px system-ui,sans-serif;align-items:center;flex-wrap:wrap;margin-top:4px';
  const opt = (label, on, fn) => {
    const l = document.createElement('label');
    const c = document.createElement('input');
    c.type = 'checkbox'; c.checked = on; c.onchange = () => fn(c.checked);
    l.append(c, ' ', label);
    box.append(l);
  };
  box.append('Options:');
  opt('Read-only', false, (v) => editor.setReadOnly(v));
  opt('Forms mode', false, (v) => editor.setForms(v));
  opt('Review tab', true, (v) => editor.setFeatures({ review: v }));
  opt('Comments', true, (v) => editor.setFeatures({ comments: v }));
  opt('Track Changes', true, (v) => editor.setFeatures({ tracking: v, markup: v }));
  opt('Insert tab', true, (v) => editor.setFeatures({ insert: v }));
  const btn = (label, fn) => { const b = document.createElement('button'); b.type = 'button'; b.textContent = label; b.onclick = fn; box.append(b); };
  btn('Open sample form', () => editor.open(Uint8Array.from(atob(FORM_DOCX), (c) => c.charCodeAt(0)), 'Sample form.docx'));
  btn('Show field values', () => say('Field values: ' + JSON.stringify(editor.getFieldValues())));
  btn('Fill sample values', () => {
    const left = editor.setFieldValues({ CustomerName: 'Acme Corp', Agree: true, Color: 'G', Due: '11/1/2026', Notes: 'First line\nSecond line' });
    say(left.length ? 'No field named: ' + left.join(', ') : 'Filled the sample form from { CustomerName, Agree, Color, Due, Notes }.');
  });
  head && head.append(box);
}
