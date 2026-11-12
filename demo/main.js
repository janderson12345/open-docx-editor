import { mountDocxEditor } from '../src/index.js';
import { exportPdf } from '../src/pdf.js';

const app = document.getElementById('app');
// A document's fonts this computer lacks: said in a line above the editor, as lines may break differently from Word's.
const note = document.getElementById('fonts-note');
const onMissingFonts = (fonts) => {
  const list = fonts.map((f) => (f.substitute ? `${f.font} (shown in ${f.substitute})` : f.standIn ? `${f.font} (shown in a font sized to it)` : f.font)).join(', ');
  note.textContent = `This computer doesn't have ${list}, so some lines may break differently from Word.`;
  note.hidden = false;
};
// (drafts: unsaved changes kept in this browser and offered back when the same file is opened again)
// (window.odxAi: an AI model for Home > AI Writing, when the page has one: the tests give a stand-in.)
const editor = mountDocxEditor(app, { pdf: exportPdf, onMissingFonts, drafts: true, ai: window.odxAi || null, fonts: { base: './dist/fonts/' }, onOpened: () => { note.hidden = true; } });
window.odx = editor; // handy for debugging and tests

// Drag a .docx onto the page to open it.
app.addEventListener('dragover', (e) => { e.preventDefault(); app.classList.add('drop'); });
app.addEventListener('dragleave', () => app.classList.remove('drop'));
app.addEventListener('drop', (e) => {
  e.preventDefault();
  app.classList.remove('drop');
  const f = e.dataTransfer.files[0];
  if (f) editor.open(f);
});

// A sample document beside the page (the published demo has one) opens at the start; without one, a blank document.
fetch('sample.docx').then((r) => (r.ok ? r.arrayBuffer() : null)).then((b) => b && editor.open(b, 'sample.docx')).catch(() => {});
