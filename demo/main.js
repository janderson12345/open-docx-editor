import { mountDocxEditor } from '../src/index.js';
import { exportPdf } from '../src/pdf.js';

const app = document.getElementById('app');
const editor = mountDocxEditor(app, { pdf: exportPdf });
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
