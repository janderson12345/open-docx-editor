// React wrapper test page: test/wrappers.mjs drives it through window.__wr.
import { createElement, createRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { DocxEditor } from '../../src/wrappers/react.js';

const log = [];
const apiRef = createRef();
let set = null;
function App() {
  const [state, setState] = useState({ src: null, readOnly: false, shown: true });
  set = setState;
  return state.shown ? createElement(DocxEditor, { ref: apiRef, src: state.src, readOnly: state.readOnly, fileName: 'form.docx', onOpened: (t, n) => log.push(`opened ${n}`), onReady: () => log.push('ready') }) : null;
}
const root = createRoot(document.getElementById('app'));
root.render(createElement(App));
window.__wr = {
  log,
  open: (bytes) => set((s) => ({ ...s, src: bytes })),
  readOnly: (on) => set((s) => ({ ...s, readOnly: on })),
  save: () => apiRef.current.save(),
  unmount: () => set((s) => ({ ...s, shown: false })),
};
