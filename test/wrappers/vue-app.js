// Vue wrapper test page: test/wrappers.mjs drives it through window.__wr.
import { createApp, h, reactive, ref } from 'vue';
import { DocxEditor } from '../../src/wrappers/vue.js';

const log = [];
const state = reactive({ src: null, readOnly: false, shown: true });
const editor = ref(null);
createApp({
  render: () => (state.shown ? h(DocxEditor, { ref: editor, src: state.src, readOnly: state.readOnly, fileName: 'form.docx', onOpened: (t, n) => log.push(`opened ${n}`), onReady: () => log.push('ready') }) : null),
}).mount('#app');
window.__wr = {
  log,
  open: (bytes) => { state.src = bytes; },
  readOnly: (on) => { state.readOnly = on; },
  save: () => editor.value.api.save(),
  unmount: () => { state.shown = false; },
};
