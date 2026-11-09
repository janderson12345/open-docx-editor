// Vue 3 wrapper: <DocxEditor :src="file" @change="..." ref="editor" />. Plain JavaScript (a render function, no SFC).
// The editor is mounted once into its own div and kept: props that change after mounting are applied through the
// editor's API (readOnly, features, forms), a new `src` opens that document, and unmounting destroys the editor.
import { defineComponent, h, onBeforeUnmount, onMounted, ref, watch } from 'vue';
import { mountDocxEditor } from '../index.js';

/**
 * Props: `src` (bytes, File, Blob or Response to open; a new value opens it), `fileName`, `author`, `readOnly`,
 * `features`, `toolbar`, `forms`, `onSave` (a function: replaces the download).
 * Events: `change(view)`, `opened(timings, fileName)`, `error(error)`, `ready(api)`.
 * The component's public instance (a template ref) has `api`, the editor's API (mountDocxEditor's): `api.save()`...
 */
export const DocxEditor = defineComponent({
  name: 'DocxEditor',
  props: {
    src: { default: null },
    fileName: { type: String, default: 'document.docx' },
    author: { type: String, default: 'Author' },
    readOnly: { type: Boolean, default: false },
    features: { type: Object, default: null },
    toolbar: { type: Boolean, default: true },
    forms: { type: Boolean, default: false },
    onSave: { type: Function, default: null },
  },
  emits: ['change', 'opened', 'error', 'ready'],
  setup(props, { emit, expose }) {
    const host = ref(null);
    let api = null;
    const fail = (e) => emit('error', e);
    onMounted(() => {
      api = mountDocxEditor(host.value, {
        fileName: props.fileName, author: props.author, readOnly: props.readOnly, features: props.features || {}, toolbar: props.toolbar, forms: props.forms,
        onChange: (view) => emit('change', view),
        onSave: props.onSave ? (bytes, name) => props.onSave(bytes, name) : undefined,
        onOpened: (timings, name) => emit('opened', timings, name),
        onError: fail,
      });
      if (props.src != null) api.open(props.src, props.fileName).catch(fail);
      else api.newDocument(); // (no document given: a blank one, as Word's New)
      emit('ready', api);
    });
    onBeforeUnmount(() => { api?.destroy(); api = null; });
    watch(() => props.src, (src) => { if (api && src != null) api.open(src, props.fileName).catch(fail); });
    watch(() => props.readOnly, (on) => api?.setReadOnly(on));
    watch(() => props.forms, (on) => api?.setForms(on));
    watch(() => props.features, (f) => { if (f) api?.setFeatures(f); }, { deep: true });
    expose({ get api() { return api; } });
    return () => h('div', { ref: host });
  },
});

export default DocxEditor;
