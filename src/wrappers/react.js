// React wrapper: <DocxEditor src={file} onChange={...} ref={apiRef} />. Plain JavaScript (no JSX), React 17+.
// The editor is mounted once into its own div and kept: props that change after mounting are applied through the
// editor's API (readOnly, features, forms), a new `src` opens that document, and unmounting destroys the editor.
import { createElement, forwardRef, useEffect, useImperativeHandle, useRef } from 'react';
import { mountDocxEditor } from '../index.js';

/**
 * Props: `src` (bytes, File, Blob or Response to open; a new value opens it), `fileName`, `author`, `readOnly`,
 * `features`, `toolbar`, `forms`, and the callbacks `onChange(view)`, `onSave(bytes, fileName)`,
 * `onOpened(timings, fileName)`, `onError(error)`, `onReady(api)`; `className` and `style` for the container.
 * The ref is the editor's API (mountDocxEditor's): `save()`, `download()`, `getFieldValues()`, `view`...
 */
export const DocxEditor = forwardRef(function DocxEditor(props, ref) {
  const { src, fileName, author, readOnly = false, features, toolbar = true, forms = false, className, style } = props;
  const host = useRef(null);
  const api = useRef(null);
  // The latest callbacks, so the editor (created once) always calls the current ones.
  const cb = useRef(props);
  cb.current = props;

  useEffect(() => {
    const editor = mountDocxEditor(host.current, {
      fileName, author, readOnly, features, toolbar, forms,
      onChange: (view) => cb.current.onChange?.(view),
      onSave: cb.current.onSave ? (bytes, name) => cb.current.onSave(bytes, name) : undefined,
      onOpened: (timings, name) => cb.current.onOpened?.(timings, name),
      onError: (e) => (cb.current.onError ? cb.current.onError(e) : console.error(e)),
    });
    api.current = editor;
    if (cb.current.src == null) editor.newDocument(); // (no document given: a blank one, as Word's New)
    cb.current.onReady?.(editor);
    return () => { editor.destroy(); api.current = null; };
  }, []); // (created once: the effects below follow the props)

  useEffect(() => {
    if (src == null || !api.current) return;
    api.current.open(src, fileName).catch((e) => (cb.current.onError ? cb.current.onError(e) : console.error(e)));
  }, [src]); // (dependencies deliberately partial)
  useEffect(() => { api.current?.setReadOnly(readOnly); }, [readOnly]);
  useEffect(() => { api.current?.setForms(forms); }, [forms]);
  useEffect(() => { if (features) api.current?.setFeatures(features); }, [JSON.stringify(features || {})]); // (dependencies deliberately partial)

  // The ref is set before effects run, so it forwards to the editor once there is one.
  useImperativeHandle(ref, () => new Proxy({}, { get: (_, k) => api.current?.[k] }), []);
  return createElement('div', { ref: host, className, style });
});
