// Reading and filling form fields without an editor (e.g. on a server at check-in).
// In Node, call setXmlImpl(...) with a DOMParser/XMLSerializer first.
import { DocxSession } from './ooxml/docx.js';
import { EditorState } from 'prosemirror-state';
import { fieldValues, listFields, setFieldValues } from './editor/fields.js';

/** { name: value } for every named content control in a .docx. */
export function readFieldValues(bytes) {
  return fieldValues(new DocxSession(bytes).doc);
}

/** A copy of the .docx with fields filled by Tag or Title. Returns { bytes, missing }. */
export function fillFields(bytes, values) {
  const session = new DocxSession(bytes);
  const view = { state: EditorState.create({ doc: session.doc }), dispatch(tr) { this.state = this.state.apply(tr); } };
  const missing = setFieldValues(view, values);
  return { bytes: session.save(view.state.doc), missing };
}

export { listFields };
