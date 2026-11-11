// Types for the one-call setup (mountDocxEditor) and the session; the rest of the API is exported untyped.

export type DocxSource = Uint8Array | ArrayBuffer | Blob | File | Response;

export interface OpenTimings { read: number; parse: number; editor: number; layout: number; total: number }

export interface FieldInfo { tag: string | null; alias: string | null; kind: string; value: string; checked?: boolean; items?: { text: string; value: string }[]; block?: boolean }

export interface MountOptions {
  onChange?: (view: unknown) => void;
  /** Replaces the default download when the user saves (e.g. upload to your server). */
  onSave?: (bytes: Uint8Array, fileName: string) => void;
  onOpened?: (timings: OpenTimings, fileName: string) => void;
  onError?: (error: unknown) => void;
  fileName?: string;
  /** Names the tracked changes and comments made here. */
  author?: string;
  readOnly?: boolean;
  /** Turn tabs or groups off by name, e.g. { review: false, comments: false }. */
  features?: Record<string, boolean>;
  toolbar?: boolean;
  /** Fill-in mode: only content controls (form fields) can be changed. */
  forms?: boolean;
  /** The PDF exporter (exportPdf from 'open-docx-editor/pdf'): File > Export then downloads a PDF of the pages. */
  pdf?: (editor: DocxEditorApi, options?: { title?: string; author?: string }) => Promise<Uint8Array>;
  /** Co-editing: a room from collaboration() in 'open-docx-editor/collab'. Mount after the provider has synced. */
  collab?: unknown;
  /** Replaces the PDF download (with `pdf`). */
  onSavePdf?: (bytes: Uint8Array, fileName: string) => void;
  /** After each open, the fonts the document asks for that this computer lacks (only when there are any). Lines set in a
   *  substitute may break differently from Word's; import 'open-docx-editor/fonts.css' for metric-compatible ones. */
  onMissingFonts?: (fonts: MissingFont[]) => void;
  /** Keep unsaved changes in this browser (IndexedDB) and offer them back on opening; off by default. `key` names the
   *  document (its file name if not given), `delayMs` how long after an edit it's kept (3000). */
  drafts?: boolean | { key?: string; delayMs?: number };
  /** In place of the built-in notice: a draft of unsaved changes was found on opening. */
  onDraft?: (draft: Draft, actions: { restore(): Promise<void>; discard(): Promise<void> }) => void;
  /** This user's names besides `author` (an e-mail address, DOMAIN\name), as Word names the editors of a protected
   *  document's editable parts. `features.protection: false` ignores Restrict Editing, Mark as Final and Always open
   *  read-only (for a host that enforces its own rights). */
  editors?: string[];
}

/** A document's Restrict Editing (Word's Review > Restrict Editing, Mark as Final, Always open read-only). */
export interface Protection {
  edit: 'readOnly' | 'comments' | 'trackedChanges' | 'forms' | null;
  enforced: boolean;
  /** Limited to the styles not locked, with no direct formatting. */
  formatting: boolean;
  /** A password is needed to stop it. */
  password: boolean;
  readOnlyRecommended: boolean;
  passwordToModify: boolean;
  final: boolean;
  /** Held read-only (marked final, read-only recommended, a password to modify) until editAnyway(). */
  readOnlyNow: boolean;
}

/** Unsaved changes kept in the browser (see MountOptions.drafts). */
export interface Draft { key: string; name: string; savedAt: number; bytes: Uint8Array; }

/** A font the document asks for that isn't available, and the font drawn instead (null: the browser's generic one). */
/** `standIn`: drawn in one of the editor's own stand-ins, another font sized to this one's widths (substitute null). */
export interface MissingFont { font: string; substitute: string | null; standIn: boolean; }

export interface DocxEditorApi {
  readonly view: unknown;
  readonly fileName: string;
  readonly readOnly: boolean;
  readonly forms: boolean;
  readonly formsProtected: boolean;
  /** The open document's protection, or null with none open. */
  readonly protection: Protection | null;
  /** Start enforcing protection (saved with the document; Word enforces it too). */
  protect(options: { edit?: Protection['edit']; formatting?: boolean; password?: string }): Promise<void>;
  /** The zoom the pages are shown at (1 is 100%). The layout, printing and PDF export are always at 100%. */
  readonly zoom: number;
  /** Show the pages at a zoom (0.1 to 5), or fitted to the window ('pageWidth', 'onePage') as it resizes. */
  setZoom(zoom: number | 'pageWidth' | 'onePage'): void;
  /** Stop protection: false when `password` isn't the document's. */
  unprotect(password?: string): Promise<boolean>;
  /** Edit Anyway (a document marked final or to open read-only; with a password to modify, it): whether it can be edited now. */
  editAnyway(password?: string): Promise<boolean>;
  /** True once the document has been edited since it was opened or last saved with markSaved(). */
  dirty: boolean;
  timings: OpenTimings | null;
  /** The fonts the open document asks for that this computer lacks. */
  missingFonts(): Promise<MissingFont[]>;
  /** `reopen`: the same document again (after a page setup change, or a restored draft): it keeps its unsaved state. */
  open(src: DocxSource, fileName?: string, options?: { share?: boolean; reopen?: boolean }): Promise<OpenTimings>;
  newDocument(options?: { share?: boolean }): void;
  /** Word's Update Fields for the body (page numbers, PAGEREF, REF, STYLEREF, the table of contents): how many changed. */
  updateFields(): number;
  /** An agent's view of the open document: its edits are tracked changes under `author` ('AI') in this editor. */
  agent(options?: { author?: string }): import('./agent').AgentDocument;
  save(): Uint8Array;
  download(fileName?: string): void;
  /** The pages as a PDF (needs the `pdf` option). */
  exportPdf(): Promise<Uint8Array>;
  downloadPdf(fileName?: string): Promise<Uint8Array>;
  /** Clears `dirty` and, with `drafts`, the draft kept in the browser. */
  markSaved(): void;
  /** The draft of unsaved changes kept for the open document, or null (with `drafts`). */
  draft(): Promise<Draft | null>;
  setReadOnly(on: boolean): void;
  setForms(on: boolean): void;
  setFeatures(features: Record<string, boolean>): void;
  getFields(): FieldInfo[];
  getFieldValues(): Record<string, string | boolean>;
  setFieldValues(values: Record<string, string | boolean>): unknown;
  destroy(): void;
}

export function mountDocxEditor(container: HTMLElement, options?: MountOptions): DocxEditorApi;

export class DocxSession {
  constructor(bytes: Uint8Array | ArrayBuffer, options?: Record<string, unknown>);
  static blank(options?: Record<string, unknown>): DocxSession;
  doc: unknown;
  save(doc?: unknown): Uint8Array;
}

export function blankDocxBytes(): Uint8Array;
export function readFieldValues(bytes: Uint8Array): Record<string, string | boolean>;
export function fillFields(bytes: Uint8Array, values: Record<string, string | boolean>): { bytes: Uint8Array; missing: string[] };

export const schema: unknown;
export const commands: Record<string, unknown>;
export const track: Record<string, unknown>;
export function createEditor(...args: unknown[]): unknown;
export function saveEditor(...args: unknown[]): unknown;
export function createToolbar(...args: unknown[]): unknown;
export function setXmlImpl(impl: { DOMParser: unknown; XMLSerializer: unknown }): void;
export function listFields(...args: unknown[]): unknown;
export function fieldValues(...args: unknown[]): unknown;
export function setFieldValues(...args: unknown[]): unknown;
export function setForms(...args: unknown[]): unknown;
