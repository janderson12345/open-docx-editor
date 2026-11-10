/** A paragraph as an agent sees it. `id` ('p0', 'p1'...) names it in edits. */
export interface AgentParagraph {
  id: string;
  style: string;
  /** Outline level 1-9 for a heading, else null. */
  heading: number | null;
  /** Its list number or bullet as Word shows it ("2.", "•"), else null. */
  list: string | null;
  table: { table: number; row: number; cell: number } | null;
  /** The text as it reads with its tracked changes accepted. */
  text: string;
}
export interface AgentMatch { id: string; text: string; match: string; start: number; end: number; }
export interface AgentComment { id: string; author: string; date: string; text: string; quote: string; replyTo: string | null; resolved: boolean; }
export interface AgentChange { kind: 'insertion' | 'deletion'; author: string; date: string; text: string; id: string; }
type Tracked = { tracked?: boolean };

/** A .docx to read and edit without a browser. Edits are tracked changes under the agent's name unless `tracked: false`. */
export interface AgentDocument {
  paragraphs(): AgentParagraph[];
  outline(): { id: string; level: number; text: string }[];
  /** Markdown with each paragraph's id in a comment before it; `from`/`to` (ids) for part of it. */
  markdown(options?: { from?: string; to?: string; ids?: boolean }): string;
  search(query: string, options?: { regex?: boolean; matchCase?: boolean; wholeWord?: boolean }): AgentMatch[];
  comments(): AgentComment[];
  changes(): AgentChange[];
  fields(options?: { details?: boolean }): Record<string, string | boolean> | object[];
  replace(query: string, replacement: string, options?: Tracked & { id?: string; all?: boolean; regex?: boolean; matchCase?: boolean; wholeWord?: boolean }): number;
  setParagraphText(id: string, text: string, options?: Tracked): boolean;
  insertParagraph(after: string | null, text: string, options?: Tracked & { style?: string }): string;
  deleteParagraph(id: string, options?: Tracked): boolean;
  addComment(id: string, text: string, options?: { quote?: string; author?: string }): string;
  reply(commentId: string, text: string, options?: { author?: string }): string;
  setFields(values: Record<string, string | boolean>): string[];
  resolveChanges(options?: { accept?: boolean; id?: string }): void;
  save(): Uint8Array;
}

/** Opens a .docx for an agent. In Node, call setXmlImpl(...) first. `author`: the name on its edits and comments ('AI'). */
export function openDocument(bytes: Uint8Array | ArrayBuffer, options?: { author?: string }): AgentDocument;

/** The same on a document open in an editor (an editor view): edits appear in it, tracked under `author`. */
export function editorAgent(view: unknown, options?: { author?: string }): AgentDocument;
