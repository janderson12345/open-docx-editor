#!/usr/bin/env node
// An MCP server (Model Context Protocol, over stdin/stdout) that lets an AI assistant read and edit .docx files in one
// folder: open, read as Markdown, outline, search, replace text, rewrite / insert / delete paragraphs, comment, list
// comments and tracked changes, fill form fields, save. Every edit is a tracked change under the assistant's name
// unless the call says otherwise, so a person reviews them in Word. It reads and writes only inside the folder it's
// given (--root, else the current folder), never overwrites a file unless asked, and makes no network requests.
//   usage: open-docx-editor-mcp [--root <folder>] [--author <name>]
import fs from 'node:fs';
import path from 'node:path';
import readline from 'node:readline';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../src/ooxml/xml.js';
import { openDocument } from '../src/agent.js';

setXmlImpl({ DOMParser, XMLSerializer });
const arg = (name, fallback) => { const i = process.argv.indexOf(name); return i > 0 && process.argv[i + 1] ? process.argv[i + 1] : fallback; };
const ROOT = fs.realpathSync(path.resolve(arg('--root', process.cwd())));
const AUTHOR = arg('--author', 'AI');
const VERSION = JSON.parse(fs.readFileSync(new URL('../package.json', import.meta.url), 'utf8')).version;
const open = new Map(); // full path -> AgentDocument

/** A path inside the folder (relative to it, or absolute within it); anything outside is refused. */
function inside(p, { mustExist = true } = {}) {
  if (typeof p !== 'string' || !p) throw new Error('A file path is needed.');
  const full = path.resolve(ROOT, p);
  const real = mustExist ? fs.realpathSync(full) : path.join(fs.realpathSync(path.dirname(full)), path.basename(full));
  if (real !== ROOT && !real.startsWith(ROOT + path.sep)) throw new Error(`${p} is outside the folder this server works in.`);
  if (!/\.docx$/i.test(real)) throw new Error(`${p} isn't a .docx file.`);
  return real;
}
function doc(p) {
  const full = inside(p);
  if (!open.has(full)) open.set(full, openDocument(new Uint8Array(fs.readFileSync(full)), { author: AUTHOR }));
  return open.get(full);
}
const json = (v) => JSON.stringify(v, null, 1);
const S = (description, extra = {}) => ({ type: 'string', description, ...extra });
const B = (description) => ({ type: 'boolean', description });
const PATH = S('The .docx file, relative to the folder the server works in.');
const PARA = S('A paragraph id from read_document, get_outline or search, e.g. "p12".');
const TRACKED = B('Make the edit a tracked change under the assistant\'s name (default true). Use false only when the user asks for untracked edits.');

const TOOLS = [
  { name: 'list_documents', description: 'The .docx files in the folder the server works in (and its subfolders).', inputSchema: { type: 'object', properties: {} },
    run: () => {
      const out = [];
      const walk = (d) => { for (const e of fs.readdirSync(d, { withFileTypes: true })) { if (e.name.startsWith('.') || e.name === 'node_modules') continue; const f = path.join(d, e.name); if (e.isDirectory()) { if (out.length < 500) walk(f); } else if (/\.docx$/i.test(e.name) && !e.name.startsWith('~$')) out.push(path.relative(ROOT, f)); } };
      walk(ROOT);
      return out.join('\n') || 'No .docx files here.';
    } },
  { name: 'read_document', description: 'A document as Markdown: headings, lists, tables, each paragraph preceded by its id (<!-- p12 -->). Long documents: read a part with from/to ids.', inputSchema: { type: 'object', properties: { path: PATH, from: PARA, to: PARA }, required: ['path'] },
    run: (a) => doc(a.path).markdown({ from: a.from, to: a.to }) },
  { name: 'get_outline', description: 'The headings, with their levels and paragraph ids.', inputSchema: { type: 'object', properties: { path: PATH }, required: ['path'] },
    run: (a) => json(doc(a.path).outline()) },
  { name: 'get_paragraphs', description: 'Every paragraph with its id, style, heading level, list number, table cell and text, as JSON.', inputSchema: { type: 'object', properties: { path: PATH }, required: ['path'] },
    run: (a) => json(doc(a.path).paragraphs()) },
  { name: 'search', description: 'Where text appears: each match with its paragraph id and the paragraph\'s text.', inputSchema: { type: 'object', properties: { path: PATH, query: S('The text to find (or a regular expression with regex).'), regex: B('The query is a regular expression.'), match_case: B('Match capitals (default false).') }, required: ['path', 'query'] },
    run: (a) => json(doc(a.path).search(a.query, { regex: !!a.regex, matchCase: !!a.match_case })) },
  { name: 'replace_text', description: 'Replaces text: the first match (in one paragraph if given), or every one with all. With regex, $1, $2 in the replacement are the groups.', inputSchema: { type: 'object', properties: { path: PATH, find: S('The text to replace.'), replace: S('The new text.'), paragraph: PARA, all: B('Every match, not only the first.'), regex: B('find is a regular expression.'), tracked: TRACKED }, required: ['path', 'find', 'replace'] },
    run: (a) => `${doc(a.path).replace(a.find, a.replace, { id: a.paragraph, all: !!a.all, regex: !!a.regex, tracked: a.tracked !== false })} replaced.` },
  { name: 'set_paragraph_text', description: 'Rewrites one paragraph\'s text, keeping its style; only what differs is changed.', inputSchema: { type: 'object', properties: { path: PATH, paragraph: PARA, text: S('The paragraph\'s new text.'), tracked: TRACKED }, required: ['path', 'paragraph', 'text'] },
    run: (a) => (doc(a.path).setParagraphText(a.paragraph, a.text, { tracked: a.tracked !== false }) ? 'Rewritten.' : 'Unchanged: it already reads that way.') },
  { name: 'insert_paragraph', description: 'Inserts a paragraph after another (or at the start with no after). Returns its id.', inputSchema: { type: 'object', properties: { path: PATH, after: PARA, text: S('Its text.'), style: S('A style name, e.g. "Heading 2" or "Normal" (default: the paragraph before\'s).'), tracked: TRACKED }, required: ['path', 'text'] },
    run: (a) => `Inserted ${doc(a.path).insertParagraph(a.after ?? null, a.text, { style: a.style, tracked: a.tracked !== false })}.` },
  { name: 'delete_paragraph', description: 'Deletes a paragraph.', inputSchema: { type: 'object', properties: { path: PATH, paragraph: PARA, tracked: TRACKED }, required: ['path', 'paragraph'] },
    run: (a) => { doc(a.path).deleteParagraph(a.paragraph, { tracked: a.tracked !== false }); return 'Deleted.'; } },
  { name: 'add_comment', description: 'Adds a comment on a paragraph, or on words within it (quote). Returns its id.', inputSchema: { type: 'object', properties: { path: PATH, paragraph: PARA, text: S('The comment.'), quote: S('The words in the paragraph to comment on (optional).') }, required: ['path', 'paragraph', 'text'] },
    run: (a) => `Comment ${doc(a.path).addComment(a.paragraph, a.text, { quote: a.quote })} added.` },
  { name: 'reply_to_comment', description: 'Replies to a comment.', inputSchema: { type: 'object', properties: { path: PATH, comment: S('The comment\'s id, from list_comments.'), text: S('The reply.') }, required: ['path', 'comment', 'text'] },
    run: (a) => `Reply ${doc(a.path).reply(a.comment, a.text)} added.` },
  { name: 'list_comments', description: 'The comments: author, date, text, the words they\'re on, replies.', inputSchema: { type: 'object', properties: { path: PATH }, required: ['path'] },
    run: (a) => json(doc(a.path).comments()) },
  { name: 'list_changes', description: 'The tracked changes: insertions and deletions, who made them, and the paragraph.', inputSchema: { type: 'object', properties: { path: PATH }, required: ['path'] },
    run: (a) => json(doc(a.path).changes()) },
  { name: 'get_fields', description: 'The form fields (content controls): their names and values.', inputSchema: { type: 'object', properties: { path: PATH }, required: ['path'] },
    run: (a) => json(doc(a.path).fields({ details: true })) },
  { name: 'set_fields', description: 'Fills form fields by name: {"CustomerName": "Acme", "Agree": true}.', inputSchema: { type: 'object', properties: { path: PATH, values: { type: 'object', description: 'Field name -> value (text, or true/false for a check box).' } }, required: ['path', 'values'] },
    run: (a) => { const missing = doc(a.path).setFields(a.values || {}); return missing.length ? `Filled; no field named: ${missing.join(', ')}.` : 'Filled.'; } },
  { name: 'save_document', description: 'Saves the edits: to save_as (a new file in the folder), or over the original only with overwrite.', inputSchema: { type: 'object', properties: { path: PATH, save_as: S('A new file name for the edited copy, e.g. "contract (edited).docx".'), overwrite: B('Save over the original (only when the user asked for that).') }, required: ['path'] },
    run: (a) => {
      const d = doc(a.path);
      let target;
      if (a.save_as) {
        target = inside(a.save_as, { mustExist: false });
        if (fs.existsSync(target) && !a.overwrite) throw new Error(`${a.save_as} already exists; give another name, or overwrite.`);
      } else if (a.overwrite) target = inside(a.path);
      else throw new Error('Give save_as (a new file name), or overwrite: true to save over the original.');
      fs.writeFileSync(target, d.save());
      return `Saved ${path.relative(ROOT, target)}.`;
    } },
];

// --- JSON-RPC over stdin/stdout, one message per line (MCP's stdio transport) ---------------------------------------
const send = (msg) => process.stdout.write(JSON.stringify({ jsonrpc: '2.0', ...msg }) + '\n');
async function handle(msg) {
  const { id, method, params } = msg;
  if (id == null) return; // (notifications: initialized, cancelled... nothing to answer)
  try {
    if (method === 'initialize') {
      return send({ id, result: { protocolVersion: params?.protocolVersion || '2025-06-18', capabilities: { tools: {} }, serverInfo: { name: 'open-docx-editor', version: VERSION },
        instructions: 'Edits Word documents in one folder. Read with read_document (paragraph ids like p12), edit by id; edits are tracked changes for the user to review. Save to a new file with save_document save_as unless the user asks to overwrite.' } });
    }
    if (method === 'ping') return send({ id, result: {} });
    if (method === 'tools/list') return send({ id, result: { tools: TOOLS.map(({ run, ...t }) => t) } });
    if (method === 'tools/call') {
      const tool = TOOLS.find((t) => t.name === params?.name);
      if (!tool) return send({ id, error: { code: -32602, message: `Unknown tool: ${params?.name}` } });
      try {
        const text = await tool.run(params.arguments || {});
        return send({ id, result: { content: [{ type: 'text', text: String(text) }] } });
      } catch (e) {
        return send({ id, result: { content: [{ type: 'text', text: e.message }], isError: true } });
      }
    }
    send({ id, error: { code: -32601, message: `Method not found: ${method}` } });
  } catch (e) {
    send({ id, error: { code: -32603, message: e.message } });
  }
}
readline.createInterface({ input: process.stdin }).on('line', (line) => {
  if (!line.trim()) return;
  let msg;
  try { msg = JSON.parse(line); } catch { return send({ id: null, error: { code: -32700, message: 'Parse error' } }); }
  handle(msg);
});
