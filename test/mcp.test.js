// The MCP server (bin/odx-mcp.mjs) driven as an AI assistant drives it: JSON-RPC over stdin/stdout.
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { unzipSync, strFromU8 } from 'fflate';

function server(root) {
  const child = spawn(process.execPath, [fileURLToPath(new URL('../bin/odx-mcp.mjs', import.meta.url)), '--root', root, '--author', 'Claude'], { stdio: ['pipe', 'pipe', 'inherit'] });
  let buf = '';
  const waiting = new Map();
  child.stdout.on('data', (d) => {
    buf += d;
    let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const msg = JSON.parse(buf.slice(0, i));
      buf = buf.slice(i + 1);
      waiting.get(msg.id)?.(msg);
    }
  });
  let n = 0;
  const call = (method, params) => new Promise((resolve) => { const id = ++n; waiting.set(id, resolve); child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n'); });
  const tool = async (name, args) => { const r = await call('tools/call', { name, arguments: args }); return { text: r.result.content[0].text, error: !!r.result.isError }; };
  return { call, tool, close: () => child.kill() };
}

test('an assistant opens, reads, edits (tracked), comments and saves a copy; nothing outside the folder', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'odx-mcp-'));
  fs.copyFileSync(new URL('./fixtures/fields.docx', import.meta.url), path.join(root, 'report.docx'));
  const s = server(root);
  try {
    const init = await s.call('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '1' } });
    assert.equal(init.result.serverInfo.name, 'open-docx-editor');
    const tools = (await s.call('tools/list', {})).result.tools.map((t) => t.name);
    for (const t of ['read_document', 'get_outline', 'search', 'replace_text', 'set_paragraph_text', 'insert_paragraph', 'delete_paragraph', 'add_comment', 'list_changes', 'save_document']) assert.ok(tools.includes(t), t);
    assert.equal((await s.tool('list_documents', {})).text, 'report.docx');
    const md = (await s.tool('read_document', { path: 'report.docx' })).text;
    assert.match(md, /# 2\. Methods of measurement/);
    const outline = JSON.parse((await s.tool('get_outline', { path: 'report.docx' })).text);
    assert.equal((await s.tool('replace_text', { path: 'report.docx', find: 'Results', replace: 'Findings', paragraph: outline[2].id })).text, '1 replaced.');
    assert.match((await s.tool('add_comment', { path: 'report.docx', paragraph: outline[1].id, text: 'Say which instruments.' })).text, /^Comment \d+ added\.$/);
    const changes = JSON.parse((await s.tool('list_changes', { path: 'report.docx' })).text);
    assert.ok(changes.length === 2 && changes.every((c) => c.author === 'Claude'));
    // Saving: a new file by default; over the original or an existing file only when asked.
    assert.ok((await s.tool('save_document', { path: 'report.docx' })).error);
    assert.equal((await s.tool('save_document', { path: 'report.docx', save_as: 'report (edited).docx' })).text, 'Saved report (edited).docx.');
    assert.ok((await s.tool('save_document', { path: 'report.docx', save_as: 'report (edited).docx' })).error);
    const xml = strFromU8(unzipSync(fs.readFileSync(path.join(root, 'report (edited).docx')))['word/document.xml']);
    assert.match(xml, /<w:ins [^>]*w:author="Claude"[^>]*>.*Findings/);
    // Outside the folder: refused, for reading and for saving.
    const outside = await s.tool('read_document', { path: '../escape.docx' });
    assert.ok(outside.error && /outside the folder|no such file/i.test(outside.text));
    assert.ok((await s.tool('save_document', { path: 'report.docx', save_as: '../escape.docx' })).error);
    assert.ok(!fs.existsSync(path.join(root, '..', 'escape.docx')));
    assert.ok((await s.tool('read_document', { path: 'notes.txt' })).error);
    assert.equal((await s.call('tools/call', { name: 'no_such_tool', arguments: {} })).error.code, -32602);
  } finally {
    s.close();
    fs.rmSync(root, { recursive: true, force: true });
  }
});
