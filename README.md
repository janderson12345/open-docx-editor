# open-docx-editor

Open, edit and save Word documents (`.docx`) in the browser. 100% JavaScript, no server, no
Microsoft Office or SharePoint, MIT licensed.

[![Latest release](https://img.shields.io/github/v/release/janderson12345/open-docx-editor)](https://github.com/janderson12345/open-docx-editor/releases)

It opens and saves real documents without losing what it doesn't understand, and lays them out on pages as Word
does: measured against Word's own PDFs of 151 documents, every one on Word's page count and 99.0% of lines within 1px
of Word's. Text, tables, lists, pictures and shapes (with their text wrapping), text boxes, frames, columns, headers
and footers, footnotes, fields, equations and charts; tracked changes and comments; page setup, Find and Replace
through every part of the document, printing and PDF export. See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the design and roadmap.

**Try it:** [live demo](https://janderson12345.github.io/open-docx-editor/), the editor with a sample document; open your own
with File > Open or by dragging a .docx onto the page. (It runs entirely in your browser: nothing is uploaded.)

## Install

```sh
npm install open-docx-editor
```

To try it without installing, clone the repository and run `npm install` and `npm run dev`, then drag a .docx onto the
page at http://localhost:8080.

## Quick start

With a bundler (Vite, webpack, esbuild...):

```js
import { mountDocxEditor } from 'open-docx-editor';
import 'open-docx-editor/style.css';

const editor = mountDocxEditor(document.getElementById('editor')); // a sized element, e.g. height: 100vh

// open from a URL, File or bytes
const bytes = await (await fetch('/files/report.docx')).arrayBuffer();
await editor.open(bytes, 'report.docx');

// get the edited file back (Uint8Array), e.g. to upload it
const out = editor.save();
await fetch('/files/report.docx', { method: 'PUT', body: out });
```

Without a build step: download `open-docx-editor.js` and `open-docx-editor.css` from the [Releases page](https://github.com/janderson12345/open-docx-editor/releases)
(or build them with `npm run build`, into `dist/`) and put them beside your page:

```html
<link rel="stylesheet" href="open-docx-editor.css">
<div id="editor" style="height: 100vh"></div>
<script type="module">
  import { mountDocxEditor } from './open-docx-editor.js';
  const editor = mountDocxEditor(document.getElementById('editor'));
  await editor.open(await (await fetch('report.docx')).arrayBuffer(), 'report.docx');
</script>
```

Using it from ASP.NET (MVC, Blazor or Web Forms)? See [docs/aspnet.md](docs/aspnet.md).

Lower-level pieces are exported too: `DocxSession` (read/write a package), `createEditor`,
`createToolbar`, `schema` and `commands`, so you can build your own UI.

### Headless (Node)

```js
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { DocxSession, setXmlImpl } from 'open-docx-editor';
setXmlImpl({ DOMParser, XMLSerializer });

const s = new DocxSession(fs.readFileSync('in.docx'));
console.log(s.doc.textContent);       // ProseMirror document
fs.writeFileSync('out.docx', s.save());
```

## What works

- **Open** any `.docx`: paragraphs, styles (from the file's `styles.xml` and theme fonts), bold /
  italic / underline / strike / colour / size / font / highlight / super- and subscript,
  alignment, spacing and indents, bullet and numbered lists (labels computed from
  `numbering.xml`, nested levels), tables (column widths, horizontal and vertical merges),
  hyperlinks, inline images, tabs, line and page breaks, tracked changes (shown, recorded with Track on, accepted
  or rejected one at a time or all at once), footnotes and endnotes, paragraph borders and shading, table styles (borders, shading,
  header row and banding from `tblLook`), cell borders, shading and vertical alignment, table
  indents, centred and floating tables (text wraps beside them), frames (paragraphs with `w:framePr`, placed and
  wrapped as Word places them), imported content (`w:altChunk` HTML or text, shown as the paragraphs Word makes of
  it, and saved back as it was until edited), and charts (column, bar, line,
  area, scatter, pie and doughnut, drawn from the values cached in the file). Pictures and
  shapes it can't draw yet keep their size, so the text around them sits where Word puts it.
- **Tracked changes and markup**: Review > Markup view shows All Markup, Simple Markup, No Markup or Original (the
  document as it was before its changes, as Word shows it). Columns inserted in a table while tracking are recorded as
  cell insertions; revision times are kept as Word writes them.
- **Pages** like Word's print layout: page size and margins from the document, page breaks,
  keep-with-next headings, and the default header and footer on every page with live
  `PAGE` / `NUMPAGES` fields. Section breaks start new pages. Tables continue on the next page
  and repeat their header rows; a row that doesn't fit breaks inside its cells, line by line,
  unless it is marked "can't split". Long paragraphs continue line by line with Word's
  widow/orphan control and keep-lines-together, as in Word.
- **Measured like Word**: line heights use each font's own "single" spacing (measured in the
  browser for fonts it doesn't know), no kerning or ligatures unless the document asks, thin
  table borders take their true width, and older documents keep Word's older table rules,
  so documents land on the same pages as in Word. Checked line by line against PDFs exported by Word itself: 151
  documents, every one on Word's page count, 99.0% of their lines within 1px of Word's (docs/LAYOUT.md). The same in
  a standards-mode page (any page with a doctype) as in a quirks-mode one. Older documents' compatibility rules are
  followed by name where Word's PDFs show they change the layout (a page break's paragraph mark, list labels' tab
  stops, automatic paragraph spacing; see docs/LAYOUT.md).
- **Right-to-left text** (Hebrew, Arabic): right-to-left paragraphs (w:bidi) start at the right margin, with their
  indents, first-line indents, alignment and tab stops mirrored as in Word, and numbers and English inside them in
  order; right-to-left tables (w:bidiVisual) put their first column at the right. Checked against Word's PDFs for a
  Hebrew and an Arabic test document (`test/fixtures/rtl-*.docx`): every line where Word puts it. PDF export draws
  them shaped and in order when it has the real font (with the PDF standard fonts, Hebrew and Arabic come out as "?").
- **The smaller things Word draws**, each checked against Word's own PDF of a test document (`test/fixtures/mkfeatures.mjs`):
  text and picture watermarks, legal numbering ("1.1.1" under a Roman "I."), picture bullets, text effects (outline,
  shadow, emboss, imprint), emphasis marks, borders round runs of text, a section's vertical alignment (centred,
  bottom, spread over the page), and line numbers (every line or every nth, restarting each page or section or going
  on, at the distance the section sets). They print and go into PDF export too.
- **Fit text** (text spread or squeezed to a width, Word's Asian Layout > Fit Text), **numbering** in every common
  format (letters, Roman, ordinals, words, circled and full-width numbers, Arabic, Hebrew, Russian and Chinese), and
  list numbers in their own bold, italic and colour, each checked against Word's PDF of a test document.
- **Strict Open XML** documents (Word's "Strict Open XML Document") open like any other; they are saved in the
  ordinary (Transitional) form, which every version of Word opens.
- **Fast on long documents**: a 300-page file opens in about a second in Chromium; off-screen
  pages skip layout, and pagination runs when you pause typing.
- **Edit** with Word-like keys: Ctrl+B/I/U, Ctrl+E/L/R/J, Enter continues lists and turns a
  heading into its "next" style, Tab / Shift+Tab change list level or move between table cells,
  Shift+Enter line break, Ctrl+Enter page break, undo/redo. Toolbar: styles, font, size, colour,
  lists, links, tables (insert, add/delete rows and columns, merge/split cells).
- **AutoCorrect as you type**, Word's everyday set: curly quotes, (c) (r) (tm) to © ® ™, three dots to an ellipsis,
  arrows (--> <-- ==>), dashes (a double hyphen between words to an em dash, a spaced hyphen to an en dash), ½ ¼ ¾, a
  sentence's first word and a lone "i" capitalised, and common misspellings. Ctrl+Z after a correction gives back what
  was typed. Off with `features: { autocorrect: false }`.
- **Find and Replace** (Ctrl+F, Ctrl+H): match case, whole words, and patterns (regular expressions: `\d{4}`,
  `colou?r`), with what a pattern's parentheses caught reused in the replacement as `$1`, `$2`. As in Word, it goes
  through the text boxes, headers, footers, footnotes and endnotes after the main text, opening each with the match
  selected, and Replace All changes them all.
- **Fields worked out as Word works them out**: in headers and footers, a running head (STYLEREF "Heading 1"), page
  references (PAGEREF) and cross-references (REF) on every page, with PAGE and NUMPAGES. In the body, fields keep their
  saved results, as in Word, until updated: F9, References > Update Fields or `api.updateFields()` gives every page
  number, page and cross-reference (with its paragraph number, or "above", "below", "on page 3") and the table of
  contents its value from the pages as laid out (one undoable step). Printing and PDF export update them first, as
  Word does when it prints.
- **Drafts** (opt in with `drafts: true`): unsaved changes are kept in the browser a moment after each edit and offered
  back, Restore or Discard, when the same file is opened again after a closed tab or a crash; cleared once saved. They
  stay on that computer, in the site's own browser storage.
- **Save** back to `.docx`. Only the parts an edit touches are rewritten (`word/document.xml`, and the headers,
  footers, notes, comments, relationships or numbering it changed); every other part of the package is kept byte for
  byte, and paragraphs keep the identities Word gave them (`w14:paraId`, `w:rsid*`). Content the editor doesn't model (or models only in part: fields, shapes, equations' own XML...) is
  carried through as raw XML.
- **Headers and footers**: double-click a page's top or bottom margin to edit its header or footer
  (or Insert > Header & Footer). The bar over it has Word's options: Different First Page (per section),
  Different Odd & Even Pages, and Link to Previous for later sections.
- **Pictures**: TIFF pictures (scans and faxes, which browsers can't show themselves) are shown too, decoded when the
  document opens and kept in the file as they were. Insert > Picture adds one from the device, and pasted pictures are
  stored in the document; select a picture and drag its handles to resize
  it (corners keep its proportions). Saved at the new size in its own drawing XML. Cropped pictures are shown cropped.
- **Wrap Text** (Layout > Wrap Text, or a picture's context menu): In Line with Text, Square, Tight, Through, Top and
  Bottom, Behind Text and In Front of Text, for a picture or shape, saved as Word writes it.
- **Text boxes**: Insert > Text Box adds one in line with the text; double-click one to edit its text; click away (or Escape) to put it back. Both the shape
  and the copy Word keeps for older versions are updated. In a group of shapes, the box you double-click opens.
  Select one and drag a handle to resize it (its text wraps again), or drag a floating one to move it; from the
  keyboard, the arrow keys move a selected floating box (Ctrl+arrows by 1px), Shift+arrows resize it, Esc returns to
  the text.
- **Equations** (Word's Office Math): drawn as Word lays them out (fractions, roots, sums and integrals with their
  limits, brackets, matrices, functions, limits, accents, equation arrays), in the text or on lines of their own, on
  screen, in print and in PDF export. They are saved exactly as they were; editing an equation isn't there yet.
- **Spelling**: the browser's own spell checker, in each piece of text's language: the document's, or another one a
  run is marked with (w:lang); text Word is told not to check (w:noProof: code, names) isn't checked.
- **Charts**: double-click one (or select it and press Enter, or right-click > Edit Chart Data...) to change its title,
  its type (column, bar, line, area; a pie stays a pie) and its data in a grid, categories down and series across,
  adding or removing either. Saving writes the chart's cached values and the workbook embedded beside it (its table
  resized), so Word shows the new chart and its Edit Data opens the same numbers. One undoable step.
- **Show/Hide ¶** (Home > Paragraph, Ctrl+Shift+8): spaces, paragraph marks, tabs, line and page breaks, drawn over
  the page without moving any text, and never printed.
- **Page Setup** (Layout > Page Setup): margins, gutter, header and footer distances, orientation and paper size (a
  custom one too), for the section the cursor is in or the whole document; Layout's Margins, Orientation and Size
  menus change the cursor's section, as Word's do.
- **Columns**: Layout > Columns sets one, two or three equal columns for the cursor's section, or columns of their
  own widths and gaps (More Columns), with or without a line between; right-to-left sections fill from the right.
- **Comments**: add (Ctrl+Alt+M), reply, edit, resolve and delete, in the margin as in Word; replies and
  resolved state are written to Word 2013+'s commentsExtended part so Word shows the threads.
- **Restricted editing** (Word's Restrict Editing, Mark as Final, Always open read-only) enforced as in Word; see
  [Restricted editing](#restricted-editing).
- **New** blank documents, generated in code.

## Not yet

- Tracked cell deletions (w:cellDel) are shown, accepted and rejected, but not recorded while tracking (the editor
  has no "delete cells, shift left" of its own).
- Co-editing shares the body, comments, pictures, page setup, headers, footers and notes (each as it's closed after
  editing), but not styles or numbering definitions added after opening: the others get those when the document is
  saved and reopened.
- PDF export falls back to the PDF standard fonts where it can't get the real one: in browsers without Local Font
  Access (Firefox, Safari), when the user declines it, and for fonts neither installed nor embedded in the document.
  Lines and pages are the layout's either way, but characters outside Windows-1252 (Hebrew, Arabic, Greek beyond
  the Symbol font's) come out as "?".
- In newspaper columns: a table longer than a whole column; line numbers sit left of the page's text, not each column.
- Safari (WebKit) rounds each line to a whole pixel, so long documents can break a page a line earlier than Word.

See the roadmap in docs/ARCHITECTURE.md.

## Develop

```sh
npm install
npm test            # round-trip + edit tests in Node
npm run dev         # demo at http://localhost:8080 (drag a .docx onto it)
npm run test:e2e    # opens every fixture in Chromium, edits, saves, checks the output; accessibility checks
node tools/check-ooxml.mjs test/out   # saved files: well-formed, every part and relationship in place
```

See CONTRIBUTING.md for where things are and how changes are checked.

Fixtures in `test/fixtures` are generated by scripts in the same folder (the `docx` npm
package, python-docx, and a script of their own for each feature; `other-app.docx` was saved from another word processor,
to test what other programs write), so the repository contains no third-party documents.

## Legal

This is an independent implementation of the published Office Open XML standard
(ECMA-376 / ISO/IEC 29500). It contains no Microsoft code, templates, fonts or other assets.
Microsoft Word is a trademark of Microsoft Corporation; this project is not affiliated with
or endorsed by Microsoft.

## Fonts

Word documents mostly use Microsoft's fonts (Calibri, Cambria, Arial, Times New Roman, Courier New). The editor lays
text out in the fonts installed on the computer, as Word does, so on a computer without those fonts lines can break
in different places. Two things help:

- **Metric-compatible fonts**, with the same character widths as Microsoft's: Carlito (Calibri), Caladea (Cambria),
  and Arimo, Tinos and Cousine (Arial, Times New Roman, Courier New), named in the stylesheet Liberation Sans, Serif
  and Mono, the names documents use for them. They're bundled (SIL Open Font License
  and Apache License 2.0, 1.6 MB, Latin and Latin Extended); import them once and they're used wherever the
  original is missing. Each matches its original's character widths and line height:

  ```js
  import 'open-docx-editor/fonts.css';
  ```

  Without a bundler, link the package's `src/fonts/fonts.css` (or `dist/fonts/fonts.css` from a build), keeping its
  `files/` folder beside it.
- **Aptos** (Word's newer default font, which browsers rarely have) is drawn by stand-in fonts built into the editor's
  stylesheet, each letter at Aptos's width, so lines break where Word breaks them.
- **Knowing when a font is missing**: `onMissingFonts` reports the fonts a document asks for that the computer lacks,
  and what it shows instead, so you can tell the user.

  ```js
  mountDocxEditor(el, { onMissingFonts: (fonts) => console.warn(fonts) }); // [{ font: 'Gill Sans', substitute: null, standIn: false }]
  const missing = await editor.missingFonts();
  ```

## PDF export

`open-docx-editor/pdf` draws a PDF from the editor's own page layout with pdf-lib: real, selectable text where the
layout put each line, pictures, borders, shading and list labels, a PDF page per page. Text is set in the real fonts,
embedded as subsets: the document's own embedded fonts, else the computer's installed ones (Chromium browsers' Local
Font Access, which asks the user once; `exportPdf(editor, { localFonts: false })` doesn't ask), else the PDF standard
fonts. Kept out of the main bundle; pass it in and File > Export downloads the PDF.

```js
import { exportPdf } from 'open-docx-editor/pdf';
const editor = mountDocxEditor(el, { pdf: exportPdf, onSavePdf: (bytes, name) => upload(bytes, name) });
const bytes = await editor.exportPdf(); // or editor.downloadPdf()
```

## Co-editing

`open-docx-editor/collab` edits the document together in real time over Yjs: the body through y-prosemirror, comments,
headers, footers and notes too, the .docx it came from shared so every editor lays out the same pages, others'
cursors from your provider's awareness,
and an undo of each editor's own edits. Bring the transport (y-websocket, y-webrtc...); mount once it has synced.

```js
import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import { collaboration } from 'open-docx-editor/collab';
const ydoc = new Y.Doc();
const provider = new WebsocketProvider('wss://your-server', 'contract-42', ydoc);
provider.once('sync', () => {
  const collab = collaboration(ydoc, { awareness: provider.awareness, user: { name: 'Ann', color: '#c2185b' } });
  mountDocxEditor(el, { collab, author: 'Ann' }); // the room's document, or a new one as its first
});
// Or mount at once and let the editor wait for the sync before choosing (a blank page meanwhile):
// collaboration(ydoc, { awareness, user, ready: new Promise((r) => provider.once('sync', r)) })
```

## Documents in a WOPI host

`open-docx-editor/wopi` edits a file a WOPI host serves (WOPI is the open protocol document stores use to let a web
editor work on their files). Your host gives the page the file's WOPI URL and an access token for the user; the
editor reads the file's information and contents, locks it while it's open (refreshing the lock every 15 minutes),
saves it back a moment after each edit and on Ctrl+S, and unlocks it when the page closes. A user who may only read,
or a file someone else has locked, opens read-only.

```js
import { mountWopiEditor } from 'open-docx-editor/wopi';
const app = await mountWopiEditor(el, {
  src: 'https://files.example.com/wopi/files/42', // the file's WOPI URL
  accessToken,                                      // the host's token for this user
  onSaved: () => showStatus('Saved'),
  onError: (e) => showStatus(e.message),            // e.status: the host's HTTP status (409: locked by someone else)
});
// app.save(), app.close() (saves, unlocks, removes the editor); app.wopi.info: the host's CheckFileInfo
```

The host must allow the editor page's origin (CORS) on its WOPI URLs, including the `X-WOPI-Override` and
`X-WOPI-Lock` request headers and the `X-WOPI-Lock`, `X-WOPI-LockFailureReason` and `X-WOPI-ItemVersion` response
headers. Keep the token out of logs and addresses: hand it to the page the way the protocol suggests (a form POST, or a message
from your host's frame) and give it a short lifetime. (`demo/wopi.html` takes it in the URL's fragment, which isn't
sent to servers, for trying things out; the browser's history keeps that address.) Not supported: proof keys (a host that requires them
checks a signature only a server can make), Save As (PutRelativeFile) and the host frame's message interface.
Tested against a mock host on another origin (`test/wopi-host.mjs`, `test/wopi.mjs`), not yet against a production one.

## AI agents

An AI assistant (or any program) can read and edit Word documents through `open-docx-editor/agent`, in plain terms
rather than raw XML: paragraphs with ids (`p12`) that stay with their paragraph through edits, the outline, a Markdown
view, search, and edits such as replace text, rewrite a paragraph, insert or delete one, comment, reply, and fill form
fields. **Every edit is a tracked change under the agent's own name** unless asked otherwise, so a person reviews it
in Word or in the editor and accepts or rejects it. Saving keeps everything else in the file as it was.

```js
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from 'open-docx-editor';
import { openDocument } from 'open-docx-editor/agent';
setXmlImpl({ DOMParser, XMLSerializer });                    // in Node
const doc = openDocument(bytes, { author: 'Contract assistant' });
doc.markdown();                                              // <!-- p0 --> # Agreement ...
doc.replace('thirty (30) days', 'sixty (60) days', { id: 'p14' });
doc.addComment('p22', 'This clause conflicts with 4.2.', { quote: 'exclusive' });
const edited = doc.save();                                   // .docx bytes
// In the editor: mountDocxEditor's api.agent({ author }) gives the same on the open document, live.
```

**MCP server.** `open-docx-editor-mcp` makes these tools available to AI assistants that speak the Model Context
Protocol (Claude Desktop, Claude Code and others): `list_documents`, `read_document`, `get_outline`,
`get_paragraphs`, `search`, `replace_text`, `set_paragraph_text`, `insert_paragraph`, `delete_paragraph`,
`add_comment`, `reply_to_comment`, `list_comments`, `list_changes`, `get_fields`, `set_fields` and `save_document`.
It reads and writes only inside the folder it's given, saves to a new file unless told to overwrite, and makes no
network requests. In Claude Desktop's configuration:

```json
{ "mcpServers": { "word": { "command": "npx", "args": ["-y", "-p", "open-docx-editor", "open-docx-editor-mcp",
  "--root", "C:\\Users\\me\\Documents\\Contracts", "--author", "Claude"] } } }
```

## React, Vue and Angular

Thin wrappers around `mountDocxEditor`, taking its main options as props or inputs (`fileName`, `author`, `readOnly`,
`features`, `toolbar`, `forms`, and the `onChange`, `onSave`, `onOpened` and `onError` callbacks); for the others
(`pdf`, `collab`, `drafts`, `onMissingFonts`...) mount with `mountDocxEditor` itself. A new `src` opens that document;
`readOnly`, `forms` and `features` can change at any time; unmounting destroys the editor. Import the stylesheet once.

```js
// React 17+
import { DocxEditor } from 'open-docx-editor/react';
import 'open-docx-editor/style.css';
const ref = useRef(null); // ref.current.save(), .download(), .getFieldValues()...
<DocxEditor ref={ref} src={file} author="Ann" readOnly={false} onChange={(view) => {}} onSave={(bytes, name) => upload(bytes)} />
```

```js
// Vue 3
import { DocxEditor } from 'open-docx-editor/vue';
// <DocxEditor ref="editor" :src="file" author="Ann" @change="..." @opened="..." />   then this.$refs.editor.api.save()
```

```ts
// Angular 15+ (standalone component, TypeScript source compiled by your app)
import { DocxEditorComponent } from 'open-docx-editor/angular';
// imports: [DocxEditorComponent]
// <odx-docx-editor [src]="file" author="Ann" (changed)="..." (ready)="api = $event" />
```

With no `src` the editor starts with a blank document. Types for the API are in `src/index.d.ts`.

## Read-only mode and turning features off

```js
const editor = mountDocxEditor(el, {
  readOnly: true,                         // show the document only (no ribbon, nothing editable)
  features: { review: false, comments: false }, // hide a tab by name, or a group by its label on any tab
  toolbar: true,                          // false hides the ribbon but keeps editing
});
editor.setReadOnly(false);                // switch at any time
editor.setFeatures({ review: true });
```

Tab names: `file`, `home`, `insert`, `layout`, `references`, `review`, `view`.
Group names are their labels in lower case with dashes, for example `comments`, `tracking`, `markup`,
`header-and-footer`, `footnotes`, `table-of-contents`, `bookmarks`, `page-background`.
`comments: false` also hides comments in the margin and blocks adding them.
Read-only and forms mode show the document with no tracked-change markup (as Word's "No Markup").

## Restricted editing

A document protected with Word's **Review > Restrict Editing** is enforced as Word enforces it:

- **No changes (read only)**: nothing can be changed, except the parts left editable for everyone or for this user
  (highlighted, as in Word); **Comments**: only comments can be added (and those parts edited); **Tracked changes**:
  every edit is tracked, Track Changes can't be turned off and changes can't be accepted or rejected; **Filling in
  forms**: forms mode, with any sections the form leaves unprotected edited freely.
- **Limit formatting to a selection of styles**: the locked styles aren't offered and direct formatting is refused.
- **Mark as Final**, **Always open read-only** and a **password to modify** open the document read-only, with Edit
  Anyway (or the password) in a notice above the ribbon.

Review > Restrict Editing starts and stops protection, with an optional password (hashed as Word hashes it, so Word
accepts it, and Word's own passwords work here). From code:

```js
const editor = mountDocxEditor(el, { author: 'Pat Lee', editors: ['pat@example.com'] }); // who may edit which parts
editor.protection;                                      // { edit: 'comments', enforced: true, password: true, ... }
await editor.protect({ edit: 'trackedChanges', password: 'secret' });
await editor.unprotect('secret');                       // false if the password is wrong
await editor.editAnyway();                              // a document marked final, or to open read-only
```

Like Word's, this protection steers the editor; it isn't encryption, and anyone with the file can remove it. Where it
must hold, the server enforces it (a WOPI host that says the user can't write gets a read-only editor);
`features: { protection: false }` ignores it, for a host that enforces its own rights.

## Form fields (content controls)

Word content controls (Developer tab: Plain/Rich Text, Check Box, Drop-Down List, Combo Box, Date Picker)
are read and written by their **Tag**, or by their **Title** when there is no tag.

```js
const editor = mountDocxEditor(el, { forms: true }); // only the fields can be filled in; no ribbon
await editor.open(bytes);
editor.setFieldValues({ CustomerName: 'Acme Corp', Agree: true, Color: 'G', Due: new Date() }); // returns names not found
editor.getFieldValues();  // { CustomerName: 'Acme Corp', Agree: true, Color: 'G', Due: '11/1/2026', ... }
editor.getFields();       // [{ tag, alias, kind, value, checked?, items?, block }] in document order
editor.setForms(false);   // back to normal editing
```

Values: check boxes are `true`/`false`; drop-downs use the item's value (falling back to its text);
an empty string puts the field's placeholder back; a multi-paragraph field takes `\n` between lines.

Without an editor (for example on the server when a document is checked in):

```js
import { setXmlImpl, readFieldValues, fillFields } from 'open-docx-editor';
setXmlImpl({ DOMParser, XMLSerializer });          // in Node, e.g. from @xmldom/xmldom
const values = readFieldValues(bytes);             // store in your database
const { bytes: filled, missing } = fillFields(bytes, values); // merge values into a template
```

## Accessibility (Section 508 / WCAG 2.1 AA)

- Checked with axe-core (WCAG 2.0/2.1 A and AA rules) in light and dark mode, on every ribbon tab, with menus and Find open: no violations in the editor's own interface.
- Keyboard: F6 moves between the document and the ribbon (as in Word). Left/Right/Home/End switch tabs, Tab moves through a tab's controls, Enter or Space runs them. In menus, the arrow keys move between items and Esc closes the menu and returns to its button. Esc in the ribbon goes back to the document.
- Focus is always visible. Ribbon tabs, menus and the document carry ARIA roles and names. The document has its language set from the .docx.
- The document's own content is the author's responsibility: Review > Check Accessibility flags missing alt text, empty headings and similar.
- Conformance draft with evidence per criterion: docs/ACCESSIBILITY.md. Tested with the NVDA screen reader in Edge:
  headings, toggles, menus, comments, fields and Word's shortcuts are announced. (Bullets and numbering aren't
  announced as a list: Edge doesn't expose a list role on editable paragraphs.)

## Browser support

Developed and measured in Microsoft Edge and Chrome (Chromium) on Windows, where the layout matches Word's best: it
uses the fonts installed on the computer, as Word does. Other Chromium browsers behave the same.

Checked against Edge on all 151 measured documents:

- **Firefox** lays them out the same: every paragraph on the same page with the same lines in all but a few dozen of
  about 19,000, page counts the same in 150 of 151. Justified lines can differ slightly.
- **Safari (WebKit)** sets each line's height to a whole pixel, so long documents drift up a little line by line and a
  page can break a line earlier (page counts the same in 142 of 151, checked with Playwright's WebKit on Windows).

Editing works in all three. PDF export in Firefox and Safari uses the PDF standard fonts (no Local Font Access).

## Documentation

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md): how it works and a map of every module
- [docs/LAYOUT.md](docs/LAYOUT.md): how layout is matched to Word
- [docs/ACCESSIBILITY.md](docs/ACCESSIBILITY.md): WCAG 2.1 AA conformance notes
- [docs/aspnet.md](docs/aspnet.md): using it from ASP.NET
- [CONTRIBUTING.md](CONTRIBUTING.md): setting up, testing, and changing the code
- [CHANGELOG.md](CHANGELOG.md), [SECURITY.md](SECURITY.md)

## License

MIT: see [LICENSE](LICENSE). The libraries it uses are permissively licensed too: see
[THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
