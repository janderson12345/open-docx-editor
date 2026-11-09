# open-docx-editor

Open, edit and save Word documents (`.docx`) in the browser. 100% JavaScript, no server, no
Microsoft Office or SharePoint, MIT licensed.

Version 0.5: it opens and saves real documents without losing what it doesn't understand, and lays them out on
pages as Word does: measured against Word's own PDFs of 151 documents, 99.0% of lines land within 1px of Word's, on
the same pages. See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the design and roadmap.

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
  indents, centred and floating tables (text wraps beside them), and charts (column, bar, line,
  area, scatter, pie and doughnut, drawn from the values cached in the file). Pictures and
  shapes it can't draw yet keep their size, so the text around them sits where Word puts it.
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
  a standards-mode page (any page with a doctype) as in a quirks-mode one.
- **Fast on long documents**: a 300-page file opens in about a second in Chromium; off-screen
  pages skip layout, and pagination runs when you pause typing.
- **Edit** with Word-like keys: Ctrl+B/I/U, Ctrl+E/L/R/J, Enter continues lists and turns a
  heading into its "next" style, Tab / Shift+Tab change list level or move between table cells,
  Shift+Enter line break, Ctrl+Enter page break, undo/redo. Toolbar: styles, font, size, colour,
  lists, links, tables (insert, add/delete rows and columns, merge/split cells).
- **Save** back to `.docx`. Only `word/document.xml` is rewritten (plus relationships or
  numbering when you add a link or a list); every other part of the package is kept byte for
  byte. Content the editor doesn't model (or models only in part: fields, shapes, equations' own XML...) is
  carried through as raw XML.
- **Headers and footers**: double-click a page's top or bottom margin to edit its header or footer
  (or Insert > Header & Footer). The bar over it has Word's options: Different First Page (per section),
  Different Odd & Even Pages, and Link to Previous for later sections.
- **Pictures**: Insert > Picture adds one from the device; select a picture and drag its handles to resize
  it (corners keep its proportions). Saved at the new size in its own drawing XML.
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
- **Columns**: Layout > Columns sets one, two or three equal columns for the section.
- **Comments**: add (Ctrl+Alt+M), reply, edit, resolve and delete, in the margin as in Word; replies and
  resolved state are written to Word 2013+'s commentsExtended part so Word shows the threads.
- **New** blank documents, generated in code.

## Not yet

- Tracked cell insertions and deletions (w:cellIns / w:cellDel) are shown, accepted and rejected, but not recorded
  while tracking (the editor has no "insert cells, shift right" of its own).
- Co-editing shares the body, comments, pictures, page setup, headers, footers and notes (each as it's closed after
  editing), but not styles or numbering definitions added after opening: the others get those when the document is
  saved and reopened.
- PDF export falls back to the PDF standard fonts where it can't get the real one: in browsers without Local Font
  Access (Firefox, Safari), when the user declines it, and for fonts neither installed nor embedded in the document.
  Lines and pages are the layout's either way.
- In newspaper columns: a table longer than a whole column.

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

## React, Vue and Angular

Thin wrappers around `mountDocxEditor`, with the same options as props or inputs. A new `src` opens that document;
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
uses the fonts installed on the computer, as Word does. Other Chromium browsers behave the same. Firefox and Safari
aren't tested; PDF export there uses the PDF standard fonts (no Local Font Access).

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
