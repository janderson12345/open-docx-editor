# Architecture and roadmap

## Goal

View and edit `.docx` files inside any web app, without SharePoint, Office Online or a server.
Pure JavaScript, MIT licensed, published on GitHub and npm.

"Reverse engineering" here means implementing the **public** Office Open XML specification
(ECMA-376 / ISO/IEC 29500, plus Microsoft's published notes [MS-OI29500] and [MS-DOCX]), which
Microsoft covers under its Open Specification Promise. We never copy Microsoft code, templates,
fonts or icons, and we don't call the product "Word".

## Built on

| Package | Licence | Used for |
|---|---|---|
| [ProseMirror](https://prosemirror.net) | MIT | The editing engine (model, transactions, undo, tables) |
| [fflate](https://github.com/101arrowz/fflate) | MIT | Reading and writing the .docx zip |
| [pdf-lib](https://pdf-lib.js.org) | MIT | PDF export |
| [Yjs](https://yjs.dev) | MIT | Co-editing |

Tests only: [docx](https://github.com/dolanmiu/docx) generates some test documents, and [mammoth.js](https://github.com/mwilliamson/mammoth.js)
reads saved files back as an independent check that their text survived.

## Design principles

1. **Model, layout and drawing are separate.** The browser lays out lines; `src/editor/pages.js` measures them and
   assigns them to pages, splitting paragraphs and table rows across pages.
2. **Every action is a named command**, callable from the ribbon, the keyboard or the API (`src/editor/commands.js`).
3. **Round-trip tests on every change**: every test document is opened, saved and read back.
4. **No vendor assets**: no Microsoft code, templates, fonts or icons; open-licensed fonts with the same widths are
   bundled for computers without Microsoft's (`src/fonts/`).
5. **The original package is kept.** Every part and every XML fragment the editor doesn't model is kept as it came,
   and only what the user changed is rewritten, so files go back to Word users intact.

## How it works

```
 .docx bytes ──fflate──► DocxPackage (all parts, byte-exact)
                              │
            styles.xml, theme, numbering.xml ──► CSS + list labels
                              │
     word/document.xml ──reader.js──► ProseMirror doc ◄──► EditorView (+ toolbar, commands)
                                                │
     word/document.xml ◄──writer.js─────────────┘   other parts untouched ──► .docx bytes
```

### Module map

**`src/ooxml/`: the file.** Reads and writes the .docx; knows nothing of the editor and runs in Node too.

| Module | Job |
|---|---|
| `package.js` | The zip: every part kept byte for byte, relationships, content types |
| `xml.js` | Parsing, serializing, namespaced lookups, escaping |
| `docx.js` | `DocxSession`: opens a file into a ProseMirror document, saves it back; headers, footers, notes, comments, pictures, charts and text boxes as the editor needs them |
| `reader.js` / `writer.js` | `document.xml` to ProseMirror JSON and back. Anything not modelled is kept as raw XML (`opaque_*` nodes, `extra`/`rextra` attrs) and written back as it was |
| `props.js` | Run and paragraph properties to plain objects and CSS; font stacks, line factors and print offsets per font |
| `styles.js` | Styles (`basedOn`), theme, numbering and table styles (conditional formatting) to CSS and list labels |
| `fonts.js` | Embedded fonts (de-obfuscated) and the font table's stand-ins |
| `comments.js`, `notes.js` | Comments and footnotes/endnotes parts |
| `shapes.js`, `emf.js`, `chart.js`, `omml.js` | Drawing what the browser can't: shapes and groups, EMF pictures, charts (SVG) and equations (MathML) |
| `chartedit.js` | A chart's data, title and type, read and rewritten with its embedded workbook |
| `ffdata.js` | Legacy form fields |
| `template.js` | A blank document, generated in code |
| `sectpr.js` | Page setup edits of a `w:sectPr` (margins, orientation, size, columns) as pure functions |
| `units.js` | Unit conversions (EMU, twips, px) |
| `strict.js` | Strict Open XML packages turned into the Transitional form as they open |
| `tiff.js` | TIFF pictures decoded (utif) and shown as PNG; the package keeps the TIFF |

**`src/editor/`: editing and layout.** ProseMirror plugins and the view.

| Module | Job |
|---|---|
| `schema.js` | Nodes and marks shaped like WordprocessingML |
| `editor.js` | Creates the view; the keymap; the paragraph pass that turns styles into per-paragraph decorations (lists, borders, contextual spacing, headings) |
| `pararules.js` | The paragraph pass's rules, one function each: list label and indent, line size, a picture beside text, a bullet's line, tracked-change classes, heading role |
| `commands.js` | Every action as a named command (toolbar, keys and API share them) |
| `toolbar.js` | The ribbon: its tabs and what is on them, and their state on each transaction |
| `ribbonkit.js`, `findbar.js`, `ctxmenu.js` | The ribbon's building blocks (buttons, groups, the drop-down menu); the Find and Replace bar; the right-click menu and mini toolbar |
| `editor.css` | All styles, in sections (page, comments and changes, ribbon, panels, print); the theme variables a host may set are listed at its top |
| `pages.js` | Pagination: splitting paragraphs and tables over pages, columns, footnotes, floats, drawing the sheets, print setup |
| `sections.js`, `hfstory.js`, `pagefields.js`, `rowfit.js`, `printcopies.js`, `linenumbers.js` | What pagination reads and draws: each section's page, columns and breaks; headers and footers as drawn; their fields (PAGE, NUMPAGES, document properties); Word's row sizing for merged and exact-height rows; print copies for mixed page sizes; line numbers in the margin |
| `runs.js`, `lhfix.js`, `widths.js`, `justify.js`, `tablefit.js`, `tabs.js`, `joins.js` | What makes lines match Word's: a paragraph's shared run size, line pitch, text widths, justified squeeze, autofit tables, tab stops, joined paragraphs |
| `baseline.js` | Print baseline shift per font |
| `textspans.js` | Each text node in a span of its own, so a line is as tall as its text in any document mode (Word's rule; see docs/LAYOUT.md) |
| `track.js`, `comments.js`, `balloons.js` | Tracked changes, comments and the markup column |
| `fields.js` | Content controls and forms mode |
| `hfedit.js`, `notes.js`, `txbxedit.js` | Editing headers and footers, notes and text-box text in their own small editors |
| `imgresize.js`, `charts.js` | Resizing and moving pictures and shapes; the chart dialog |
| `find.js`, `bookmarks.js`, `toc.js`, `marks.js`, `ruler.js`, `a11y.js`, `live.js` | Find and replace, bookmarks, table of contents, Show/Hide, ruler, accessibility checker, screen-reader announcements |
| `autocorrect.js`, `drafts.js`, `fontcheck.js` | AutoCorrect as you type; unsaved changes kept in the browser (IndexedDB); the fonts a document asks for that the computer lacks |

**Elsewhere:** `src/engine/printlines.js` (the print layer), `src/engine/metrics.json` (font metrics),
`src/index.js` (the public API, `mountDocxEditor`), `src/collab.js` (co-editing over Yjs), `src/wopi.js` (files a
WOPI host serves), `src/pdf.js` (PDF export), `src/fieldsNode.js` (form fields without an editor), `src/wrappers/` (React, Vue, Angular).

### From a keystroke to pages

1. A key or toolbar command dispatches a ProseMirror transaction (`commands.js`); `track.js` turns it into tracked
   changes when tracking is on.
2. The paragraph pass (`editor.js`) recomputes the decorations of the paragraphs that changed.
3. After the browser lays the column out, `pages.js` measures it, moves page-starting lines and rows to their
   pages with spacers (decorations, never document content), and draws sheets, headers and footers behind the text.
4. Printing or PDF export: `printlines.js` places each line exactly; `pdf.js` reads the same positions.
5. Saving: `docx.js` writes `document.xml` (and any changed header, footer, note, comment or chart part) into the
   package; every other part goes out as it came in.

### Testing

- `npm test` (Node): round trips (open, save, open gives the same model; other parts byte-identical; well-formed XML;
  mammoth reads the same text), edits, tracked changes, styles, content controls, equations, charts and shapes.
- `npm run test:e2e` (Chromium): every fixture opened, edited and saved; accessibility (axe and our own scan, light
  and dark, narrow widths); framework wrappers; PDF export; co-editing; forms; charts; text boxes.
- Layout against Word: see `docs/LAYOUT.md` and `tools/word-compare`.
- Saved files: `tools/check-ooxml.mjs` checks they're well-formed with every part and relationship in place.

## Roadmap

**Phase 1: MVP (done).** Open/edit/save, styles, lists, tables, links, images
displayed, lossless pass-through. v0.2 added approximate pages, headers/footers with page
numbers, table styles and borders, and large-document performance.

**Phase 2: fidelity (done).** Headers and footers (shown, then editable), footnotes/endnotes, insert
and resize images, paragraph spacing/indent UI, tab stops, borders and shading, table styles
with conditional formatting, theme colours, section breaks and columns, comments shown.
Build a corpus of real-world documents (with permission) and track fidelity per feature.

**Phase 3: exact page layout (done: 99.0% of lines within 1px of Word on 151 documents; see docs/LAYOUT.md).** A layout layer that measures the laid-out lines,
paginates with widow/orphan control, keep-with-next and table row splitting, and positions
headers, footers and footnotes per page; render pages as separate sheets. This is the largest
piece.

**Phase 4: review.** Tracked changes are done (src/editor/track.js: insertions, deletions, paragraph marks, formatting and paragraph-property changes, kept in Word's own w:ins/w:del/w:rPrChange/w:pPrChange form). Comments with replies, tracked table rows (inserted and deleted, recorded while tracking, accepted and rejected) and moved text (w:moveFrom / w:moveTo) are done too. So are tracked cell insertions and deletions (w:cellIns / w:cellDel: shown, accepted and rejected, a removed cell's row shifting left with a w:gridAfter, as Word does).

**Phase 5: collaboration and output (done).** Real-time co-editing via Yjs (`src/collab.js`, y-prosemirror) of the
body, comments, headers, footers, notes and pictures; PDF export drawn from the page layout (`src/pdf.js`, pdf-lib,
with the document's or the computer's own fonts); printing; React, Vue and Angular wrappers; accessibility checks.

**Phase 6: richer content (done).** Equations drawn (MathML), charts edited (data, title, type, with their workbook),
text boxes moved and resized, spelling in each run's language.

**Publishing.** GitHub repo with CI (GitHub Actions running both test suites), npm package,
hosted demo on GitHub Pages, CONTRIBUTING guide, and an OOXML schema validation step in CI.
