# Changelog

## Unreleased

- **TIFF pictures** (scans and faxes: uncompressed, LZW, PackBits, Deflate, JPEG and CCITT fax compression) are shown, on
  screen, in print and in PDF export; the document keeps the TIFF. Browsers can't show TIFF themselves.

## 0.6.0 (2026-11-10)

- **Fonts for computers without Office's**: `open-docx-editor/fonts.css` bundles open-licensed fonts with the widths of
  Calibri, Cambria, Arial, Times New Roman and Courier New (Carlito, Caladea, Arimo, Tinos, Cousine); with them, 99.0%
  of lines still land within 1px of Word's. `onMissingFonts` and `missingFonts()` report the fonts a document uses
  that the computer lacks.
- **Right-to-left text** (Hebrew, Arabic): right-to-left paragraphs and tables laid out as Word does, mirrored indents,
  alignment and tab stops; PDF export draws them shaped and in order.
- **The smaller things Word draws**: text and picture watermarks, legal numbering, picture bullets, text effects
  (outline, shadow, emboss, imprint), emphasis marks, run borders, a section's vertical alignment and line numbers, on
  screen, in print and in PDF export, each compared with Word's own PDF.
- **Strict Open XML** documents open (saved as Transitional).
- **Older documents' compatibility rules** followed by name where they change the layout: a page break's paragraph
  mark (splitPgBreakAndParaMark), list labels' tab stops (doNotUseIndentAsNumberingTabStop) and automatic paragraph
  spacing (doNotUseHTMLParagraphAutoSpacing; without it, no automatic spacing between two items of one list).
- **WOPI**: `open-docx-editor/wopi` opens, locks, saves and unlocks a file a WOPI host serves (`mountWopiEditor`,
  `openWopi`, `WopiError`), read-only for a user who may only read or a file someone else has locked.
- **Typing comforts**: AutoCorrect as you type (curly quotes, symbols, dashes, fractions, capitals, common misspellings;
  Ctrl+Z undoes a correction), Find and Replace with patterns (regular expressions, $1 in replacements), and drafts of
  unsaved changes kept in the browser and offered back on opening the same file (opt in: `drafts: true`; `onDraft` and
  `api.draft()` for a host's own notice). `features: { newDocument: false, openFile: false }` takes New and Open off the
  File menu.
- **Firefox** lays documents out as Chromium browsers do: text width corrections are measured as the page lays text
  out.
- **Fixed**: PDF export left out table cell borders.
- **Security**: a document's colours, highlights, paragraph alignment and font names reach the page's CSS only in safe
  forms. A crafted document could add CSS of its own (in 0.5.0), loading images from elsewhere or restyling the page
  around the editor; it could not run script.
- `missingFonts()` and `onMissingFonts` give `standIn: true` for a font drawn in one of the editor's own stand-ins.

## 0.5.0 (first public release)

The first version published. What it does, in brief (the README has the details):

- **Opens and saves .docx** in the browser with no server and no Office. Only what was edited is rewritten; anything
  the editor doesn't model is kept as it came and written back.
- **Word's page layout**: measured against Word's own PDFs of 151 real documents, 99.0% of lines land within 1px of
  Word's and every document has Word's page count. Sections, columns, headers and footers, footnotes and endnotes,
  floating pictures and tables, tables split across pages with repeated header rows. The same in standards and quirks
  mode pages.
- **Editing**: a Word-style ribbon (classic or simplified), styles, fonts, lists, tables, pictures, text boxes, links,
  bookmarks, table of contents, page setup, Find and Replace, Format Painter, the context menu, Show/Hide ¶.
- **Review**: tracked changes (recorded, shown, accepted and rejected), comments with replies and resolving.
- **Charts** drawn and editable (title, type, data, with the embedded workbook updated); **equations** drawn as Word
  lays them out.
- **Forms**: content controls filled in a forms mode, read and written by tag, also without a browser.
- **Print and PDF export** from the same layout, with the real fonts embedded where available.
- **Co-editing** over Yjs, and **React, Vue and Angular** wrappers.
- **Accessibility**: keyboard access to everything, ARIA roles and live announcements, checked with axe-core and
  tested with NVDA; conformance notes in docs/ACCESSIBILITY.md.

Earlier versions (0.1 to 0.4) were developed before the repository was published.
