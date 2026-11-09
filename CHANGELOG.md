# Changelog

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
