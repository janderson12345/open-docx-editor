# Changelog

## Unreleased

## 0.7.0 (2026-11-11)

Since 0.6.0: frames, columns of their own widths, Page Setup, Wrap Text, the Original markup view, imported content
(altChunk), Find and Replace through every part of the document, reference fields and Update Fields, restricted
editing, an interface and MCP server for AI agents, TIFF pictures, Aptos drawn evenly, faster typing in long
documents, and many smaller matches to Word in reading, drawing, editing and saving.

- **Aptos drawn evenly**: on computers without Aptos (Word's default font), its stand-in drew Arial at a different size
  for each letter to reach Aptos's widths, so a capital T or J came out small and a t or an l tall. The stand-in is now
  a set of small fonts built from Arimo (Arial's design), each letter narrowed or widened sideways only, all at one
  height. Letter widths are unchanged; in the comparison with Word, 4 of 35,101 lines break one word differently.

- **Imported content (w:altChunk)**: the HTML or plain text a document carries for Word to convert when it opens it
  (as mail-merge systems write rich-text fields, often in table cells) is shown as the paragraphs Word makes of it, in
  place of a placeholder: HTML paragraphs in Times New Roman 12pt with automatic spacing, text outside them in the
  document's Normal style, headings, lists, links, tables, bold, italic, underline, fonts, sizes and colours, a blank
  field as one empty paragraph. It can be searched, counted, printed and edited. Saved unedited, each chunk is written
  back as it was, its part kept; once edited, as the paragraphs it now is, as Word saves it. (RTF and whole-document
  chunks stay as placeholders.)

- **Faster typing in long documents**: the layout's own measurements (tab widths, row heights, column widths) no longer
  look like edits to the editor, which had re-read and redrawn much of a document of tables at each keystroke; tabs,
  merged rows and print baseline shifts are measured again only where something changed; tracked
  changes' balloons are drawn in one pass. A keystroke in a 20MB document of tables: about 0.9s, was about 3s; in an
  800-row table with 7,200 tracked changes: about 0.1s, was about 3s; in a 750-page report: about 0.3s, was 0.5s.
- **Page Setup** (Layout > Page Setup): margins, gutter, header and footer distances, orientation and paper size (a
  custom one too), applied to the section the cursor is in or to the whole document. Layout's Margins, Orientation,
  Size and Columns menus change the section the cursor is in (they had changed the last section only).
- **Wrap Text** (Layout > Wrap Text, and a picture's context menu): In Line with Text, Square, Tight, Through, Top and
  Bottom, Behind Text, In Front of Text, for the selected picture or shape; drawn so and saved as Word writes it.
- **Original view** (Review > Markup view > Original): the document as it was before its tracked changes, as Word shows
  it: insertions left out, deletions as plain text, paragraphs split or joined by a tracked paragraph mark as they were.
- **Frames** (paragraphs with w:framePr): placed as Word places them, across the page from the page, the margins or the
  text, down the page from its top or below where they come, the text beside them; paragraphs with the same frame
  settings make one frame. They had been drawn in the flow of the text.
- Paragraph borders and shading are printed and exported to PDF (they had been left out).
- **Columns of their own widths** (Layout > Columns > More Columns with "Equal column width" off): each column at its
  width and place, the line between them midway in each gap; a paragraph going on in a column of another width is laid
  out again at that width, as Word does.
- **Find and Replace in every part of the document**: after the main text, Find goes on through the text boxes,
  headers, footers, footnotes and endnotes, as Word's does, opening each with the match selected; Replace All
  replaces in them all.
- **Restricted editing**, as Word enforces it: documents protected read-only, for comments or for tracked changes,
  with their editable parts for everyone or for named users (`editors`) highlighted and open; forms with unprotected
  sections; formatting limited to the unlocked styles; Mark as Final, Always open read-only and a password to modify
  (read-only until Edit Anyway). Review > Restrict Editing starts and stops protection with an optional password,
  hashed as Word hashes it (`api.protection`, `protect()`, `unprotect()`, `editAnyway()`).
- **AI agents**: `open-docx-editor/agent` (`openDocument`, and `api.agent()` in the editor) reads a document as
  paragraphs with stable ids, an outline or Markdown, and edits it in plain terms (replace, rewrite, insert, delete,
  comment, reply, fill fields), each edit a tracked change under the agent's name. `open-docx-editor-mcp`, an MCP
  server, gives AI assistants those tools on the .docx files in one folder. Word reads their tracked changes and
  comments as its own.
- `@xmldom/xmldom` is now a dependency (the MCP server reads XML in Node with it).
- **Reference fields**: STYLEREF running heads, PAGEREF and REF in headers and footers worked out for each page, as
  Word does; Update Fields (F9, References > Update Fields, `api.updateFields()`) for the body's page numbers,
  cross-references and table of contents; printing and PDF export update them first, as Word does. A running head
  that names a character style finds text in it. Checked against
  Word's PDF of a test document. A table of contents' page numbers follow a section's own numbering (roman
  numerals, restarting) when it's updated.
- **TIFF pictures** (scans and faxes: uncompressed, LZW, PackBits, Deflate, JPEG and CCITT fax compression) are shown, on
  screen, in print and in PDF export; the document keeps the TIFF. Browsers can't show TIFF themselves.
- **Saving keeps more as Word wrote it**: paragraph identities (`w14:paraId`, `w14:textId`, `w:rsid*`); footnotes and
  endnotes nothing refers to; revision and comment times as Word keeps them (wall-clock `w:date`, `w16du:dateUtc`);
  table rows inside nested content controls (a repeating section and its items), which had been dropped; a highlight
  colour Word has no name for is written as run shading, not an invalid `w:highlight`.
- **Drawn as Word draws it**: cropped pictures; floating drawings stacked in Word's drawing order; drawings placed
  against the page or margins vertically at that height on their page; footnotes longer than a page continue on the
  following pages; percentage and solid shading patterns; complex-script formatting (`w:bCs`, `w:iCs`, `w:szCs`) for
  right-to-left runs; Chinese counting and Hebrew list and page numbers; hidden list numbers left out; built-in
  Heading 1 to 9 and Title used without being defined; newspaper columns filling from the right in right-to-left
  sections, with the line between midway in each gap; autofit tables whose first row is merged across the columns; a
  table of contents entry whose number doesn't fit moved to the next line with its leader.
- **Editing as in Word**: Hebrew or Arabic typed or pasted becomes right-to-left text, with Right-to-Left and
  Left-to-Right buttons and Distributed alignment; Insert Table as wide as the text where it goes, and new rows take the
  formatting of the row they're made from; pasted pictures stored in the document, linked pictures followed on
  Ctrl+click, SVG pictures shown from their SVG; pasted content keeps its spacing, indents and cell shading; an inserted
  note reference takes the formatting around it; Remove Hyperlink removes the Hyperlink style too; the caret's
  formatting survives list upkeep after Enter; word count includes text boxes; Find matches across comment ranges and
  bookmarks.
- **Tracked changes**: a column inserted while tracking is recorded as cell insertions (`w:cellIns`) and rejecting takes
  it away; text from input methods (Chinese, Japanese...) is one tracked insertion; accepting a deleted paragraph leaves
  the next its own properties; a paragraph moved whole leaves no empty paragraph; Accept and Reject resolve every
  selected table cell.

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
