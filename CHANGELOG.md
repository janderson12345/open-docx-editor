# Changelog

## Unreleased

- **Stand-ins for 151 more Microsoft fonts**, from open-metric-fonts (open fonts with the originals' character widths
  and line spacing): Grandview, Seaford, Tenorite, Skeena, Bierstadt, Avenir Next LT Pro, Aptos's other weights, Aptos
  Narrow, Serif and Mono, Bahnschrift, Sitka, Century Gothic, Garamond, Franklin Gothic, the Lucida family and the
  rest of the Office and Windows set. Where the computer lacks one, the editor draws it in its stand-in and lays its
  lines out at the original's line height, so they break where Word breaks them. Loaded only for fonts a document uses,
  one file per alphabet its text needs, from npm's `@open-metric-fonts` packages on jsDelivr, or from your own copy
  (`fonts.omf`; README "Self-contained hosting"), or not at all (`fonts.omf: false`).
- **Self-contained hosting, documented and tested**: README "Self-contained hosting" shows how to serve the stand-in fonts
  from your own site (no CDN) and the Content-Security-Policy the editor needs (only `'self'`, with `data:`/`blob:`
  images and generated styles); a test runs the editor under that policy. Fixed on the way: the ribbon's drop-downs and
  colour pickers also wrote their handlers into the page as inline-script attributes (inert, but refused under a strict
  policy).
- **Resizing shapes as in Word.** A shape's corner handles now resize it freely, as Word's do (Shift keeps its
  proportions, as does a shape locked to them; pictures keep theirs, as before). Its yellow handles change its adjust
  values (an arrow's head and shaft, a rounded rectangle's corners, a callout's pointer, a block arc's angles: each of
  the standard's handles for all 187 preset shapes), the shape redrawn as it's dragged. A line has a handle at each end
  instead of eight round a box: dragging one moves that end, the line turning to follow (Shift: at 15-degree steps), a
  floating line's place moving with it. Each is one undoable step, saved as Word saves it (checked by opening the saved
  files in Word and comparing its print with ours).
- **Turning shapes, as in Word.** A selected shape has a rotate handle above it (Shift: 15-degree steps), and Alt+Left
  or Alt+Right turn it 15 degrees (Ctrl+Alt: 1), as Word's do; its other handles turn with it. Turned shapes are now
  drawn whole (they were cut off at their box) and where Word draws them: in line with the text, its room round the
  turned outline (which turning writes, as Word does) taken on the line, the shape drawn as far along as Word draws it;
  floating, turned about its centre, the text wrapping round the turned outline. Checked against shapes Word turned
  itself, printed by Word.
- **Drag a picture or shape to another place in the text**, as in Word: a selected one in line with the text, dragged by
  its body, goes where it's dropped, a caret showing where while it's dragged; with Ctrl, a copy. One undoable step.
- **Text on both sides of a floating picture or shape**, as Word wraps it (square, tight or through, "both sides"): each
  line beside it is filled from the margin up to it with as many whole words as fit, and goes on past it; a side too
  narrow for a word stays empty. Text on its left only, its right only, or the larger side, as the document says (left
  only had put the text on its right). A floating picture aligned to the column's centre is drawn there (it was at the
  margin). Checked against Word's prints of shapes at several distances from the margin.
- **A floating picture reaching past the page's foot** stays on its page, as in Word, drawn up to end at the page's edge
  when it would go past it, the lines above it going round it; its paragraph had moved to the next page. (A text box
  grown to fit its text still goes on to the next page, as Word does.)
- **Printing: lines beside a float from the paragraph above** print beside it, as on screen; they printed under it.
- **A shape drawn past its box**, as a callout's pointer below it, is now drawn whole in line with the text (it was cut
  off at the box); it takes the same room on its line as before.
- **Fixed: a shape inserted with Insert > Shapes and given an arrowhead** had its arrowhead cut off at its box until
  the document was opened again.

## 0.7.8 (2026-11-12)

- **Aptos stand-ins look more like Aptos.** Where a computer lacks Aptos (an Office cloud font browsers can't use), its
  letters are now Source Sans 3's (Adobe, SIL Open Font License) fitted to Aptos's widths and x-height, instead of
  Arial's design: closer a, t, y and G, lighter and the right height. Letter widths are unchanged, so lines break where
  Word breaks them as before.
- **Stand-in fonts load only when needed, from a CDN by default.** The editor links its stand-ins' stylesheet from
  jsDelivr (this release's files on npm) and the browser downloads a font only when a document uses it and the computer
  lacks it; the Aptos stand-ins no longer sit inside the editor's stylesheet (about 250 KB smaller). The Calibri,
  Cambria, Arial, Times New Roman and Courier New stand-ins come the same way, with no import needed. `fonts.base`
  serves them from your own server instead (copy the package's `src/fonts`), or `false` if you import
  `open-docx-editor/fonts.css` yourself.
- **Choose the font for a missing one**: `fonts.substitutes` ({ 'Font': 'Family' | { family, url } }) draws your choice
  where a font is missing, a font the browser has or one loaded from a URL you give (not fitted to the missing font's
  widths).
- **Fixed: tables on later pages jumped up when clicking.** In a document with a table at the top of each page (after
  a page break), each click that moved the caret into or out of a table pulled tables up onto the page before, on top
  of the ones there; the document itself was unchanged. Tables now stay on their pages.

## 0.7.7 (2026-11-12)

- **Command search (Alt+Q)**, as Word for the web's Search box: a box in the ribbon's tab row; type part of a command's
  name and pick it from the list of every ribbon command, each with its tab and keyboard shortcut. Up/Down and Enter
  run it (a menu opens, a list takes the focus, anything else is done and the keys go back to the document); Esc goes
  back to the document. A combo box with its list for screen readers, the number of matches said as you type.
- **Table of contents styles.** Inserting a table of contents into a document without Word's TOC styles adds them
  (TOC Heading, TOC 1–3, as Word's own template defines them), so its lines are indented by level and Word shows them
  under their usual names.
- **Documents from Word in other languages.** Word keeps a built-in style's English name but gives it the language's own
  id ("Ttulo1" for Heading 1 in Spanish). A table of contents, a pasted heading and HTML imported into the document
  now use the document's own Heading and TOC styles instead of adding English copies alongside them.
- **Insert Table uses Word's Table Grid style**, adding it to the document if it lacks it. Its lines now show while
  editing (they were only in the saved file), and the cells' text has no space after, so rows are as tall as Word's.
  A pasted table without table properties is also drawn with the lines it is saved with.
- **A built-in heading style a document uses without defining it** (no Heading 1 in its styles.xml) is now Normal, as
  in Word: plain body text, not a heading in the navigation pane, table of contents or to screen readers. It was drawn
  as Word's built-in Heading 1, which Word itself doesn't do (checked with and without latent styles, and with no
  styles part).
- **Columns for part of a document.** With text selected, Layout > Columns applies to the selected text only, with
  continuous section breaks before and after it, as Word's menu does; "One/Two/Three, from here on" applies from the
  cursor on (More Columns' "This point forward"). A break at a paragraph's start or end splits nothing. The menu's tick
  now shows the cursor's section, not the last one.
- **Shapes.** Shapes are drawn from DrawingML's own geometry: guide formulas, adjust values, arcs and the shapes'
  shaded faces (a cylinder's top, a cube's sides), for about 75 of Word's preset shapes (stars, arrows, flowchart
  shapes, brackets and braces, block arcs, polygons...) and for custom geometry, where only 7 were drawn before
  (others came out as rectangles). Each preset was checked against Word's own drawing of it. Lines get their
  arrowheads (triangle, stealth, open arrow, diamond, oval; small to large), sized as Word's. Theme colours are mixed as
  Word mixes them (a shade or tint in linear light, "darker 25%" on HSL luminance), a document without a theme takes the
  current Office theme's colours, and text in a shape without a colour of its own takes the shape's font colour. An
  inline line's arrowheads are drawn past its box as in Word, without moving the text.
- **Screen readers (checked with NVDA)**: command search no longer moves its active option as you type (NVDA said
  "unknown" and the page's name on every key); Down goes into the list and Enter runs the first match. A menu item
  with a hint reads with a pause between the two ("Edit individual documents, One document: ..."), where they ran
  together. The Mailings tab, AI Writing and Insert > Shapes were read through as expected.
- **All 187 preset shapes**, from the definitions in the Office Open XML standard (ECMA-376), generated into
  `src/ooxml/presets-data.js` by `tools/presets/build-presets.py`.
- **SmartArt** is drawn, from the drawing Word saves with each diagram: its shapes, connectors and text (DrawingML
  text: insets, anchoring, alignment, bullets, spacing, sizes and colours), colourful variations and styles. Checked
  against Word's print of eight layouts (process, block list, organization chart, cycle, bullet list, Venn, pyramid,
  block process): text within 2px. Not yet: gradient fills (drawn in their first colour), and a diagram Word lays out
  again on opening rather than as saved (the pyramid's text colour).
- **Insert > Shapes**: a gallery of the shapes, each inserted in line with the text, styled as Word's (accent 1 fill,
  a darker outline), and a line or an arrow.
- **AI writing in the editor** (Home > AI Writing, with the host's model as the `ai` option): rewrite the selection
  or the cursor's paragraph (shorter, longer, more formal or casual, spelling and grammar), summarize it or the
  document, or ask about it. Rewrites go in as tracked changes under `aiAuthor` ('AI'), only the edited words marked;
  an answer is shown to insert below or put in place of the selection (tracked) or copy. A request is cancelled when
  another replaces it, not applied if the document changed meanwhile, and a model's error is said. The editor sends
  nothing anywhere itself.
- **Mail merge**, as Word's Mailings: a Mailings tab with Select Recipients (the host's records or a CSV file), Insert
  Merge Field, Preview Results (each record's values in place of the fields on screen, record by record) and Finish
  & Merge (one document, each record on a new page). Merge fields' switches as Word formats them (`\* Upper/Lower/Caps/
  FirstCap`, `\# "$#,##0.00"`, `\@ "MMMM d, yyyy"`, `\b`, `\f`), `MERGEREC`, simple and complex fields, formatting
  kept. API: the `mergeData` and `onMerge` options, `setMergeData()`, `previewMerge()`, `mailMerge()`, and
  `mailMerge()`/`mergeFieldNames()` without an editor. Checked by opening the merged document in Word. Not yet: IF
  fields worked out per record, and merge fields in headers and footers.
- **Automatic hyphenation.** A document with Word's automatic hyphenation on (`w:autoHyphenation`) now hyphenates
  English words at line ends, on screen and in print, in every browser, as Word does. A browser that hyphenates
  English itself (Firefox) does it; where it can't (Edge and Chrome on Windows), break points come from the TeX
  English patterns (at least two letters before a break, three after), shown as soft hyphens that never enter the
  document's text. Paragraphs or styles with "Don't hyphenate" (`w:suppressAutoHyphens`) are left alone, and words in
  capitals when the document says so. Against Word's own print every line lands where Word's does; Word's dictionary
  sometimes splits a word a syllable earlier. Adds the `hyphen` package (its English patterns: about 50KB).
- **Justified text:** a paragraph's last line that fits only with its spaces squeezed (as Word 2013+ does) was
  squeezed several pixels too far; it now ends where Word's does, within a pixel.

## 0.7.6 (2026-11-11)

Zoom, and what a screen reader (NVDA) pass over the newer pieces turned up.

- **Zoom and fit-to-width**: View > Zoom (zoom out and in, a percentage, 100%, One Page, and Page Width, which keeps
  fitting as the window resizes), Ctrl+wheel, `api.zoom` and `api.setZoom()`. Only the view changes: the layout is
  measured at 100%, so the pages break the same, and printing and PDF export are at full size. The ruler, comments,
  picture handles and the header, footer, note and text box editors follow the zoom; the zoom carries over to the next
  document opened.
- With the comments column showing, the page moves aside for it only as far as the window has room, so a page wider
  than the window is never cut off at the left.
- **Find, with a screen reader**: the find bar says which match it went to and where ("4 of 8, in the header"), and
  the keys stay in it when a match is in a header, footer, note, text box or comment (their editors had taken the
  focus, so the next Enter typed into them). Ctrl+H in the find bar turns it into Replace, as in Word, rather than
  opening the browser's history.
- The notices above the ribbon (marked as final, open read-only, restricted editing, a kept draft) are read out when
  they appear, with where their buttons are.
- A comment box being edited took the keys back whenever the comments column was drawn again (an edit, an undo
  elsewhere in the document); it takes them once, as it opens. A comment Find opened closes with the find bar, and Esc
  in "Replace with" closes the find bar too.

## 0.7.5 (2026-11-11)

Since 0.7.0: fit text, more numbering formats, list numbers in their own formatting, Find in comments, pasted
formatting kept, and fixes for comments with tables or pictures and for nested fields. Each change that shows on the
page was checked against Word's own PDF of a test document.

- **Fit text** (`w:fitText`, Word's Asian Layout > Fit Text): a run, or the runs sharing its id, spread or squeezed to
  its width, the difference after each character but the last. On screen and in PDF export, the text after it starts
  where Word starts it, to the pixel.
- **Numbering formats**: ordinal ("1st"), cardinalText ("One"), ordinalText ("First"), numbers in circles, in
  brackets and with full stops, full-width digits, Arabic letters (alphabetic and abjad), Russian letters, and the
  Chinese stems and branches (digits past them, as Word does). Every label of a test document's 13 lists matches
  Word's.
- **List numbers in their own formatting**: the bold, italic, colour and underline their list level sets, else the
  paragraph mark's, as Word draws them.
- **Find and Replace in comments** too, the comment opening in the margin with the match selected.
- **Paste** keeps the text's colour, size and font (black, which some editors set on every run, and generic families
  are left out).
- **A document with no styles part** takes Word's own defaults: Aptos 12pt, 8pt after, 1.15 lines.
- **Tracked changes made here carry their UTC time** (`w16du:dateUtc`, beside `w:date`, which is on the author's clock),
  with the namespace declared and ignorable as Word declares it.
- **Comments with a table or a picture**: their table's text is shown, and they're kept as they are (Edit is offered
  only for text comments: editing here would lose the rest). A save no longer writes again, from its text, a comment
  that had no thread id, which lost such content.
- **Nested fields**: the result of a field inside another's code (an IF around a PAGE) is part of the code, not shown,
  as in Word.
- **No keyboard trap**: where a protected document can only be read, Tab goes on to the next control unless the caret
  is in a part this user may edit.
- The agent interface's text leaves out hidden text and shows symbols and note references (`[^1]`).
- The save check (`tools/check-ooxml.mjs`) also finds a prefix that `mc:Ignorable` names but nothing declares.
- pdf.js, used only by the layout measuring tool, updated past its security advisory (`npm audit`: none).
- New tests: input-method typing in the body, text boxes, headers, footers and notes; Track Changes edge cases; keys
  inside a protected document's editable parts; five save cycles giving the same file; a field written in one run;
  pictures sharing one id, as templating tools write them; unusual content controls.

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
- Right-aligned tab stops (a table of contents' page numbers) reach the margin where Calibri isn't installed and its
  stand-in is drawn: tabs are sized again when the text in their line moves or a web font arrives, and a title that
  overruns the stop by a fraction of a pixel takes its number to the next line, as in Word.
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
