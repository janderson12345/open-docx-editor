# Changelog

## Unreleased

Layout fixes found by comparing the editor with Word's prints of about 4,000 real-world documents: lines placed
within 1px of Word's went from 47% to 76%, and documents laid out exactly as Word from 1,157 to 1,820.

- **Word squeezes a justified line by a rule of its own, now the editor's** (66 generated justified documents, 4,516
  words that stuck out of a line): Word takes the word on when narrowing the line's spaces for it costs each space no
  more than half of what leaving it out would widen each by, and never by more than a quarter of a space. Every plain
  line of the set breaks as Word's, full lines and lines of three to eight words; the fitted limit this replaces got
  31 of the 4,516 wrong, the rule seven (three on the quarter's limit, four in letter-spaced text within 3% of the
  rule). Lines whose runs carry letter-spacing of their
  own (every run of a document converted from a PDF does) are squeezed as any other: they were left alone, and none
  of 27 such documents broke as Word's.
- **Tab stops are taken in rising order, as Word takes them** (34 generated lines): a stop written after one further
  along is not a stop. Stops written 900 then 880 twips sent a tab to 880; Word sends it to 900.
- **A list number is highlighted with its paragraph mark** (five generated lists): the highlight covers the number
  and its room up to the text; the mark's shading does not colour it.
- **An underline is the colour of its text** where that colour also names a theme's (a link's blue): it was black.
  And a table's line where it breaks across pages is as wide as the table's own lines (it was a hairline under
  0.75pt lines).
- **A page a fraction of a pixel wider than whole pixels prints in place**: a page of 12242 by 15842 twips was wider than
  its printed sheet, and the browser cut every page a pixel late, so each page's text stood a pixel higher than the last.
- **A bullet's taller first line holds over a tracked deletion** (16 generated lists): where a paragraph opened with a
  deletion, hidden in the view printed, the extra room for a Symbol bullet went with it, and the page ran short.
- **A header line of several floating shapes is a full line** (12 generated headers): with two or more it was a pixel
  short in Calibri, and the page's text under the header moved up with it.
- **A font's line is measured to its own units** (85 generated documents): measured at 100px it was rounded, a line of
  Tempus Sans ITC 0.08px short at 12pt.
- **A table row at a page's top starts under its own top line only** (102 generated tables): where the line over a
  row is the upper row's bottom line, the row's text starts at the page's top, moved whole or split; it stood the
  line's width lower, every page of a long table. A merged cell's lower part that sets a top line counts as the row's.
- **A run in another font deepens a Calibri line only past Calibri's own depth** (20 generated documents): runs in
  Garamond or Consolas made 11pt Calibri lines 0.2 to 0.4px taller than Word's.
- **A paragraph is no taller than its font's own line** (18 generated documents in Garamond, Georgia and Courier
  New): fonts whose single line is under 1.15em had every single-spaced paragraph, empty or not, 0.2 to 0.4px too
  tall, 6 to 15px down a page.
- **Garamond's line and baseline are Word's** (42 generated documents, 8 to 16pt): 1.125em a line, the text 0.862em
  under the line's top whatever the spacing.
- **A line crossing the foot of an exact-height row is cut there in print** (27 generated tables): Word draws such a
  line down to the row's foot and no further; the print drew it whole, over what followed the row.
- **A table row of exact height never breaks across pages** (24 generated tables): Word moves such a row whole to
  the next page, its text cut to the row or not; automatic rows break line by line. Printing broke one whose text was
  taller than the row: the row could not grow across the page's foot, the rows after it ran off the page over the
  footer, and every page after came out shifted (an 82-page document: 9% of its words in Word's place, 97% now).
- **A right-aligned list number ends at its place whatever its width** (18 generated lists): roman numbers wider than
  their hanging indent pushed their text on and ended past their place; their text starts at the indent, or at a tab
  stop of the paragraph's own before it. And a list number's width is measured as it is drawn, without kerning: a
  capital and its full stop measured a pixel narrow and kept their text against them where Word sends it to the next
  stop (two generated lists of 26 letters).
- **Justified paragraphs that hold a tab are squeezed too** (18 generated documents, 1,046 lines): a label and a tab
  before a paragraph's text, a tab that opens it, or one to a stop after its first words left the whole paragraph
  unsqueezed, and 246 of the lines broke as Word's; 1,039 do now. Word justifies, and squeezes, a line after its last
  tab.
- **A continuous section's first page is the page its text starts** (18 generated documents): a section set to have a
  different first page that starts part way down a page shows the ordinary header on the next page, as before; one
  whose text starts a page (a page break before it, or its opening headings carried over with their text) has that
  page as its first, with the first-page header and footer. A document's second page stood 27px low under the wrong
  header and ran a page over.
- **A list level's first number counts the deeper levels before it** (14 generated lists): a paragraph of a deeper
  level before any of an upper one uses that level's first number, so the upper level's own first paragraph is its
  second: after "1.1", the first top-level paragraph is "II".
- **A list number has its level's letter-spacing and size, and its text starts at the paragraph's own stop** (36
  generated lists): the number is condensed or expanded as its level says and drawn at the level's size where that
  is larger than its style's; the text after it starts at a tab stop of the paragraph's own when one lies between
  the number's end and the hanging indent.
- **Tab stops are measured from their own section's margin** (15 generated pages): in a document whose sections have
  different side margins, stops were set from the last section's, every tabbed line of the earlier sections off by
  the difference.
- **No line break at punctuation inside unspaced text** (31 generated paragraphs): Word breaks a line at its spaces
  and after hyphens and dashes; the browser also broke between a closing bracket and an opening one, after a full
  stop before a bracket, and after a question mark or an ellipsis before a letter ("(a)(1)(i)" split at a line's
  end).
- **A header's or footer's text box is set from the page's edge and shows each page's own number** (six generated
  footers): a shape placed from the page's edge stood in from the margin instead, and up to 1.7px high; a page
  number inside a text box showed the number it was saved with on every page.
- **Underlines as Word draws them: in their own colour, across tabs, through descenders** (11 generated documents):
  an underline given a colour of its own was not drawn at all; an underlined tab had no line (Word draws it across
  the tab's whole width); the browser left a gap in the line round each letter that reaches below it, where Word's
  is unbroken; and underlined spaces at a paragraph's end had a line where Word draws none.
- **A link without the Hyperlink style is drawn as its text** (eight generated documents): Word gives a link no
  colour or line of its own, only the Hyperlink style it puts on the links it makes; a contents page of unstyled
  links was blue and underlined. A link made in the editor is now given the Hyperlink style, as Word gives the
  links it makes (it was saved without it, and opened in Word as plain text).
- **A paragraph none of whose text is in its own size sits on Word's baseline** (16 generated documents, 176 lines
  within 0.7px): its lines are measured as its runs',
  not as a line of the paragraph's size holding larger text (a 12pt list paragraph in an 11pt style printed a pixel
  low). And a paragraph pushed to a page's top in one layout pass and back in a later one keeps its own space before
  (it lost it, standing 13px high).
- **A page's last line is fitted exactly, as Word fits it** (pinned with 189 generated documents printed from Word,
  all now breaking where Word's do; 143 before): Word keeps a line that ends on the bottom margin and moves one that
  passes it by a twip, whatever the line spacing. The browser lays lines and spacing out in 64ths of a pixel, a third
  of a pixel off over a page either way; the page fit now takes that rounding back. Paragraph spacing and cell
  margins are written to the nearest 64th (they were cut short: a page of table rows came out 1.3px short of Word's;
  96 generated documents now end their page within 0.5px of Word's).
- **A cell's text keeps clear of the lines at its sides** (80 generated documents): it starts and ends its side
  margin in from the cell's edge, or half the line's width where that is more. A table with no side margins held a
  third of a pixel more on each side than Word's and kept a word on a line that Word wraps, every row after it a
  line off: one 78-page document went from 21% of its words in Word's place to all but none.
- **A line may end after a hyphen or a dash whatever follows it** (112 generated paragraphs): Word breaks after a
  hyphen, an en dash or an em dash before a quote, a full stop or a currency sign too, where the browser took the
  word before the hyphen down with what followed.
- **A page number takes its field's shaded opening mark's look** (33 generated footers): where the run of a page
  number field's opening mark is shaded, Word draws the number in that run's colour, on its shading, whatever its
  saved result is set in (the page-number gallery's footer with a mention's look left on its marks: blue on grey on
  every page).
- **Underlines sit where the font puts them, as in Word** (84 generated lines in six fonts, regular and bold): Word
  draws an underline at the font's own distance under the baseline and as thick as the font makes it, the browser a
  pixel under every font; a link's underline was 0.7px above Word's under Calibri and 2.5px under Courier New. Now
  within half a pixel (the browser keeps the line on whole pixels).
- **A line holding a check box set in an East Asian font is as tall as Word's** (70 generated documents): Word's
  check box control writes its box in MS Gothic, and Word makes that line 332/256 of the text's size whatever the
  font beside it; a form of 25 such lines had ended its page a pixel low.
- **Hardening**: the document's default paragraph style id is escaped where table styles write it into a selector (a
  quote in it could have closed the selector and started a rule of its own).
- **Table styles place and repeat their tables** (60 and 12 generated documents): a table with no alignment or indent
  of its own takes its style's, one wider than its column set right overhangs at the left, and a style's header
  rows repeat on each page. **A table starts half its left line's width in** from its indent, the line being its first row's first cell's, the
  cell's own where it sets one (93 documents).
- **Exact-height rows** (60 generated documents): the last paragraph's space after counts when a cell centres or
  bottom-aligns its text, while text and spacing fit the row; a cell holding more than its row starts at the top and
  no longer makes the row taller. **A row split across pages** keeps its cell's last paragraph's space after, and a
  centred cell's paragraph spacing counts in its part of the row.
- **List numbers** (pinned with 46 generated documents): a number wider than its room pushes its text to the next
  tab stop; a level that ends its number with a space; a restarted list restarts at its first paragraph only;
  lists sharing a definition share the count; a paragraph holding only a page break is counted and shows no number;
  a number with no font of its own takes its paragraph mark's font, then its style's.
- **Tabs**: a stop between a paragraph's right indent and the margin is a stop (a contents entry's page number at
  the margin); in a centred or right-aligned paragraph tabs go to their stops and the line is then aligned as a
  whole; a boxed paragraph's line ends inside its side line. **A raised or lowered list label**, and a symbol bullet
  with no plain character, drawn in its own font.
- **Text boxes and drawn shapes**: automatic paragraph spacing is 14pt and none above a box's first paragraph; a
  shape's picture is a whole number of pixels, so its text is no longer drawn 0.2% small; italic text measured at
  its own width; justified lines.
- **Links take their look from their character style** (the editor's own blue only for links made in the editor);
  contents entries are drawn plain, as Word prints them; an underline is in its run's own colour.
- **Long documents split paragraphs across pages in print** (they had moved whole, leaving gaps); a paragraph
  holding only a page break stays at its page's foot; the Normal style's font goes over a table style's.
- **List labels in their own font size their line as Word does** (one rule, pinned with 368 generated lists printed
  from Word, each within 0.1px): a bullet or number in a font of its own makes its line taller only by what the
  label's top reaches above the text's, above the first line, whatever the line spacing: a Symbol bullet over Calibri,
  Times New Roman or Courier New, a Segoe UI bullet over Calibri, a larger Wingdings bullet; a Courier New "o" adds
  nothing until it is larger than the text.
- **Headers and footers**: spaces are kept as typed (text set in place with runs of spaces had moved left), each in
  its own size; a FILENAME field shows the file's current name; a footer is measured with its tabs at their stops, so
  a page no longer comes out a line short under a footer line that just fits; an empty last paragraph is as tall as
  its mark's theme font.
- **Tables**: the top lines set on the hidden lower parts of merged cells count in their rows' heights; a table
  fitted to its text takes the column widths Word saved where the fitting comes within 1.5px of them, and keeps its
  fitted widths when it is redrawn (a document had printed two ways); a cell's own "no fill" clears its table
  style's band shading; a character style that switches bold off doesn't undo the bold of the table style's first
  row, and its size counts as the run's own; a repeated header row prints its text (it had been painted under its
  shading); lines under a pixel wide print at their own width; the lines at a table's page break sit where Word
  draws them, in their own colour and width.
- **Text**: the automatic colour is drawn as automatic, not as the style's colour; the spaces that end a centred or
  right-aligned paragraph take no part in the alignment; a missing font is drawn in the other name its file gives
  it, as Word draws it (pinned with 182 generated documents); an empty paragraph whose line passes the bottom margin
  only by its line spacing's extra stays on its page.
- **A long run of paragraphs kept with the next no longer takes a page each**: a run of paragraphs each set to keep
  with the next, longer than a page, breaks where the page ends, as in Word (a document of 120 such paragraphs came
  out at 113 pages for Word's 14). **A list from a paragraph style takes that style's own indent** over the list
  level's, as in Word (a numbered heading had started 2in to the right and wrapped).
- **AutoCorrect with Track Changes on keeps the caret where you type**: capitalising a sentence's first word left the
  caret after the new capital, so the next words went inside the word ("Typed bolded" came out "Tboldedyped").
- **Section breaks and table lines as Word lays them out** (rules pinned with 265 generated documents printed from
  Word): an empty paragraph that carries a continuous section break takes no room and none of its own spacing, the
  space after the paragraph above and the space before the one below both kept (first in the document or under a
  table it keeps its line; so does one that is all there is of its section); a cell's own bottom line, wider than the
  top line of the row under it, is the line between the two rows and lowers that row, its room taken from a cell's
  top margin where the cell draws a top line of its own; a cell with no top line of its own starts under the widest
  line over its row; an empty paragraph taller than an exact-height row no longer stretches the row.

- **Hardening from the release review**: a document, pasted text or co-editor naming a font "constructor" or
  "__proto__" no longer stops the document opening or breaks the editor; a quoted ">" in a piece of XML kept from a
  part no longer has its namespaces declared twice; the extra layout passes for floating tables placed down past a
  page's foot are capped.
- **Documents with Word 2023's tracked-change dates open again**: a tracked table change carrying `w16du:dateUtc`
  was re-read without its namespace and the document failed to open.
- **Saving an untouched document no longer adds tracked changes**: with Track Changes on, a document with merged-cell
  tables saved them a second time as tracked insertions (a probe the editor measured lines with was read as an edit).
- **Printing and PDF export of long documents is far faster**: the slowest documents tried, which took 65-230s, now
  take 1-13s (a 76-page document of 2,700 table paragraphs went from 53s to 13s); pages of more than one size print
  from copies holding only what each page shows.
- **Headers**: shapes placed against the page or margins are drawn there, out of the header's flow; a header's floats
  don't push the body down; a page with no header of its kind keeps its top margin; all-caps text boxes draw capitals.
- **Floating tables** follow Word's placement: one with no room beside it is set against the margin or laid in the
  flow and broken row by row, one placed down from its text stays a float with the text after it filling the room
  above, the room beside one is measured in its own section's column, and a table that doesn't fit beside a float goes
  below it.
- **Tables and page breaks**: a table kept with the heading above it breaks where it runs out of room; rows that keep
  with the next move on with the paragraph after them; contextual spacing works across a table's edges; an empty
  numbered paragraph in a hidden-mark cell keeps its line; a section's closing paragraph of drawings stays on its page.
- **Font size in table cells**: text in the Normal style takes the size Normal sets, over the table style's and the
  document defaults' (defaults 11pt, Normal 12pt: 12pt in the cells, where the editor drew 11pt and wrapped
  differently). The table style's size applies where Normal sets none. In a document of Word 2010's layout or older, a
  12pt Normal still gives way to a table style that formats its text or paragraphs, as it does in Word.
- **Paragraph spacing**: a paragraph of spaces or no-break spaces takes its mark's line; multiple spacing below one
  line is as short as Word's; a cell's space-only paragraph keeps its table style's size; shadowed text has Word's
  taller line, a shadowed space none.
- **Paragraph spacing, older rule**: in a document that turns off HTML paragraph spacing
  (`doNotUseHTMLParagraphAutoSpacing`) the space after a paragraph and the space before the next are both kept, as in
  Word, where the editor kept the larger. A paragraph of nothing but tabs takes its font's line times its line spacing.
- **Headers**: a bookmark's end between two header paragraphs takes no room (it was drawn as an 18px block, every page
  that much low).
- **A list item moved to the next page inside a split table row** starts at the top of the cell there, its bullet
  beside it, where its text sat a line low.
- **Lists**: auto spacing between items of one list is dropped each side on its own, as in Word, in cells too; two
  numbering definitions over one list are one list; a smaller bullet whose level sets its size is drawn at that size.
- **Fonts**: "Times" is drawn and measured as Times New Roman, as Windows substitutes it; Arial Narrow's line is
  Word's (each line at 1.08 or 1.15 spacing was 0.2pt short, about 10px a page).

## 0.8.4 (2026-11-18)

- **Missing fonts as Word substitutes them**: Palatino in Palatino Linotype, Futura Md BT and Futura Lt BT in Century
  Gothic, Lucida Grande in Segoe UI, Helvetica Neue in Arial, Tms Rmn, Times-Roman, TimesNewRomanPSMT, CG Times (W1)
  and New York in Times New Roman, "Calibri (Body)" in Calibri, Charter BT and Inter in Cambria, Courier in Courier
  New. A face named as a font ("Arial Bold", "Times New Roman Bold") is drawn in that face, and bold text under it is
  no heavier.
- **Hardening from the release review**: repeating sections bound to custom XML are filled within one budget of copied
  content and a section inside another's item is left as it is (a crafted file could multiply them into millions of
  items as it opened); a bound picture shown by many controls is decoded and stored once, a bounded number and size of
  them; a placeholder building block's run formatting is copied into controls only when it is small; a co-editor's
  control properties that aren't text are ignored, and the controls written around a block held to a few levels.
- **More of Word's bound content controls**: an empty node shows the control's placeholder from the document's
  glossary (or Word's default for its kind), and an emptied control shows it again; a repeating section gets an item
  for each node it is bound to, each item editing its own node; a picture control shows the picture its node holds,
  fitted in its box; bound values typed by one co-editor reach the others' controls.
- **A document with a section of columns between one-column text no longer hangs the tab** when a paragraph in the
  columns is taller than the column it moves to and can't be cut there: since 0.8.1 it was moved on and walked again
  without end, and the page never drew. Each such paragraph is now walked again a bounded number of times.
- **A justified, hyphenated document with Word's consecutive-hyphen limit settles once opened**: since 0.8.2 its
  paragraphs were drawn again every 85ms for as long as it stayed open (measuring them for the limit was read as a
  change in the page), using the processor and at times making a click miss.
- **Bold over the Microsoft 365 cloud SemiBold and Demi families** (Aptos SemiBold, Avenir Next LT Pro Demi, Georgia
  Pro Semibold, Verdana Pro Semibold) is laid out at the regular face's widths, as Word lays out its faked bold: lines
  break on the same words as Word's in every browser (Firefox's faked bold had widened each letter, ending bold lines
  a word or two early).
- **open-metric-fonts 1.1.0**: the stand-in fonts load from the 1.1.0 packages (cleaner letters: overlaps and cramped
  or gapped letters fixed; same widths and line spacing), and the bundled Aptos stand-ins are rebuilt from Morrow Sans
  1.1.0. `OMF_VERSION` is 1.1.0; the README's self-hosting steps name the 1.1.0 download.

## 0.8.3 (2026-11-17)

- **Hardening from the release review**: bound controls' lookups are remembered and bounded (a crafted part of
  millions of nodes with thousands of controls bound past them had hung the tab), a co-editor's control properties are
  read in one pass however many unclosed tags they hold, characters XML can't hold are left out of saved text (a pasted
  control character made a file Word wouldn't open), and very deeply nested control content no longer stops a file
  opening.
- **Content controls bound to custom XML show their values** as Word does, instead of their placeholder text, in the
  body, headers and footers and table cells; empty values show blank. Typing into a bound control changes every control
  bound to the same value, the headers' too, and saving (or `fillFields`) writes the new value into the document's
  custom XML part.
- **Lines holding text in different sizes or fonts, or raised and lowered text, are as tall as Word makes them**: each
  run's top and foot from its own font, runs of only spaces or tabs not counted (lines were up to 0.5px off at single
  spacing, 1.5px at double). Headers and footers get the same line-height matching as the body.
- **Shapes placed against their paragraph sit where Word puts them under space before**: the offset counts from the
  paragraph's top above that space (less the space after the paragraph above, from the top margin on a new page), and
  text under a top-and-bottom shape gets the space before again. One-sided floating shapes no longer print 0.8pt low.
- **Lines beside shapes with text on both sides split as in Word**: a line reaching any distance into a shape's room is
  split, offsets keep fractions of a pixel, overlapping shapes are wrapped as one room, and a shape the page layout
  moves gets its lines split again.
- **Compound borders drawn as Word prints them**: paragraph borders of 2.25pt and over (a triple shows three lines, a
  thick-thin pair's lines their own widths), and table cells' compound borders at Word's width, including where lines
  meet, with rows as tall as Word's line widths make them (text in such tables was up to 4.5pt off). The PDF export
  draws the cells' lines too.
- **Finish & Merge takes time in proportion to the merge fields and records**, not their square: 7,400 fields in one
  paragraph from 2.7s a record to 39ms, 300 records with six headers and footers from 26s to 0.3s. The merged documents
  are the same as before, and the scaling test covers Finish & Merge.

## 0.8.2 (2026-11-17)

- **Hardening from the release review**: a co-editor's paragraph borders of any shape are drawn safely (sizes and
  spaces as bounded numbers, colours as hex; a space sent as text had stopped the paragraph being drawn), and a
  paragraph of thousands of shapes with text on both sides looks for text above only a bounded number of them.
- **Exactly spaced lines print where Word puts their text**, the baseline at four fifths of the line's height whatever
  the font; they printed up to 4px high (19px in a 60pt line). A line with a larger run after its first prints on that
  run's ascent, as Word's (10pt text with a 12pt run printed 2-3px high). Documents over 1,500 paragraphs get their line
  heights matched to Word's too (they had come out up to 1px a page short).
- **A paragraph holding a shape with text on both sides splits over two pages** as in Word, the shape staying with its
  first lines (it moved to the next page whole); such a shape reaching past the bottom margin is drawn up to end there,
  one above the top margin down to it. A float drawn up at a page's foot no longer narrows its paragraph's lines on the
  next page.
- **A text box's text, opened for editing, sits exactly where the box draws it**, inset from a thick outline, instead
  of jumping about 12px when editing starts or ends.
- **Thin compound borders** (triple, medium-gap, double and small-gap under 2.25pt) are laid out at Word's width, so
  the text below starts where Word puts it, and their lines drawn as Word prints them (they were up to 22pt off).
- **Long documents with many comments, merge fields (Preview Results), empty paragraphs with formatting marks shown,
  tracked paragraph joins, or automatic hyphenation** open and respond to typing in time growing with the document, not
  its square: 2,400 joined paragraphs from 333ms a key to 14ms, a 2,400-paragraph hyphenated document from 291ms to
  about 1ms. Automatic hyphenation updates only the edited paragraphs.

## 0.8.1 (2026-11-16)

- **Security fixes from the release reviews**: a crafted document could make the editor hang or crawl through work
  without a bound: text measured beside a shape with text on both sides, a WordArt word broken letter by letter, bevel
  sizes and reflection blurs of any size, settings, form fields and note separators read by patterns that went back
  over long runs of unclosed tags, a note of thousands of lines fitted into columns, borders millions of pixels wide,
  columns by the million, paragraphs of thousands of pictures. All are bounded now; font tables are read for their own
  entries only.
- **The arrow keys in tables**: Left and Right go on from cell to cell in right-to-left tables too (at a
  right-to-left cell's end the caret had stayed stuck there), and from the paragraph under or over a table into it,
  not over it; Up and Down in and out of tables keep the caret's column, as in Word, instead of going to the start of
  the cell.
- **Mail merge: NEXT, NEXTIF and SKIPIF**, as in Word's merge: NEXT and NEXTIF move on to the next record within the
  same document (several records to a page), and SKIPIF leaves a document out. Headers and footers are filled from the
  record the document ended on, MERGESEQ counts the merged documents, and Preview Results shows them as Word's preview
  does.
- **Footnotes and endnotes in columns**: a band of footnotes taller than its page stays in the section's columns,
  carried on to the next page's band, as Word lays it out (it fell back to one band the text's width). Endnotes in
  columns, right to left too, go under the text's end in the section's direction, and on a page of their own from the
  top, as Word prints them; the text's columns are no longer cut short for them. Note heights are worked out again only
  from the first band that changed (400 notes in two columns: 0.8s to 0.14s).
- **Lines holding an inline picture are as tall as Word makes them** in every font and line spacing: the picture, the
  font's descent and all of multiple spacing's extra under it; at-least lines the larger of their length and the
  picture, exact lines their length with the picture cut to it, a picture alone on its line no shorter than a line of
  its font. Aptos picture lines were about 0.5pt short each. Text beside a floating picture prints on Word's baseline.
- **Text on both sides of a shape**: the line after the shape's foot is measured from its own top, as Word does; each
  side of a split justified line squeezes by Word's limit for its spaces; a side too narrow for a word stays empty;
  long documents with many such shapes now settle, where their lines had kept changing and words were left under the
  shapes.
- **Borders of two or three lines** (thick-thin, thin-thick-thin, triple, with small, medium or large gaps) take the
  full width Word gives them, so text below them starts where Word puts it. A text box's text is inset from the inside
  of its outline (half the outline's width), as Word does.
- **Accept All and Reject All on long documents are linear** (4 times the changes took 11 times as long), and opening
  one builds its text spans a block at a time (6,400 paragraphs: 0.35s to 6ms). A scaling test of stress documents
  fails on a return to quadratic behaviour, whatever the machine's speed.
- **WordArt wider than the column wraps**, as Word's: its box (not wrapped, resized to fit its text) is as wide as its
  text but never wider than the column, its lines wrapping there (a word longer than that letter by letter), centred
  in it, while typing and when drawn; Word's own files print as in Word (checked against Word's print at 36 and 54pt,
  saved 2in and 9in wide, at the column's edge, moved in and placed against the page).
- **Bevels drawn** (Word 2010's 3-D format on text, w14:props3d: the WordArt gallery's Soft Bevel and Sharp Bevel), in
  the text and in text boxes: the letters lit from above over a rounded rim as wide as the bevel, or, for a sharp
  (angled) bevel, over flat facets of one tone each, meeting in mitred corners and in ridges along narrow strokes, lit
  softly from above and left in the text and harder from the left in text boxes, as Word lights them; an approximation
  of Word's 3-D rendering close to its print. Its extrusion and contour, which don't show face on, aren't drawn, nor in
  PDF export.
- **Reflections keep their descenders and blur**: a reflected g, j, p or y's tail is mirrored too (rising into the
  reflection's gap, as in Word's print), and the copy is softened by its blur as Word's print softens it, on the page
  and in text boxes (descenders in PDF export too).
- **Text beside a float on both sides, justified**: each side of a line split by the shape spreads on its own (the
  words before it to its edge, those after it to the margin), with Word's squeeze on each side; words had been up to
  44pt off Word's. Printed, justified lines beside a footer's shape no longer spread across the full width.
- **A float drawn up and wrapped on its left pushes the lines above it**, as in Word.
- **Shapes' outlines and floating lines drawn whole**: an inline shape's outline is no longer cut at its box, straight
  floating lines (arrowheads and all) are drawn, and a floating shape without text sits at its offset, as Word prints
  it. Drawings behind the text show over the page and under the text (they had been hidden under the page). Outline
  corners are round or mitred as Word draws them. In Word 2010 documents a drawing aligned across the page sits its
  effect extent lower, as Word places it.
- **Footnotes**: a document-wide kerning setting reaches the text and its notes (10pt Calibri notes broke a word later
  than Word's); a long note cut at a page's foot keeps as many lines as Word's; text above a footnote band keeps its
  space after clear of it; a new note keeps the space after its number.
- **Legacy text form fields**: a field whose value runs over several paragraphs is read and filled in (Word writes a
  multi-line value as paragraphs); a field keeps to its maximum length when filled in, typed into or pasted into (a
  paste too long is cut to fit).
- **Pasting** a lone heading into an empty paragraph keeps the heading's style (it took the empty paragraph's).
- **Deleting a section break** (Delete or Backspace at it, or a selection across it) takes its section's setup away
  and joins the text before it to the following section, in that section's columns and page setup, as Word does
  (columns applied to selected text, the break before or after them deleted). Delete at the end of a paragraph before
  an empty one that ends a section keeps the break, as in Word. The saved document opens in Word with its sections.
- **Hyphenation zone and consecutive-hyphen limit**: with automatic hyphenation on, a word is left whole where wrapping
  it leaves no more room at the line's end than the hyphenation zone (in documents of Word 2010's layout and older:
  Word 2013's has none), and no more lines in a row end in a hyphen than the document allows, on screen and in print.
  Checked against Word's print of test documents (a 1in zone, limits of 1 and 2); in Firefox, which hyphenates by
  itself, too.
- **Accept All and Reject All on thousands of changes are faster again**: on 6,000 tracked changes, from 1.9s to
  0.5-0.6s each, the result the same.

## 0.8.0 (2026-11-16)

- **Security fixes from the release review**: a crafted document could put CSS of its own into a picture's style (a
  request to another site on open, a picture over the page), make EMF/EMF+ pictures take minutes and gigabytes (a
  bitmap claiming any size, metafiles nested without end, an object drawn again and again, long text wrapped), or make
  a mail-merge comparison backtrack for seconds; a co-editor could do the same through a picture's attributes, write
  parts outside the package's folders with a "../" name, or loop an editor for ever with a name without an extension.
  A cross-reference to a list label of thousands of dots or spaces took time growing with its square, and a field's format switch naming "constructor" was taken as a format. All are bounded or checked now. Two slow patterns over table-of-contents switches and pasted styles were made linear,
  as were the reads of a tag's attributes in notes, pictures, shapes, page fields and text effects (seconds each on a
  crafted tag); a chart's embedded workbook no longer puts its own sheet id into a regular expression.
- **Endnotes move with the text**: when columns or the text above them changed, endnotes stayed where they were first
  drawn (in a one-section document of columns, at the page's foot). A footnote mark just after a comment's end is
  printed now.
- **Mail merge: IF fields per record, and merge fields in headers and footers**, in the merged document and the preview,
  as Word's own merge fills them.
- **Tables of contents**: `\u` alone lists paragraphs by outline level; `\t` takes entries from named styles at their
  levels; the field keeps its switches when updated.
- **Pasting**: headings from styled paragraphs (Word's own clipboard HTML, `class="Heading1"`) arrive as headings; a
  paragraph pasted into an empty one keeps its style.
- **Home and End on a table of contents entry** (or any paragraph beginning or ending in a linked field) go to the start
  or end of the paragraph; text typed straight after them is no longer misplaced. The find bar keeps its search.
- **Shapes turned and flipped inside groups**, a turned group's shapes turning with it, as Word draws them; SmartArt
  gradient fills; text in a shape centred or at the bottom as its anchor says.
- **Inner shadows drawn in body text**: Word's inner shadow (a text shadow scaled to nothing, the letters looking
  pressed in, as in the WordArt gallery's Style 15) is drawn on the page and in print; text boxes' version now falls on
  Word's side at Word's depth.
- **Glows in PDF export**, close to Word's print (with an installed or embedded font, as reflections); and a run mixing
  alphabets its font lacks (Greek and Hebrew, say) gets the Noto fallback for each, not "?" for all but the first.
- **Shapes with text on both sides draw their outline whole**, past their box as Word does, the text beside them where
  it was.
- **A float placed against the page that pushes its own paragraph to the next page goes on with it**: the paragraph
  and its picture start the next page, as Word lays it out (the picture had been left on the page before).
- **Body text wraps round shapes in headers and footers**: a square-, tight- or through-wrapped shape anchored in a
  header or footer pushes the body text aside on every page showing it, as Word lays it out (it was drawn over the
  text). Checked against Word's print of seven test documents; on the way, a header shape placed below its paragraph
  no longer pushes the header's own text aside or moves the body down.
- **Reflections drawn**: a run with Word 2010's reflection text effect is drawn on the page, in print and in PDF
  export, each word mirrored below its baseline with the gap and fade of Word's print (in PDF export when an installed
  or embedded font is used; with only the standard PDF fonts none is drawn).
- **Co-editing shares styles and numbering** added after opening (a new style, a new list's numbering, a table of
  contents' styles): the other editors in the room, and anyone who joins later, draw the text as its author does, and
  their saved files hold them. They had arrived only when the file was reopened.
- **Text boxes measured before their font arrived are drawn again**: text in shapes, text boxes and WordArt wrapped at
  a fallback font's widths when the document was read before the stand-in font loaded; once it has, those shapes are
  laid out again (on the page only: no undo step, no change to the document).
- **Accept All and Reject All on thousands of changes are faster** (6,000 changes: 3.5 to 2.2 seconds, the relayout
  included), one undo step as before.
- **Long documents with many tracked changes open fast**: 6,000 tracked changes over 177 pages opened in 24 seconds,
  now 1.5; a 118-page document in 0.45 seconds rather than 3.2. The margin's balloons and the print layout read
  positions inside paragraphs the browser skips laying out off screen, which laid each out on the spot; those far off
  screen now go by their paragraph (the balloons settle on their lines as the paragraph comes on screen). Measured on
  stress documents: typing 33-65ms a keystroke, Find's 20,000 matches in 60ms, Accept All on 6,000 changes 3 seconds.
- **PDF export of Hebrew, Arabic, Greek and Cyrillic without the real font** (Firefox, Safari, or a font the computer
  lacks): the characters a run's font lacks are set in the Noto font for their alphabet, fetched when first needed
  (`fallbackFonts`: npm's `@fontsource` packages on jsDelivr by default, your own copy, or `false`), the run's other
  characters in its own font beside them; they had come out as "?". Arabic letters are joined and right-to-left runs
  ordered as on screen.
- **Floating shapes' outlines and arrowheads drawn whole**: a thick outline or an arrow's heads past the shape's box were
  cut off at it (only shapes in line with the text had the room); now drawn past it as Word does, the shape still
  wrapping the text as before.
- **Long footnotes laid out as Word lays them out** (checked against Word's print of four test documents): the lines
  above a long note's reference stay on the page, the reference's line going on with its note; a reference's line stays
  when its note's first two lines fit under it, the rest of the note going on on the next page; a note longer than its
  page starts under the text and is cut at a line's end, not partway through one; notes on a landscape page among
  portrait ones are measured at that page's width. Printing: footnote reference numbers at a line's end print again,
  and a landscape page among portrait ones prints whole (it was cut off at the portrait width).
- **Delete Cells (shift cells left)** on the Table tab: the cells go and the cells after them move left, as Word's
  dialog does; with changes tracked, each is recorded as a deleted cell (w:cellDel), as is each cell of a column
  deleted, and goes when accepted (Word reads it as a cell deletion and accepts it the same).
- **Pictures placed against the page**: the text wraps on the side with room (one near the right margin had its text
  squeezed into a sliver on its right, a letter a line).
- **A line holding a picture** under 1.5 or double spacing has all of the spacing's extra below it, as in Word (the
  next line had been 3pt and 7pt too high).
- **Security fixes from code scanning.** A crafted document could make reading its page margins or a merge field's
  number picture take a very long time (regular expressions that backtracked): both are now read in linear time, as is
  Find's replacement text. Tracked rows' author names decode `&amp;` last (an author named with an escaped quote was
  read as another); the agent's markdown tables escape backslashes before pipes. The CI workflow asks only to read the
  repository.
- **Text in shapes, text boxes, WordArt and SmartArt in the stand-in fonts.** These are drawn as pictures, which can't
  use the page's web fonts, so where the computer lacked the document's font their text fell back to an installed one
  (Aptos's text in Calibri). The editor now shows each with the stand-in faces its text needs embedded (fetched once
  from where the page loads them; the document itself is unchanged), and measures their lines in the same font: WordArt
  in Aptos is now as wide as in Word's print.
- **Turned and flipped pictures.** A picture's turn and flips (Word's Rotate and Flip) are drawn as Word draws them, in
  line with the text or floating, a quarter-turned one taking its turned box's room in the line; the rotate handle and
  Alt+Left/Right turn pictures as they turn shapes, the turn and the room round it written as Word writes them (checked
  against Word's print of pictures turned by Word and by the editor).
- **Stand-ins for 151 more Microsoft fonts**, from open-metric-fonts (open fonts with the originals' character widths
  and line spacing): Grandview, Seaford, Tenorite, Skeena, Bierstadt, Avenir Next LT Pro, Aptos's other weights, Aptos
  Narrow, Serif and Mono, Bahnschrift, Sitka, Century Gothic, Garamond, Franklin Gothic, the Lucida family and the
  rest of the Office and Windows set. Where the computer lacks one, the editor draws it in its stand-in and lays its
  lines out at the original's line height, so they break where Word breaks them. Loaded only for fonts a document uses,
  one file per alphabet its text needs, from npm's `@open-metric-fonts` packages on jsDelivr, or from your own copy
  (`fonts.omf`: the release's web bundle unzipped, or the npm packages you want; README "Self-contained hosting"),
  or not at all (`fonts.omf: false`). `OMF_VERSION` says which release the editor's map is for.
- **The Aptos stand-in, now from Morrow Sans.** The Aptos and Aptos Display stand-ins bundled with the editor
  are now made from open-metric-fonts' Morrow Sans and Morrow Sans Display (fonts drawn to Aptos's widths), each letter
  at the width it had before (fitted to Word's prints): lines break and sit exactly where they did.
- **Self-contained hosting, documented and tested**: README "Self-contained hosting" shows how to serve the stand-in fonts
  from your own site (no CDN) and the Content-Security-Policy the editor needs (only `'self'`, with `data:`/`blob:`
  images and generated styles); a test runs the editor under that policy. Fixed on the way: the ribbon's drop-downs and
  colour pickers also wrote their handlers into the page as inline-script attributes (inert, but refused under a strict
  policy).
- **WordArt.** Insert > WordArt puts the selected text, or "Your text here", in a floating text box as wide as its
  text, in one of Word's fifteen gallery styles (written exactly as Word writes each, in the document's theme colours),
  open for typing; the box widens to fit what's typed. Word's own WordArt prints as in Word, its box as wide as its text
  (checked against Word's print of all fifteen).
- **Word 2010's text effects drawn**: a run's coloured, transparent, gradient or no fill, its outline, shadow, glow and
  reflection (w14:textFill, textOutline, shadow, glow, reflection), in the text and in text boxes, as Word draws them,
  with the extra room Word gives a line for a glow or reflection (within 0.1pt of Word's prints at 18 to 72pt). They
  were kept and saved before, but not shown.
- **Drawing shapes, as in Word.** A shape picked from Insert > Shapes with the mouse is drawn by dragging on the page
  (Shift: a square, a circle, or a line at 15-degree steps; Ctrl: from its centre out; a click: at Word's size; Esc
  stops), floating in front of the text where it was drawn and anchored to the paragraph beside it, as Word's are (one
  Undo each; checked by opening the saved file in Word, whose print has them where the editor drew them). Picked from
  the keyboard, a shape still goes in line with the text at the caret.
- **EMF+ pictures drawn.** A picture saved as EMF+ (GDI+'s metafile records: charts and drawings copied from Office
  and other programs) is now drawn from those records: fills, transparency, gradient, hatch and picture brushes, pens
  with their widths, dashes, caps and joins, paths, curves, arcs and pies, text, bitmaps, clipping and the world, page
  and container transforms. An EMF+ only file (no GDI copy to fall back on) was blank; a dual file now shows what its
  GDI copy lacks, such as transparency and Visio's colours. Checked against Word's print of pictures drawn by Windows'
  own GDI+. A dual file whose EMF+ records need something not read yet is drawn from its GDI copy, as before. Text
  wraps in its layout box and is cut off at it, path gradients take each corner's colour, hatches are drawn in the
  picture's pixels, and pens keep a different cap at each end, as GDI+ draws them.
- **Footnotes in newspaper columns, as Word lays them out.** A page's footnotes are one band across its foot, every
  column on the page cut at its top, the notes in the section's columns, balanced as Word balances them (the first
  column taking more, a note split between columns only where each part keeps two lines), right to left in a
  right-to-left section; `w15:footnoteColumns` gives them a number of columns of their own (1: one band the text's
  width). Each note is on the page of its reference, also when the reference is in lines carried to the next column
  (they went on the wrong page, or were lost, in any section of columns). The separator is 2in long, as Word's, at
  the right when its paragraph is right to left. Checked against Word's prints of 13 documents.
- **`ribbon: 'simple'`**: a page can start the editor with the one-line simplified ribbon (`'classic'`: the full one)
  instead of the user's last choice in the browser; the ribbon's own button still switches, and `editor.setRibbon()`
  and `editor.ribbon` change and read it. Also on the React, Vue and Angular components, and the demo's `?ribbon=`.
- **Footnotes in columns of older documents** (compatibility mode under 15, Word 2010 and before): each column has
  its own note area and separator under it, holding the notes referenced in it, and only that column is cut at it,
  as Word lays them out for those documents (checked against Word's print of a compat-14 document).
- **Balanced columns as Word balances them**: a section of columns before a continuous break is cut no lower than it
  must be for its last column to fit, as Word does (a paragraph keeping two lines a side had moved whole to the last
  column and overfilled it), and what follows starts a line under the lowest column's last line, not half a leading
  lower.
- **Column breaks in a section of columns** among other sections are acted on (they were ignored there): the text
  after one goes on in the next column, a page break in such a section on the next page, and a column break before a
  next-page section break keeps its own column. A section of columns ending at a continuous break with a column break
  of its own isn't balanced, as in Word, and what follows starts under its lowest column.
- **A table longer than a column** goes on over as many columns as it takes, cut at a row at each column's foot, its
  header rows again above each part, as in Word (it was cut once, its rest running down onto the next page).
- **A paragraph longer than a column** goes on over as many columns as it takes, cut at each one's foot with two
  lines kept on each side, as in Word (it was carried once, its rest running down through the margin onto the next
  page), in a document of columns or a section of columns among others; its line numbers and footnotes follow it
  into each column. Endnotes in a document of columns are balanced in
  its columns as Word balances them, after the text's last columns, the first taking the odd line.
- **Columns hold Word's lines.** A paragraph going on in the next column was measured a little too tall at the
  column's foot, so a column could hold a line fewer than Word's, and its lines in the next column sat about 2pt high;
  both now as Word lays them out (two and three columns, tables, right to left, unequal widths).
- **Section breaks after columns, and odd and even pages.** A next-page break from one section of columns into
  another starts the new section on a new page (it began under the columns); a page filled to its foot keeps its
  section's empty last paragraph, so the next section doesn't start a page late; and an odd- or even-page break adds
  the blank page Word adds to reach the right side. Checked against Word's prints of 20 documents.
- **Line numbers in columns, as Word numbers them.** A paragraph's lines carried to the next column are numbered there,
  in reading order (they were numbered where they're cut off, in the bottom margin, and the next column's count was
  off from there); a right-to-left section's numbers go right of each column, as in Word; and a section that restarts
  its numbers each page goes on counting down the page it starts on after a continuous break (it started again at 1).
  Checked against Word's prints of two and three columns, unequal columns, a table, a right-to-left section and a
  section change on the page.
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
