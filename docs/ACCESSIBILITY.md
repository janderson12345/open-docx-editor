# Accessibility conformance (DRAFT): Section 508 / WCAG 2.1 AA

Scope: the editor's own interface (ribbon, menus, right-click menu, Find, comments, form fields) as built in
this repository. The content of a document is its author's responsibility; Review > Check Accessibility reports on it.

**Method.** Source and rendered-page testing, plus automated NVDA runs on Microsoft Edge, Windows: NVDA 2026.2
(Guidepup, speech over NVDA Remote) for runs 1-5, NVDA 2026.3 portable for run 6 (real key input through the OS, so
NVDA's own keyboard hook sees it; speech read from NVDA's input/output log). No human listener, JAWS, Narrator, voice
control or magnifier yet.

NVDA run 5 (2026-10-23): ribbon tabs, Tab through ribbon buttons, Esc/F6, Bold toggle ("toggle button, not pressed" →
"pressed"), the right-click menu and its items all spoke correctly.

NVDA run 6 (2026-11-06; a harness kept outside the repository, on rich.docx and form.docx). Heard as expected: headings ("heading,
level 1, Introduction"), the table ("table with 3 rows and 3 columns, row 1, column 1, A1", a merged cell "row 2 through
3"), a tracked insertion ("added"), a picture's alt text ("graphic, A blue box"), the ribbon tabs ("Insert, tab, selected,
3 of 7"), buttons with their shortcuts, Esc back to "Document, edit, multi line", the File menu's first item ("File menu,
New, 1 of 6"), the right-click menu's items, the comment alerts ("Comment added. Type it in the comment box", "Comment
posted"), Review > Accept ("Accept this change, 1 of 2", "Change accepted"), Ctrl+B/I/U ("Bold on", "Bold off"), and in
forms mode every field on Tab, Shift+Tab and Ctrl+Home ("Agree, check box, not checked", "Color, drop-down list, Red",
"Due date, date picker, empty"). Fixed by this run: headings had no heading role; formatting shortcuts said nothing;
focus fell to the page after posting a comment and after choosing from a ribbon menu, and for a moment while a comment
box opened; arrowing through a drop-down field picked the first item moved to and closed it (Enter then reopened it).
A check box content control now has a checkbox role with its state ("check box, not checked", arrowed to outside forms
mode; it had been read by its glyph). Open: list items are read as their label and text ("●, Bullet one", "1., Step
one"), with no "list" or item count. Tried: giving a bullet's label "bullet" as CSS alternative text (`content: ... /
"bullet"`): Edge and NVDA then read nothing for it, so the glyph stays. And `role="listitem"` with `aria-level` on list
paragraphs: Edge drops a list item that has no list around it (the accessibility tree still says "paragraph"), and the
editing model keeps every paragraph directly in the document, so there is no list element to give them; reparenting
them with `aria-owns` would take them out of the text box NVDA reads the caret in.

The rule used for every row: claim *Supports* only with evidence
(a file, a test); anything untested that could plausibly fail is *Partially Supports*, with the test that would settle it.

Evidence runs on every `npm run test:e2e`:
- `test/a11y.mjs`: axe-core (WCAG 2.0/2.1 A and AA rules) on every ribbon tab, the File menu and the right-click
  menu, in light and dark mode; plus keyboard checks (F6 to the ribbon, Esc back to the document).
- `test/a11y-scan.js`: pattern scan for what axe doesn't cover (visible label missing from the accessible name,
  targets under 24 px, unnamed widgets, broken id references, positive tabindex, mouse-only elements, unhidden icons).
- Floors make the checks fail if they stop finding the controls. The label-in-name check was broken on purpose once
  and seen to fail before the fix was restored.

| Criterion | Level | Remark | Evidence | What would close the gap |
|---|---|---|---|---|
| 1.1.1 Non-text content | Supports | Icons are hidden from screen readers; every button has a text name. | `src/editor/toolbar.js` (ICON_PATHS, `aria-label`), `test/a11y-scan.js` | |
| 1.3.1 Info and relationships | Partially Supports | Ribbon tabs, menus and toolbars carry ARIA roles. In the document, NVDA announces headings (with their level), tables (rows, columns, merged cells), tracked insertions and pictures' alt text (NVDA run 6). List items are read as label and text, without list semantics. | `test/a11y.mjs`, NVDA run 6 | List semantics in the editable area; a JAWS pass. |
| 1.4.3 Contrast (minimum) | Supports | Interface text passes in light and dark mode. Document text keeps the author's colors. | `test/a11y.mjs` (axe color-contrast, both schemes) | |
| 1.4.4 Resize text / 1.4.10 Reflow | Supports | At 700 and 320 CSS px wide (200% and 400% zoom) the ribbon stays one row, with groups that don't fit behind its More button (which brings them all into view), the tabs wrap, and nothing but the document sheet (a two-dimensional layout, exempt) scrolls sideways. | `test/a11y-more.mjs` | |
| 1.4.11 Non-text contrast | Supports | Ribbon icons and form-field borders measure at least 3:1 against what's behind them, light and dark (field borders were 1.46:1 and were darkened); focus outline is solid blue (#2b6cdf). | `test/a11y-more.mjs`, `src/editor/editor.css` | |
| 1.4.12 Text spacing | Supports | With WCAG's text-spacing values applied to the interface, no button, tab or group label's text is cut off (style-gallery tiles are previews of the document's styles and are cut on purpose). | `test/a11y-more.mjs` | |
| 2.1.1 Keyboard | Supports | Every ribbon control, menu, the right-click menu (Shift+F10) and form fields work from the keyboard. The page area is focusable for scrolling. | `src/editor/toolbar.js` (`el` keydown), `src/editor/fields.js` (Tab, Space), `test/a11y.mjs` | |
| 2.1.2 No keyboard trap | Supports | Esc closes menus and returns to their button; F6 and Esc leave the ribbon. | `test/a11y.mjs` | |
| 2.1.4 Character key shortcuts | Supports | All shortcuts use Ctrl, Alt, Shift or function keys. | `src/editor/editor.js` keymap | |
| 2.4.3 Focus order | Partially Supports | Tab order follows the ribbon layout; with NVDA (run 6), F6/Esc, menus, comments and form fields return focus to the document or the next control. | NVDA run 6 | A JAWS pass; a human review. |
| 2.4.7 Focus visible | Supports | `:focus-visible` outlines on all controls. | `src/editor/editor.css` | |
| 2.5.3 Label in name | Supports | Accessible names contain the visible label. | `test/a11y-scan.js`, `test/a11y.mjs` | |
| 2.5.8 Target size (WCAG 2.2) | Supports | Ribbon targets are at least 24 px. | `src/editor/editor.css`, `test/a11y-scan.js` | |
| 3.1.1 Language of page | Supports | The editor takes the document's language from the .docx; the host page sets its own. | `src/ooxml/docx.js` (`lang`), `src/editor/editor.js` | |
| 3.1.2 Language of parts | Supports | Text the document marks as another language (w:lang) carries that language, so screen readers voice it and the browser's spell checker uses its dictionary; text marked not to be checked (w:noProof) isn't. | `src/editor/schema.js` (rextra) | |
| 4.1.2 Name, role, value | Partially Supports | Tabs have `aria-selected`, menu buttons `aria-expanded`, toggle buttons (Bold, alignment, Ruler…) `aria-pressed`, confirmed in NVDA. Form fields' name, kind and state are announced in forms mode (heard in NVDA run 6). | `src/editor/toolbar.js`, `test/a11y.mjs` | |
| 4.1.3 Status messages | Partially Supports | Find results and notices use `role=status`. Accepting or rejecting changes, adding, posting, replying to, resolving and deleting comments, and Ctrl+B/I/U are announced through a live region; heard in NVDA run 6 (comment added and posted, change accepted, bold on/off). | `src/editor/toolbar.js`, `src/editor/live.js`, `src/editor/track.js`, `src/editor/comments.js`, NVDA run 6 | Reply, resolve and delete heard in NVDA. |

## Not done yet
- Screen readers (NVDA + Chrome or Firefox, JAWS + Chrome, Narrator + Edge), magnifier, Windows High Contrast, voice control.
- 200% zoom, reflow, text spacing.
- A full VPAT 2.5 Rev 508 document built from these rows.
