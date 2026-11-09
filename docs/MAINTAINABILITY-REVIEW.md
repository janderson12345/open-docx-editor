# Maintainability review (2026-11-08)

What would slow down a developer taking this code over, ranked by how much it costs them, with the fix and its risk.
Measured with a script kept outside the repository (sizes, function lengths, nesting, CSS, inline styles, duplicate helpers) and read
by hand. Correctness was reviewed separately (two code reviews, 20 findings, 19 fixed).

**Scale:** 14,400 lines of source (870 KB) in 60 files under `src/`; 55 unit tests, 9 browser suites, a 151-document
layout comparison against Word and a 19-document corpus.

## 1. Private document names in the repository (fix first)

Code comments and the layout notes name the private sample documents the layout was fitted on: about 25 code comments
and well over 100 lines of `docs/LAYOUT-WORK.md` cite them by file name, among them a personal résumé and client and
company names. A new developer can't open those documents, and the names shouldn't be published with an MIT
repository.

- **Fix:** replace each name with a neutral label (`sample A`, `corpus 2: c2-07`) or with the property it shows ("a
  12pt Symbol bullet over 11pt text"); keep a private mapping outside the repo. Rewrite `docs/LAYOUT-WORK.md` the same
  way, or move it out of the repository as a private log and keep a short public summary of the rules it found.
- **Risk:** none to behaviour (comments and docs only). The history still holds the old text: if the repository is ever
  published, start it from a fresh history or rewrite it.

## 2. Very large functions

| Function | Lines | What it holds |
|---|---|---|
| `createPages` (`editor/pages.js`) | 1,605 | the whole page view: measuring, splitting paragraphs and tables, columns, footnotes, headers and footers, print, PDF hooks, as one closure over shared `let` state |
| `layout` (inside `createPages`) | 508 | one pass of pagination |
| `createToolbar` (`editor/toolbar.js`) | 783 | every ribbon tab, menu, the context menu and their commands |
| paragraph styling (`editor/editor.js`, `compute` / `visit`) | 260 | every per-paragraph rule (lists, bullets, borders, contextual spacing, tracked marks, headings) |
| `draw` (`pdf.js`) | 280 | the PDF export |

A developer has to hold hundreds of lines of shared closure state in their head to change one rule. `pages.js` nests
eight levels deep.

- **Fix:** split along the seams that already exist as inner functions: `createPages` into modules for measurement
  (`paragraphLines`, `lineStart`), splitting (`splitParagraph`, `splitTable`, `splitRow`), columns (`colSplit`,
  `colSplitTable`, `colRegions`), notes, headers and footers (`drawSheets`, `pageAt`), and print (`beforePrint`), with the
  shared state (`margins`, `pages`, `gap`, caches) passed as one explicit object. `createToolbar` into one file per
  ribbon tab plus the menu and context-menu helpers. The paragraph pass into small named rule functions called in order.
- **Risk:** high for `createPages` (pagination is what the 99% layout match rests on): one seam per commit, each run
  through all tests and the 151-document comparison, which must not move. Low for the toolbar (the browser suites
  cover it).

## 3. Repeated helpers and magic numbers

- The same small helpers are written in several files: `color` (shapes, pdf), `roman` (styles, pdf), the `SKIP`
  selector (printlines, pdf), the Office colour palette (chart, shapes), child-element lookups `els` / `child` (chartedit,
  shapes, omml), `esc` (emf, xml), `onOff` (writer, xml), `nextId` / `now` (comments, track), `idOf` (comments, notes).
- Unit conversions are inline numbers: `9525` (EMU per px) 23 times, `/ 15` (twips per px) 30 times, `* 4 / 3` and
  `0.75` (pt and px) throughout.
- **Fix:** one `ooxml/units.js` (`emuToPx`, `pxToEmu`, `twipsToPx`, `ptToPx`...) and one place for each helper,
  imported where needed.
- **Risk:** low; mechanical, covered by the tests and the layout comparison.

## 4. CSS

`editor/editor.css` is 685 lines but 110 KB: 582 rules, 271 lines over 160 characters, one selector of 582 characters.

- **76 `!important`**, most to win over styles the editor writes inline from JavaScript (display, height, background,
  padding). They make the cascade hard to follow: a new rule can lose without any visible reason.
- **146 hard-coded colours** beside 96 uses of `var(--odx-…)`; the ribbon blue `#2b6cdf` appears 24 times, `#fff` 33.
  The `--odx-*` colour variables are never defined, only used with fallbacks, and share their namespace with internal
  layout variables (`--odx-lh`, `--odx-pt`, `--odx-bl`), so a host page can't tell which ones are meant for theming.
- **42 selectors are set in more than one rule** (`.odx-content` in six), so a property's value has to be hunted down.
- **Inline styles from JavaScript:** `pages.js` sets styles in 41 places, the toolbar in 13. Positions and sizes that
  depend on measurement belong there; fixed looks (colours, borders, display) belong in classes.
- **Fix:** split the file by area (ribbon, document, page and print, comments and balloons, dialogs), define the theme
  tokens once on the editor's root (`--odx-color-*`, documented as the theming API) apart from layout variables
  (`--odx-l-*`), replace repeated literals with them, merge duplicate selectors, turn fixed inline looks into classes,
  then remove the `!important`s that no longer have anything to win against.
- **Risk:** medium: the document's own look is measured against Word, so changes to `.odx-content` rules go through the
  layout comparison; ribbon and dialog changes through the accessibility suites (contrast, focus, reflow).

## 5. Comments and naming

- Comments are thorough about why (what Word does, which document showed it), but long parenthetical asides often
  carry history ("tried at run N, reverted") a newcomer doesn't need, and refer to private documents (item 1).
- Some names only make sense with history: `odx-pl`, `lsize`, `lalt`, `bla`, `fx`, `runmult`, `TIE`, `inset`, `lead0`.
- **Fix:** keep the why, drop the history (it's in `LAYOUT-WORK.md` and git), and give short names a one-line
  definition where they're declared.
- **Risk:** none to behaviour.

## 6. Documentation for the next developer

- `docs/ARCHITECTURE.md` (113 lines) doesn't mention most editor modules: `baseline`, `lhfix`, `widths`, `justify`,
  `joins`, `printlines`, `hfedit`, `txbxedit`, `tablefit`, `imgresize`, `charts`, `chartedit`, `omml`.
- `CONTRIBUTING.md` (48 lines) explains the tests, but not the layout loop: how to print the samples with Word, run the
  151-document comparison and read its report, and what "within 1px" means.
- The measurement tools the layout work depends on live outside the repository (`odx-test/`): a new developer can
  run the unit and browser tests, but not the Word comparison.
- No linter or formatter config: style is consistent today by hand.
- **Fix:** a module map in `ARCHITECTURE.md` (what each file does, how a keystroke becomes pages and a printed PDF);
  the layout loop in `CONTRIBUTING.md`; the comparison tools moved into `tools/word-compare` with instructions (the
  documents themselves stay private); an ESLint and Prettier config that matches the current style.
- **Risk:** none to behaviour.

## What's in good shape

- No dead code to speak of (3 unused exports), no TODO or debugging leftovers.
- Clear layering: `ooxml/` reads and writes the file and knows nothing of the editor; `editor/` is ProseMirror;
  `engine/` is the print layer; `index.js` is the public API, with type declarations.
- Tests at every level: unit tests in Node, browser suites for editing, accessibility, forms, charts, text boxes,
  co-editing and PDF, and the layout comparison against Word.

## Suggested order

1. Private names out of the code and docs (item 1).
2. Documentation and lint config (item 6): no behaviour change, and it helps every later step.
3. Shared helpers and units (item 3).
4. CSS (item 4).
5. Comments and naming (item 5), alongside the above.
6. Splitting the large functions (item 2), toolbar first, `createPages` last, one seam per commit.

## Status after the clean-up (2026-11-08)

Every step went through the unit tests, the nine browser suites and the 151-document comparison, which stayed at
35084/35432 lines within 1px, 151/151 page counts; the 19-document corpus stayed at 2406/2462.

| Item | State |
|---|---|
| 1. Private names | Done: comments and docs use `sample S###` labels; the run log is out of the repository, its rules in `docs/LAYOUT.md`. The old names are still in the git history. |
| 2. Large functions | Done. `createToolbar` 783 → 561 lines (`ribbonkit.js`, `findbar.js`, `ctxmenu.js`, page setup as pure functions in `ooxml/sectpr.js` with unit tests). The paragraph pass's `visit` 260 → 80 (`pararules.js`). PDF `draw` 280 → 135 (one function per layer; the test PDFs render pixel-identical). `createPages`: header and footer drawing, their fields, row sizing, print copies and section reading are their own modules; `layout()` 508 → 259 lines, its phases named functions (`measureBlocks`, `assignNotes`, `startColumns`/`endColumns`, `placeColumns`, `applyPushes`). Still one closure: the walk and the splitting functions share `createPages`'s state (margins, notes, caches), which is the page view's model; making it an object of its own is the next step if the file grows. |
| 3. Helpers and units | Done: `ooxml/units.js`, one copy of each helper (the w:cols reading, once three, is `sections.js columnsIn`; "deleted whole" is one function for body and headers). |
| 4. CSS | Done as far as the tests can prove it. Sections by area, print last; the theme colours declared once on `:where(:root)` and listed at the top apart from the layout variables; the ribbon blue is `--odx-accent`; 11 rules set twice merged, 2 wholly overridden rules removed; 13 `!important` replaced by specificity where every competing rule (hover and pressed included) could be checked; the field pickers' and mini toolbar's fixed looks moved from inline styles into classes. Each step checked element by element with computed-style snapshots (two shades unified on purpose). Then 15 more in the ribbon, panels, comments and Reading view, checked on 18 snapshotted states (read-only, forms, comments on and off, composing, No/All Markup, Reading view, simplified and classic ribbon, open menus, the context menu, Show/Hide, hovered controls); the `odx-split` rule, which never applied (`menu()` ignored its class) and would have made three buttons smaller than the 24px target size, removed. Left on purpose: `[hidden]`; rules that override the document's own style CSS (style tiles) or styles the layout writes inline (pages, print, tracked-change views), whose removal would need the layout comparison for each; the focus ring. The remaining inline styles are geometry the layout measures. |
| 5. Comments and naming | Done: run numbers and tool paths out of comments; short CSS variable names defined in the glossary. |
| 6. Documentation | Done: module map and the keystroke-to-pages path in `ARCHITECTURE.md`, the layout loop in `CONTRIBUTING.md`, the comparison tools in `tools/word-compare`, ESLint in CI. |

Found along the way: Esc didn't close a context menu opened with the mouse (fixed); page setup and Find had no browser
test (added to the e2e suite).
