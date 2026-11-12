# Generated documents compared with Word

A test tool, run by hand: it is not part of `npm test`, the browser tests or CI. It generates documents from
combinations of features, has Word print each one to PDF, prints each one through the editor, and reports which
features and feature combinations lay out differently from Word, worst first.

Windows, with Word and Microsoft Edge installed; Node 22.13 or later, Python 3 and PowerShell 7 (as
`tools/word-compare`, whose printing and line matching it uses).

## The documents

```powershell
node tools/doc-generator/generate.mjs <work folder>          # every feature alone, every pair, 300 random mixes
node tools/doc-generator/generate.mjs --list                 # the features
```

`features.mjs` defines the features: run formatting (bold, fonts, sizes, effects...), things inserted in a paragraph
(tabs, fields, notes, comments, pictures, tracked changes...), paragraph formatting (alignment, indents, spacing,
styles, tab stops...), structures (lists, eight kinds of table, a drop cap), where the content goes (a text box, a
frame, a footnote or endnote, the header or footer) and the page (sizes, margins, columns, borders, line numbers,
headers and footers, section breaks) and document (compatibility mode, hyphenation, default font, length).

Every document has the same shape: a title line naming its features, four content paragraphs the features apply to,
and plain paragraphs after them. Features in one group are alternatives and never combined; a feature that can't go
in a header, a note or a text box isn't combined with one that puts the content there.

- `s-<feature>`: one feature alone (and `s-plain`, none).
- `p-<a>__<b>`: every pair of features that can go together.
- `r-<seed>`: three to six features picked at random from the seed (`--random N --seed S`: seeds S to S+N-1).

All of it is reproducible: a document's text comes from a seed (from its name, or its own seed), so the same command
makes the same file. To make one again from a report:

```powershell
node tools/doc-generator/generate.mjs <folder> --features table-merged,textbox
node tools/doc-generator/generate.mjs <folder> --seed 117 --random 1
```

## A run

```powershell
node tools/doc-generator/run.mjs <work folder> [--steps word,ours,score,report] [--run name]
```

- `word`: one Word, started for the run, prints every document to PDF (`word-pdf.ps1`): opened read-only and
  invisibly, closed without saving, no Word setting changed. A document that takes over two minutes is abandoned:
  the run ends the Word it started (never another) and carries on with a new one. Word's PDFs are kept: run it once.
- `ours`: builds the demo page and prints every document through it, in Edge (`print-fast.mjs`: as
  `tools/word-compare/print-ours.mjs` prints, but each browser page reused and each document printed as soon as the
  editor has laid it out, 16 at a time: the full set in about 15 minutes).
- `score`: extracts the lines of both (`lines.cjs`: `tools/word-compare`'s, but text on one baseline split
  at gaps over 12pt, so columns side by side are lines of their own) and scores each document (`score.py`).
- `report`: `report-<run>.md` (and `.json`): features worst first, the pairs worse together than either alone, the
  documents worst first with how to make each again, and anything Word or the editor couldn't print.

Steps pick up where they stopped. `word` and `ours` can run at the same time (two commands with `--steps word` and
`--steps ours`). Word prints about five documents a second, the full set (about 7,500) in under half an hour.
After a change to the editor only `ours`, `score` and `report` need running again: Word's PDFs and lines are kept.

A score is the share of Word's lines the editor doesn't place on the same page within 1px, plus lines Word doesn't
have, plus 20 per page of difference; 0 is the same as Word. Lines are matched as in `tools/word-compare`.

Keep the work folder outside the repository: the full set of documents and PDFs is a few hundred megabytes.
