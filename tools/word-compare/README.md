# Comparing the layout with Word

Measures how closely the editor places every line compared with Word: both print the same documents to PDF, the text
lines of the two PDFs are matched in order, and a line counts when it's on the same page and its baseline is within
1px of Word's. See `docs/LAYOUT.md` for what the numbers mean and where they stand.

Windows, with Word and Microsoft Edge installed (the published numbers are Edge's: its fonts and font rendering are
part of what's measured). Node 20 for the repository's Playwright, Python 3, PowerShell 7.

## Once: Word's PDFs

```powershell
pwsh tools/word-compare/word-print.ps1 <folder of .docx> <folder for Word's PDFs>
```

Word opens each document read-only, invisibly, exports a PDF and closes it without saving; no Word setting is changed.
Documents already printed are skipped. Keep the PDFs: Word's output only changes when Word does.

## Each change: a run

```powershell
pwsh tools/word-compare/run.ps1 -Docs <folder of .docx> -Word <folder of Word's PDFs> -Out <work folder> -Run 101
```

It builds the demo (`node build.mjs`), prints every document through it (`print-ours.mjs`, eight at a time), extracts
the lines of both sets of PDFs (`lines.cjs`, Word's once) and writes `<work folder>/report-101.txt`: one line per
document, then the totals:

```
TOTAL docs 151 page counts 151/151 lines on same page 35432/35436 (100.0%) within 1px 35084/35432 (99.0%) ...
```

Compare two runs, document by document (which got better, which worse):

```powershell
python tools/word-compare/diff.py <work folder>/report-100.txt <work folder>/report-101.txt
```

A layout change is kept when the total gets better and nothing that got worse is left unexplained.

Runs are not quite deterministic: with no change at all, one document (`sample S125`) comes out in one of two states,
19 lines on its later pages 0.7px apart, which reads as 2 lines within 1px more or fewer. A change that moves only that
document by 2 has changed nothing; compare its lines with an earlier run's to be sure.

## The documents

Real documents, with their owners' permission, kept out of the repository. In code comments refer to one by a label
(`sample S078`), never by its name; keep the list of labels with the documents. The 19-document corpus in
`test/fixtures/make-layout-corpus*.mjs` can be generated and shared freely.
