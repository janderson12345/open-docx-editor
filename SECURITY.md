# Security

open-docx-editor runs in the browser and opens files people give it, so a document that makes it run script, read
data it shouldn't, or hang the page is a security problem. So is anything that lets a saved file carry something its
author didn't put there.

## Reporting a problem

Please don't open a public issue for it. Use GitHub's private reporting instead: on the repository's **Security** tab,
choose **Report a vulnerability**. Include the document (or a small one that shows the problem), the browser, and
what happened. You'll get an answer within a week.

## What it does to keep documents safe

- Nothing in a document is run: macros (a .docm's VBA project) are kept as bytes and written back, never executed;
  there are no scripts and no remote content. Pictures are shown from the file itself; a picture linked to an
  outside address isn't fetched.
- A link is followed only to `http:`, `https:`, `mailto:` and `tel:` addresses, and to places in the document; any
  other address (`javascript:`, `data:`, `file:`...) stays in the file as it was but isn't followed. Pasted links
  follow the same rule.
- The document's XML is read as data and drawn through ProseMirror's rendering; equations are built as MathML from
  its escaped text, and shapes, charts and EMF pictures are drawn as SVG images, which can't run script.
- Saving writes back the parts of the file that were in it, plus only the ones an edit needs.

## Supported versions

Fixes go into the latest release.
