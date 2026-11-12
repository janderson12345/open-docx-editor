# Bundled fonts

`src/fonts/` holds metric-compatible stand-ins for Word's common fonts and `fonts.css`, which offers them under the
names the editor's font stacks use. `build-fonts.mjs` makes it from:

- the `@fontsource` packages of Carlito, Arimo, Tinos and Cousine (`npm pack @fontsource/carlito @fontsource/arimo
  @fontsource/tinos @fontsource/cousine`, then untar each into one folder): their Latin and Latin Extended files;
- Caladea's 2013 release (`crosextrafonts-20130214.tar.gz`, published by Google for ChromeOS), converted to WOFF2:
  `python -m fontTools.ttLib.woff2 compress -o Caladea-Regular.woff2 Caladea-Regular.ttf` for each face. Later
  Caladea releases, including the one on Google Fonts and fontsource, changed its widths and no longer match Cambria's.

```sh
node tools/fonts/build-fonts.mjs <folder with the unpacked @fontsource packages> <folder with the Caladea .woff2 files>
```

Each face is checked against the installed Microsoft font (widths of each letter, ascent, descent) before the files
are replaced.

## The Aptos stand-ins

`src/fonts/files/odx-aptos-*.woff2` are eight small fonts for Aptos and Aptos Display (regular, bold, italic, bold
italic), with their `@font-face` lines in `src/fonts/fonts.css`: Source Sans 3's Latin letters (its letters are closer
to Aptos's than Arial's design), each narrowed or widened sideways to the width the editor gives Aptos's letter (fitted
to Word's PDFs), the font scaled so its x-height is Aptos's. `build-aptos.mjs` measures those widths as the stand-ins
draw them now (in Microsoft Edge), builds the fonts with `aptos_standin.py` (Python with fontTools and brotli) and
writes them and their CSS. It needs Source Sans 3's TTFs (from github.com/adobe-fonts/source-sans releases):

```sh
node tools/fonts/build-aptos.mjs <folder with SourceSans3-Regular.ttf, -Bold.ttf, -It.ttf, -BoldIt.ttf>
```
