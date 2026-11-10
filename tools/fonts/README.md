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
