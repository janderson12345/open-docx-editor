// ESLint for mistakes, not formatting: names used but never defined, defined but never used, unreachable code.
// The code's style (long lines with "why" comments, compact arrow functions) is kept by hand and by .editorconfig.
import js from '@eslint/js';
import globals from 'globals';

export default [
  { ignores: ['dist/**', 'demo/dist/**', 'node_modules/**', 'test/out/**', 'coverage/**'] },
  js.configs.recommended,
  {
    files: ['**/*.js', '**/*.mjs', '**/*.cjs'],
    languageOptions: {
      ecmaVersion: 2023,
      sourceType: 'module',
      globals: { ...globals.browser, ...globals.node },
    },
    rules: {
      // (Unused arguments are often a callback's signature, kept for readability; unused locals aren't.)
      'no-unused-vars': ['error', { args: 'none', caughtErrors: 'none', varsIgnorePattern: '^_' }],
      'no-empty': ['error', { allowEmptyCatch: true }],
      // (A label inside a `switch` case block, and regexes with control characters for XML, are deliberate.)
      'no-case-declarations': 'off',
      'no-control-regex': 'off',
    },
  },
  { files: ['**/*.cjs'], languageOptions: { sourceType: 'commonjs' } },
];
