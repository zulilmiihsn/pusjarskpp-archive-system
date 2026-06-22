'use strict';

// Flat config untuk ESLint v9+/v10. Menggantikan .eslintrc.json (format lama
// yang tak lagi didukung tanpa flag kompat).
//
// Catatan GAS: seluruh file .gs berbagi SATU global scope saat di-deploy
// (clasp mengonkatenasi), jadi referensi lintas-file (AppController,
// ConfigService, requireAuth_, cleanId_, CONFIG_SHEETS, dll.) valid di runtime.
// Karena itu `no-undef` dimatikan — kalau tidak, ribuan false-positive muncul.
// Linting difokuskan ke file .gs; blok <script> di dalam .html butuh plugin
// terpisah yang tidak terpasang, jadi .html diabaikan.

const GAS_GLOBALS = {
  google: 'readonly',
  MimeType: 'readonly',
  Browser: 'readonly',
  AlertButton: 'readonly',
  ButtonSet: 'readonly',
  HtmlService: 'readonly',
  SpreadsheetApp: 'readonly',
  Session: 'readonly',
  Utilities: 'readonly',
  UrlFetchApp: 'readonly',
  ScriptApp: 'readonly',
  DriveApp: 'readonly',
  LockService: 'readonly',
  PropertiesService: 'readonly',
  SlidesApp: 'readonly',
  CacheService: 'readonly',
  Drive: 'readonly',
  Logger: 'readonly',
  console: 'readonly'
};

module.exports = [
  {
    ignores: ['node_modules/**', 'scripts/**', '**/*.html', 'eslint.config.js']
  },
  {
    files: ['**/*.gs', '**/*.js'],
    languageOptions: {
      ecmaVersion: 2019,
      sourceType: 'script',
      globals: GAS_GLOBALS
    },
    rules: {
      // Rule kustom proyek (dari .eslintrc.json lama).
      'no-unused-vars': ['warn', { args: 'none', varsIgnorePattern: '^_' }],
      'no-console': 'off',
      eqeqeq: 'error',
      'no-empty': ['error', { allowEmptyCatch: true }],
      'prefer-const': 'warn',
      'no-var': 'warn',
      'no-undef': 'off',

      // Subset rule core penangkap-bug (tak butuh paket @eslint/js).
      'no-dupe-keys': 'error',
      'no-dupe-args': 'error',
      'no-unreachable': 'error',
      'no-cond-assign': ['error', 'except-parens'],
      'no-constant-condition': ['error', { checkLoops: false }],
      'no-fallthrough': 'error',
      'no-func-assign': 'error',
      'no-irregular-whitespace': 'error',
      'use-isnan': 'error',
      'valid-typeof': 'error'
    }
  }
];
