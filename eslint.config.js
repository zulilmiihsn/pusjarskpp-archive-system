module.exports = [
  {
    ignores: [
      'node_modules/**',
      'assets/**',
      'docs/**'
    ]
  },
  {
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'script',
      globals: {
        // Apps Script Core Services
        Logger: 'readonly',
        DriveApp: 'readonly',
        SpreadsheetApp: 'readonly',
        CacheService: 'readonly',
        PropertiesService: 'readonly',
        ScriptApp: 'readonly',
        UrlFetchApp: 'readonly',
        HtmlService: 'readonly',
        LockService: 'readonly',
        Session: 'readonly',
        Utilities: 'readonly',
        XmlService: 'readonly',
        GmailApp: 'readonly',
        CalendarApp: 'readonly',
        ContactsApp: 'readonly',
        DocumentApp: 'readonly',
        FormApp: 'readonly',
        SlidesApp: 'readonly',
        Drive: 'readonly',

        // Node.js test environment globals
        console: 'readonly',
        process: 'readonly',
        require: 'readonly',
        exports: 'readonly',
        module: 'readonly',
        __dirname: 'readonly',
        __filename: 'readonly',
        setTimeout: 'readonly',
        clearTimeout: 'readonly',
        setInterval: 'readonly',
        clearInterval: 'readonly',
        Buffer: 'readonly'
      }
    },
    rules: {
      'no-undef': 'off', // GAS scripts share a single flat global namespace across files
      'no-redeclare': 'warn',
      'no-constant-condition': ['error', { checkLoops: false }],
      'no-empty': ['error', { allowEmptyCatch: true }],
      'no-var': 'warn',
      'prefer-const': ['warn', { destructuring: 'all' }],
      'eqeqeq': ['error', 'always', { null: 'ignore' }]
    }
  }
];
