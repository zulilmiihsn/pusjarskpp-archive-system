const fs = require('fs');
const path = require('path');

// Basic GAS Mocks
const Properties = {};
global.PropertiesService = {
  getScriptProperties: () => ({
    getProperty: (k) => Properties[k],
    setProperty: (k, v) => Properties[k] = v,
    deleteProperty: (k) => delete Properties[k]
  })
};

const Cache = {};
global.CacheService = {
  getScriptCache: () => ({
    get: (k) => Cache[k] || null,
    put: (k, v, exp) => Cache[k] = v,
    remove: (k) => delete Cache[k]
  })
};

global.Utilities = {
  getUuid: () => Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15),
  base64EncodeWebSafe: (str) => Buffer.from(str).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''),
  computeDigest: () => [1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6],
  DigestAlgorithm: { SHA_256: 'SHA_256' },
  Charset: { UTF_8: 'UTF-8' },
  formatDate: (d, tz, format) => d.toISOString(),
  sleep: (ms) => {}
};

global.Session = {
  getEffectiveUser: () => ({
    getEmail: () => 'admin@example.com'
  })
};

global.Logger = {
  log: console.log
};

// Very basic Drive Mock
const mockFiles = {};
global.DriveApp = {
  getFileById: (id) => {
    if (!mockFiles[id]) throw new Error('File not found: ' + id);
    return mockFiles[id];
  },
  getFolderById: (id) => {
    if (!mockFiles[id]) throw new Error('Folder not found: ' + id);
    return mockFiles[id];
  },
  createFolder: (name) => {
    const id = Utilities.getUuid();
    mockFiles[id] = { id, name, getId: () => id, getName: () => name, setTrashed: () => {}, moveTo: () => {} };
    return mockFiles[id];
  }
};

global.SpreadsheetApp = {
  openById: (id) => {
    if (!mockFiles[id]) throw new Error('Spreadsheet not found');
    return mockFiles[id];
  },
  create: (name) => {
    const id = Utilities.getUuid();
    mockFiles[id] = { id, name, getId: () => id, getName: () => name };
    return mockFiles[id];
  }
};

global.HtmlService = {
  createHtmlOutputFromFile: (name) => ({
    getContent: () => '{}'
  })
};

// Fake LockService
global.LockService = {
  getScriptLock: () => ({
    tryLock: () => true,
    releaseLock: () => {}
  })
};

// Fake ScriptApp
global.ScriptApp = {
  getProjectTriggers: () => [],
  newTrigger: () => ({ timeBased: () => ({ everyHours: () => ({ create: () => ({}) }) }) })
};

// Load actual GAS code into NodeJS sandbox (using eval)
function loadGasScripts() {
  const root = path.resolve(__dirname, '..');
  const files = [
    'ConfigConstants.gs', 'ConfigHelpers.gs', 'PureFunctions.gs', 'SecurityHelpers.gs', 'ConfigRepository.gs', 
    'ConfigService.gs', 'DriveService.gs', 'SheetHelpers.gs', 'SpreadsheetService.gs', 
    'WorkspaceSetup.gs', 'AuthService.gs', 'AccountController.gs', 'SettingsController.gs', 
    'WorkspaceController.gs', 'TemplateController.gs', 'SubActivityController.gs', 'ArchiveController.gs',
    'DriveController.gs'
  ];
  let fullCode = '';
  for (const f of files) {
    let content = fs.readFileSync(path.join(root, f), 'utf8');
    // strip out 'use strict'; at module level
    content = content.replace(/'use strict';/g, '');
    fullCode += content + '\n\n';
  }
  
  // We eval the combined code in global scope
  const script = new (require('vm').Script)(fullCode);
  script.runInThisContext();
}

module.exports = {
  loadGasScripts,
  Properties,
  Cache,
  mockFiles,
  resetMocks: () => {
    Object.keys(Properties).forEach(k => delete Properties[k]);
    Object.keys(Cache).forEach(k => delete Cache[k]);
    Object.keys(mockFiles).forEach(k => delete mockFiles[k]);
  }
};
