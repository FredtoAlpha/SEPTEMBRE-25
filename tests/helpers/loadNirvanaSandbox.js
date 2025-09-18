const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createDefaultSpreadsheet() {
  const alerts = [];
  const toasts = [];
  return {
    _alerts: alerts,
    _toasts: toasts,
    toast(message, title, seconds) {
      toasts.push({ message, title, seconds });
    },
    getUi() {
      return {
        ButtonSet: { OK: 'OK' },
        _alerts: alerts,
        alert(title, message, button) {
          alerts.push({ title, message, button });
        }
      };
    }
  };
}

function loadNirvanaSandbox(overrides = {}) {
  const spreadsheet = createDefaultSpreadsheet();

  const sandbox = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    Logger: {
      log: () => {},
      warn: () => {},
      error: () => {}
    },
    LockService: {
      getScriptLock: () => ({
        tryLock: () => true,
        releaseLock: () => {}
      })
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => spreadsheet
    },
    Date,
    JSON,
    Math,
    Number,
    String,
    Array,
    Object,
    RegExp,
    Error,
    isFinite
  };

  sandbox.global = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.self = sandbox;

  const filePath = path.join(__dirname, '..', '..', 'Nirvana_Combined_Orchestrator.js');
  const code = fs.readFileSync(filePath, 'utf8');
  vm.runInNewContext(code, sandbox, { filename: 'Nirvana_Combined_Orchestrator.js' });

  Object.assign(sandbox, overrides);
  sandbox.__spreadsheet = spreadsheet;

  return sandbox;
}

module.exports = { loadNirvanaSandbox };
