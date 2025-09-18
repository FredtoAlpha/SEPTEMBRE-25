const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadBackendSandbox(overrides = {}) {
  const sandbox = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    Logger: { log: () => {} },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheets: () => [] })
    },
  };

  sandbox.global = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.self = sandbox;

  const filePath = path.join(__dirname, '..', '..', 'BackendV2.js');
  const code = fs.readFileSync(filePath, 'utf8');
  vm.runInNewContext(code, sandbox, { filename: 'BackendV2.js' });

  Object.assign(sandbox, overrides);
  return sandbox;
}

module.exports = { loadBackendSandbox };
