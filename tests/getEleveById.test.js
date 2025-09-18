const { test } = require('node:test');
const assert = require('node:assert/strict');
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

  const filePath = path.join(__dirname, '..', 'BackendV2.js');
  const code = fs.readFileSync(filePath, 'utf8');
  vm.runInNewContext(code, sandbox, { filename: 'BackendV2.js' });

  Object.assign(sandbox, overrides);
  return sandbox;
}

test('getEleveById_ accepte les alias ID_ELEVE pour la colonne identifiant', () => {
  const data = [
    ['ID_ELEVE', 'Nom', 'Prenom', 'MOB', 'COM'],
    ['ABC123', 'Durand', 'Alice', 'LIBRE', 42],
  ];

  const fakeSheet = {
    getName: () => '6E1TEST',
    getDataRange: () => ({
      getValues: () => data,
    }),
  };

  const sandbox = loadBackendSandbox({
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheets: () => [fakeSheet],
      }),
    },
  });

  const eleve = sandbox.getEleveById_('ABC123');

  assert.ok(eleve, 'Un élève doit être renvoyé lorsque l\'identifiant correspond.');
  assert.strictEqual(eleve.id, 'ABC123');
  assert.strictEqual(eleve.nom, 'Durand');
  assert.strictEqual(eleve.prenom, 'Alice');
  assert.strictEqual(eleve.mobilite, 'LIBRE');
  assert.strictEqual(eleve.scores.C, 42);
  assert.strictEqual(eleve.scores.T, undefined);
  assert.strictEqual(eleve.scores.P, undefined);
  assert.strictEqual(eleve.scores.A, undefined);
});
