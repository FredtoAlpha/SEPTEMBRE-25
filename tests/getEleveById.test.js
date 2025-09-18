const { test } = require('node:test');
const assert = require('node:assert/strict');
const { loadBackendSandbox } = require('./helpers/loadBackendSandbox');

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
  assert.strictEqual(eleve.scores.T, 0);
  assert.strictEqual(eleve.scores.P, 0);
  assert.strictEqual(eleve.scores.A, 0);
});

test('getEleveById_ accepte la colonne strictement nommée ID mais ignore ID_PARENT', () => {
  const data = [
    ['ID_PARENT', 'ID', 'Nom'],
    ['P-001', 'XYZ999', 'Martin'],
  ];

  const fakeSheet = {
    getName: () => '5E2TEST',
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

  const eleve = sandbox.getEleveById_('XYZ999');

  assert.ok(eleve, 'La colonne ID doit être reconnue.');
  assert.strictEqual(eleve.id, 'XYZ999');
  assert.strictEqual(eleve.nom, 'Martin');
});

test('getEleveById_ privilégie ID_ELEVE même en présence de colonnes ID_PARENT', () => {
  const data = [
    ['ID_PARENT', 'ID_ELEVE', 'Nom'],
    ['P-001', 'ABC001', 'Alice'],
    ['P-002', 'ABC002', 'Bob'],
  ];

  const fakeSheet = {
    getName: () => '4E3TEST',
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

  const eleve = sandbox.getEleveById_('ABC002');

  assert.ok(eleve, 'L\'alias ID_ELEVE doit être prioritaire.');
  assert.strictEqual(eleve.id, 'ABC002');
  assert.strictEqual(eleve.nom, 'Bob');
});

test('getEleveById_ retourne null lorsqu\'aucune colonne aliasée ID n\'est disponible', () => {
  const data = [
    ['ID_PARENT', 'Nom'],
    ['P-001', 'Alice'],
    ['P-002', 'Bob'],
  ];

  const fakeSheet = {
    getName: () => '4E3TEST',
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

  const eleve = sandbox.getEleveById_('ABC002');

  assert.strictEqual(eleve, null);
});
