const { test } = require('node:test');
const assert = require('node:assert/strict');
const { loadBackendSandbox } = require('./helpers/loadBackendSandbox');

test('ElevesBackend service structure les classes et règles', () => {
  const sandbox = loadBackendSandbox();
  const ElevesBackend = sandbox.ElevesBackend;

  const dataAccess = {
    getClassSheetsForSuffix: (suffix, options = {}) => {
      assert.strictEqual(suffix, ElevesBackend.config.sheetSuffixes.test);
      assert.strictEqual(options.includeValues, true);
      return [
        {
          name: `6E1${suffix}`,
          values: [
            ['ID_ELEVE', 'Nom', 'Prenom', 'MOB', 'COM'],
            ['ABC123', 'Durand', 'Alice', 'LIBRE', 42],
            ['DEF456', 'Martin', 'Bob', '', ''],
          ],
        },
      ];
    },
    getStructureSheetValues: () => [
      ['CLASSE_DEST', 'EFFECTIF', 'OPTIONS'],
      ['6E1', 28, 'ITA=12'],
    ],
  };

  const logger = { log: () => {}, warn: () => {}, error: () => {} };
  const domain = ElevesBackend.createDomain({ config: ElevesBackend.config, logger });
  const service = ElevesBackend.createService({
    config: ElevesBackend.config,
    domain,
    dataAccess,
    logger,
  });

  const eleves = service.getElevesData();

  const expected = domain.sanitize([
    {
      classe: '6E1',
      eleves: [
        {
          id: 'ABC123',
          nom: 'Durand',
          prenom: 'Alice',
          sexe: '',
          lv2: '',
          opt: '',
          disso: '',
          asso: '',
          scores: { C: 42, T: 0, P: 0, A: 0 },
          source: '',
          dispo: '',
          mobilite: 'LIBRE',
        },
        {
          id: 'DEF456',
          nom: 'Martin',
          prenom: 'Bob',
          sexe: '',
          lv2: '',
          opt: '',
          disso: '',
          asso: '',
          scores: { C: 0, T: 0, P: 0, A: 0 },
          source: '',
          dispo: '',
          mobilite: 'LIBRE',
        },
      ],
    },
  ]);

  assert.deepStrictEqual(eleves, expected);

  const rules = service.getStructureRules();
  const expectedRules = domain.sanitize({
    '6E1': { capacity: 28, quotas: { ITA: 12 } },
  });
  assert.deepStrictEqual(rules, expectedRules);
});

test('ElevesBackend service bascule sur TEST pour un mode inconnu', () => {
  const sandbox = loadBackendSandbox();
  const ElevesBackend = sandbox.ElevesBackend;

  const suffixes = [];
  const dataAccess = {
    getClassSheetsForSuffix: (suffix) => {
      suffixes.push(suffix);
      return [];
    },
    getStructureSheetValues: () => null,
  };

  const logger = { log: () => {}, warn: () => {}, error: () => {} };
  const domain = ElevesBackend.createDomain({ config: ElevesBackend.config, logger });
  const service = ElevesBackend.createService({
    config: ElevesBackend.config,
    domain,
    dataAccess,
    logger,
  });

  const result = service.getElevesDataForMode('INCONNU');
  assert.ok(Array.isArray(result));
  assert.strictEqual(result.length, 0);
  assert.deepStrictEqual(suffixes, [ElevesBackend.config.sheetSuffixes.test]);
});
