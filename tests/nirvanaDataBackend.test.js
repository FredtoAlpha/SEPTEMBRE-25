const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');
const assert = require('node:assert/strict');

function loadNirvanaDataSandbox(overrides = {}) {
  const sandbox = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    Logger: { log: () => {}, warn: () => {}, error: () => {} },
    Date,
    Math,
    Number,
    String,
    Array,
    Object,
    JSON,
    RegExp,
    Error
  };

  sandbox.global = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.self = sandbox;

  Object.assign(sandbox, overrides);

  const filePath = path.join(__dirname, '..', 'Nirvana_DataBackend.js');
  const code = fs.readFileSync(filePath, 'utf8');
  vm.runInNewContext(code, sandbox, { filename: 'Nirvana_DataBackend.js' });

  return sandbox;
}

test('NirvanaDataBackend.prepareData construit un contexte complet', () => {
  const computeDistribution = (students, criteria) => {
    const distribution = {};
    criteria.forEach(critere => {
      distribution[critere] = { '1': 0, '2': 0, '3': 0, '4': 0 };
    });

    students.forEach(student => {
      criteria.forEach(critere => {
        const value = student?.[critere];
        const key = value !== undefined && value !== null ? String(value) : null;
        if (key && Object.prototype.hasOwnProperty.call(distribution[critere], key)) {
          distribution[critere][key] += 1;
        }
      });
    });

    return distribution;
  };

  const sandbox = loadNirvanaDataSandbox();
  const { NirvanaDataBackend } = sandbox;
  const logs = [];

  const domain = NirvanaDataBackend.createDomain({
    computeDistribution,
    buildDissocCountMap: () => ({}) ,
    logger: { log: () => {}, warn: () => {}, error: () => {} }
  });

  const service = NirvanaDataBackend.createService({
    domain,
    getConfig: () => ({ TEST_SUFFIX: 'TEST' }),
    determineActiveLevel: () => 'TROISIEME',
    loadStructure: () => ({ success: true, structure: { classes: [] } }),
    buildOptionPools: () => ({ ESP: ['3A_TEST'] }),
    loadStudents: () => ({
      success: true,
      students: [
        { ID_ELEVE: 'E1', CLASSE: '3A_TEST', SEXE: 'F', COM: 1, TRA: 2, PART: 3 },
        { ID_ELEVE: 'E2', CLASSE: '3B_TEST', SEXE: 'M', COM: 2, TRA: 3, PART: 4 }
      ],
      colIndexes: { ID_ELEVE: 0 }
    }),
    sanitizeStudents: rows => ({ clean: rows }),
    classifyStudents: (rows, criteria) => {
      logs.push({ rows: rows.length, criteria });
    },
    criteria: ['COM', 'TRA', 'PART'],
    logger: { log: () => {}, warn: () => {}, error: () => {} }
  });

  const context = service.prepareData();

  assert.equal(context.totalEleves, 2);
  assert.equal(context.nbClasses, 2);
  assert.deepEqual(Object.keys(context.classesState).sort(), ['3A_TEST', '3B_TEST']);
  assert.equal(context.classesState['3A_TEST'].length, 1);
  assert.equal(context.classesState['3B_TEST'].length, 1);
  assert.deepEqual(context.optionPools, { ESP: ['3A_TEST'] });
  assert.deepEqual(context.colIndexes, { ID_ELEVE: 0 });
  assert.equal(context.ciblesParClasse['3A_TEST'].COM['1'], 1);
  assert.equal(context.ciblesParClasse['3B_TEST'].COM['2'], 1);
  assert.equal(context.classeCaches['3A_TEST'].parite.total, 0);
  assert.equal(logs.length, 1);
});

test('NirvanaDataBackend.prepareData signale les erreurs de classification sans échouer', () => {
  const warnings = [];

  const sandbox = loadNirvanaDataSandbox();
  const { NirvanaDataBackend } = sandbox;

  const domain = NirvanaDataBackend.createDomain({
    computeDistribution: () => ({ COM: { '1': 0, '2': 0, '3': 0, '4': 0 } }),
    buildDissocCountMap: () => ({})
  });

  const service = NirvanaDataBackend.createService({
    domain,
    getConfig: () => ({}),
    determineActiveLevel: () => null,
    loadStructure: () => ({ success: true, structure: {} }),
    buildOptionPools: () => ({}),
    loadStudents: () => ({ success: true, students: [{ CLASSE: '3A_TEST' }], colIndexes: {} }),
    sanitizeStudents: rows => ({ clean: rows }),
    classifyStudents: () => {
      throw new Error('boom');
    },
    criteria: ['COM'],
    logger: { log: () => {}, warn: message => warnings.push(message), error: () => {} }
  });

  const context = service.prepareData();
  assert.equal(context.totalEleves, 1);
  assert.equal(context.nbClasses, 1);
  assert.ok(warnings.length >= 1);
});
