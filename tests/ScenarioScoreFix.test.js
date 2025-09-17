const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createSandbox(overrides = {}) {
  const sandbox = {
    Logger: { log: () => {} },
  };

  sandbox.global = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.self = sandbox;

  const filePath = path.join(__dirname, '..', 'Nirvana_Combined_Orchestrator.js');
  const code = fs.readFileSync(filePath, 'utf8');
  vm.runInNewContext(code, sandbox, { filename: 'Nirvana_Combined_Orchestrator.js' });

  Object.assign(sandbox, overrides);
  return sandbox;
}

test('calculerScoreEquilibrageScenario reads scenario scores correctly', () => {
  const sandbox = createSandbox();

  const dataContext = {
    classesState: {
      '3A': [
        { ID_ELEVE: '1', SEXE: 'F', COM: 4 },
        { ID_ELEVE: '2', SEXE: 'M', COM: 4 },
        { ID_ELEVE: '3', SEXE: 'F', COM: 2 },
      ]
    }
  };

  const scenario = 'COM';

  // Expected calculation without the bug:
  // effectifs = [2, 1] (two students with score 4, one with score 2)
  // moyenne = 1.5
  // variance = ((2-1.5)^2 + (1-1.5)^2) / 2 = 0.25
  // ecartType = sqrt(0.25) = 0.5
  // scoreClasse = 100 - (0.5 * 20) = 90
  const expectedScore = 90;

  const actualScore = sandbox.calculerScoreEquilibrageScenario(dataContext, scenario);

  assert.strictEqual(actualScore, expectedScore, 'The calculated score should be based on the correct "COM" property.');
});
