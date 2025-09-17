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

test('déclenche le fallback si la stratégie principale réussit sans faire d\'opérations', () => {
  const scenarios = ['COM'];
  let fallbackInvocations = 0;

  const sandbox = createSandbox();

  // La stratégie principale est disponible
  sandbox.lancerEquilibrageScores_UI = () => {};

  // Elle réussit, mais ne fait aucune opération
  sandbox.executerEquilibrageScoresPersonnalise = () => ({
    success: true,
    totalEchanges: 0
  });

  // Le fallback est surveillé
  sandbox.V2_Ameliore_OptimisationEngine = () => {
    fallbackInvocations += 1;
    return {
      success: true,
      nbSwapsAppliques: 5,
    };
  };

  sandbox.executerPhaseScoresSpecialisee({}, {}, scenarios);

  // Le test doit vérifier que le fallback a bien été appelé.
  assert.strictEqual(fallbackInvocations, 1, 'Le fallback V2 aurait dû être invoqué');
});
