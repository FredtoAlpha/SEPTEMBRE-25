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

test('utilise le fallback V2 quand la stratégie spécialisée échoue', () => {
  const scenarios = ['COM', 'TRA'];
  let fallbackInvocations = 0;

  const sandbox = createSandbox();
  sandbox.lancerEquilibrageScores_UI = () => {};
  sandbox.executerEquilibrageScoresPersonnalise = () => ({ success: false });
  sandbox.V2_Ameliore_OptimisationEngine = (sheet, dataContext, config) => {
    fallbackInvocations += 1;
    assert.deepStrictEqual(Array.from(config.COLONNES_SCORES_ACTIVES), scenarios);
    return {
      success: true,
      nbSwapsAppliques: 4,
      cyclesGeneraux: 2,
    };
  };

  const resultats = sandbox.executerPhaseScoresSpecialisee({}, {}, scenarios);

  assert.strictEqual(fallbackInvocations, 1, 'le fallback V2 doit être invoqué');
  assert.strictEqual(resultats.success, true, 'le résultat final doit être un succès');
  assert.strictEqual(resultats.details.strategieUtilisee, 'V2 Adaptée Scores');
});

test('le fallback V2 met à jour le nombre d\'opérations', () => {
  const scenarios = ['PART'];
  const sandbox = createSandbox();
  sandbox.lancerEquilibrageScores_UI = () => {};
  sandbox.executerEquilibrageScoresPersonnalise = () => ({ success: false });
  sandbox.V2_Ameliore_OptimisationEngine = () => ({
    success: true,
    nbSwapsAppliques: 7,
    cyclesGeneraux: 1,
  });

  const resultats = sandbox.executerPhaseScoresSpecialisee({}, {}, scenarios);

  assert.strictEqual(resultats.success, true);
  assert.strictEqual(resultats.nbOperations, 7);
  assert.strictEqual(resultats.details.cyclesGeneraux, 1);
});
