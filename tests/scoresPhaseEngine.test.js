const { test } = require('node:test');
const assert = require('node:assert/strict');

const { loadNirvanaSandbox } = require('./helpers/loadNirvanaSandbox');

test('ScoresEquilibrageEngine domain returns specialised success', () => {
  const sandbox = loadNirvanaSandbox();
  const domain = sandbox.ScoresEquilibrageEngine.createDomain({ logger: sandbox.Logger });

  const result = domain.run({
    scenarios: ['COM'],
    runCustom: () => ({
      success: true,
      nbOperations: 3,
      details: { strategieUtilisee: 'Spécialisée COM' }
    })
  });

  assert.strictEqual(result.success, true);
  assert.strictEqual(result.nbOperations, 3);
  assert.strictEqual(result.details.strategieUtilisee, 'Spécialisée COM');
  assert.ok(Array.isArray(result.details.history));
  assert.strictEqual(result.details.history.length, 1);
  const [entry] = result.details.history;
  assert.strictEqual(entry.type, 'specialisee');
  assert.strictEqual(entry.success, true);
  assert.deepStrictEqual({ ...entry.details }, { strategieUtilisee: 'Spécialisée COM' });
});

test('ScoresEquilibrageEngine domain falls back when specialised fails', () => {
  const sandbox = loadNirvanaSandbox();
  const domain = sandbox.ScoresEquilibrageEngine.createDomain({ logger: sandbox.Logger });
  let fallbackCalled = 0;

  const result = domain.run({
    scenarios: ['COM'],
    runCustom: () => ({
      success: false,
      nbOperations: 0,
      details: { strategieUtilisee: 'Spécialisée COM' }
    }),
    runFallback: () => {
      fallbackCalled += 1;
      return {
        success: true,
        nbOperations: 5,
        details: { strategieUtilisee: 'V2 Adaptée Scores' }
      };
    }
  });

  assert.strictEqual(fallbackCalled, 1);
  assert.strictEqual(result.success, true);
  assert.strictEqual(result.nbOperations, 5);
  assert.strictEqual(result.details.strategieUtilisee, 'V2 Adaptée Scores');
  assert.strictEqual(result.details.history.length, 2);
  const [, fallbackEntry] = result.details.history;
  assert.strictEqual(fallbackEntry.type, 'fallback');
  assert.strictEqual(fallbackEntry.success, true);
});

test('ScoresEquilibrageEngine domain fails gracefully without scenarios', () => {
  const sandbox = loadNirvanaSandbox();
  const domain = sandbox.ScoresEquilibrageEngine.createDomain({ logger: sandbox.Logger });

  const result = domain.run({ scenarios: [] });

  assert.strictEqual(result.success, false);
  assert.strictEqual(result.nbOperations, 0);
  assert.strictEqual(result.details.strategieUtilisee, 'Aucune');
  assert.ok(Array.isArray(result.details.history));
  assert.strictEqual(result.details.history.length, 0);
});

test('ScoresEquilibrageEngine service prepares config and triggers fallback', () => {
  const sandbox = loadNirvanaSandbox();
  const captured = [];

  const service = sandbox.ScoresEquilibrageEngine.createService({
    runCustomPhase: ({ config }) => {
      captured.push({ stage: 'specialisee', config: { ...config } });
      return {
        success: false,
        nbOperations: 0,
        details: { strategieUtilisee: 'Spécialisée COM+TRA' }
      };
    },
    runFallbackPhase: ({ config }) => {
      captured.push({ stage: 'fallback', config: { ...config } });
      return {
        success: true,
        nbOperations: 7,
        details: { strategieUtilisee: 'V2 Adaptée Scores', cyclesGeneraux: 2 }
      };
    }
  });

  const baseConfig = { MODE_AGRESSIF: false, seuil: 1 };
  const result = service.runPhase({
    dataContext: { classesState: {} },
    config: baseConfig,
    scenarios: ['COM', 'TRA']
  });

  assert.strictEqual(result.success, true);
  assert.strictEqual(result.nbOperations, 7);
  assert.strictEqual(result.details.cyclesGeneraux, 2);
  assert.deepStrictEqual(Array.from(result.details.scenarios), ['COM', 'TRA']);
  assert.strictEqual(captured.length, 2);
  assert.deepStrictEqual(Array.from(captured[0].config.COLONNES_SCORES_ACTIVES), ['COM', 'TRA']);
  assert.strictEqual(captured[0].config.MODE_AGRESSIF, true);
  assert.strictEqual(captured[0].config.MAX_ITERATIONS_SCORES, 50);
  assert.strictEqual(captured[1].config.MODE_AGRESSIF, true);
  assert.strictEqual(captured[1].config.MAX_ITERATIONS_SCORES, 50);
  assert.deepStrictEqual(Array.from(captured[1].config.COLONNES_SCORES_ACTIVES), ['COM', 'TRA']);
});
