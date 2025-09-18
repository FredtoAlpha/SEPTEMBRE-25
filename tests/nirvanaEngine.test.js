const test = require('node:test');
const assert = require('node:assert/strict');

const { loadNirvanaSandbox } = require('./helpers/loadNirvanaSandbox');

function createTimeProvider(instants) {
  let index = 0;
  return () => {
    const value = instants[index];
    if (index < instants.length - 1) {
      index += 1;
    }
    return value instanceof Date ? value : new Date(value);
  };
}

test('NirvanaEngine domain agrège les résultats et formate le message', () => {
  const sandbox = loadNirvanaSandbox();
  const domain = sandbox.NirvanaEngine.createDomain({ logger: sandbox.Logger });

  const start = new Date('2025-01-01T10:00:00Z');
  const end = new Date('2025-01-01T10:00:10Z');

  const summary = domain.runCombined({
    config: { demo: true },
    dataContext: { classesState: {} },
    runV2Phase: () => ({ swapsV2: [1, 2, 3], cyclesGeneraux: 2, cyclesParite: 1 }),
    runParityPhase: () => ({ nbApplied: 4 }),
    computeState: () => ({ scoreGlobal: 87.1234 }),
    now: createTimeProvider([end]),
    startedAt: start
  });

  assert.equal(summary.success, true);
  assert.equal(summary.swapsV2, 3);
  assert.equal(summary.cyclesGeneraux, 2);
  assert.equal(summary.cyclesParite, 1);
  assert.equal(summary.operationsParity, 4);
  assert.equal(summary.totalOperations, 7);
  assert.equal(summary.tempsMs, 10000);
  assert.equal(summary.scoreFinal, 87.1234);
  assert.equal(summary.startedAt.toISOString(), start.toISOString());
  assert.equal(summary.endedAt.toISOString(), end.toISOString());

  const message = domain.formatSuccess(summary);
  assert.match(message, /Swaps principaux: 3/);
  assert.match(message, /Corrections parité: 4/);
  assert.match(message, /Score final: 87\.12\/100/);
  assert.match(message, /Durée totale: 10\.0 secondes/);
});

test('NirvanaEngine service orchestre les verrous et interactions UI', () => {
  const sandbox = loadNirvanaSandbox();
  const domain = sandbox.NirvanaEngine.createDomain({ logger: sandbox.Logger });

  let released = false;
  const toasts = [];
  const alerts = [];

  const spreadsheet = {
    toast(message, title, seconds) {
      toasts.push({ message, title, seconds });
    },
    getUi() {
      return {
        ButtonSet: { OK: 'OK' },
        alert(title, message, button) {
          alerts.push({ title, message, button });
        }
      };
    }
  };

  const service = sandbox.NirvanaEngine.createService({
    domain,
    getConfig: () => ({ demo: true }),
    prepareData: () => ({ classesState: {} }),
    runV2Phase: () => ({ swapsV2: ['a'], cyclesGeneraux: 1, cyclesParite: 2 }),
    runParityPhase: () => ({ nbApplied: 3 }),
    computeState: () => ({ scoreGlobal: 91 }),
    spreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    lockService: {
      getScriptLock: () => ({
        tryLock: () => true,
        releaseLock: () => {
          released = true;
        }
      })
    },
    logger: {
      log: () => {},
      warn: () => {},
      error: () => {}
    },
    now: createTimeProvider([
      new Date('2025-01-01T08:00:00Z'),
      new Date('2025-01-01T08:00:05Z'),
      new Date('2025-01-01T08:00:07Z')
    ])
  });

  const result = service.runCombination();

  assert.equal(result.success, true);
  assert.equal(result.swapsV2, 1);
  assert.equal(result.operationsParity, 3);
  assert.equal(result.totalOperations, 4);
  assert.equal(result.scoreFinal, 91);
  assert.equal(result.tempsMs, 5000);
  assert.equal(released, true);
  assert.equal(toasts.length, 4);
  assert.equal(alerts.length, 1);
  assert.match(alerts[0].message, /Swaps principaux: 1/);
  assert.match(toasts[toasts.length - 1].message, /4 opérations appliquées/);
});
