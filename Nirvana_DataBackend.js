'use strict';

const NirvanaDataBackend = (function (global) {
  const noop = () => {};
  const defaultLogger = (global.Logger && {
    log: (...args) => global.Logger.log(...args),
    warn: (...args) => (global.Logger.warn ? global.Logger.warn(...args) : noop()),
    error: (...args) => (global.Logger.error ? global.Logger.error(...args) : global.Logger.log(...args))
  }) || global.console || { log: noop, warn: noop, error: noop };

  const DEFAULT_CRITERIA = ['COM', 'TRA', 'PART'];

  function ensureArray(value) {
    return Array.isArray(value) ? value.slice() : [];
  }

  function toClassKey(value) {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value).trim();
  }

  function createEmptyDistribution(criteria = DEFAULT_CRITERIA) {
    return criteria.reduce((acc, critere) => {
      acc[critere] = { '1': 0, '2': 0, '3': 0, '4': 0 };
      return acc;
    }, {});
  }

  function createEmptyCaches(classesState, criteria = DEFAULT_CRITERIA) {
    const caches = {};
    Object.keys(classesState).forEach(classe => {
      const dist = criteria.reduce((acc, critere) => {
        acc[critere] = { 1: 0, 2: 0, 3: 0, 4: 0 };
        return acc;
      }, {});
      caches[classe] = {
        dist,
        parite: { F: 0, M: 0, total: 0 },
        score: 0
      };
    });
    return caches;
  }

  function createDomain({
    computeDistribution = global.V2_Ameliore_CalculerDistributionGlobale,
    buildDissocCountMap = global.buildDissocCountMap,
    logger = defaultLogger
  } = {}) {
    function buildClassesState(students) {
      const classesState = {};
      const effectifsClasses = {};

      ensureArray(students).forEach(student => {
        if (!student) return;
        const classe = toClassKey(student.CLASSE);
        if (!classe) return;
        if (!classesState[classe]) {
          classesState[classe] = [];
        }
        classesState[classe].push(student);
      });

      Object.keys(classesState).forEach(classe => {
        effectifsClasses[classe] = classesState[classe].length;
      });

      return { classesState, effectifsClasses };
    }

    function buildTargets({ classesState, distributionGlobale, totalEleves, criteria = DEFAULT_CRITERIA }) {
      const ciblesParClasse = {};
      const total = Number.isFinite(totalEleves) ? totalEleves : 0;
      const distribution = distributionGlobale || createEmptyDistribution(criteria);

      Object.keys(classesState).forEach(classe => {
        const effectifClasse = classesState[classe].length;
        ciblesParClasse[classe] = {};
        criteria.forEach(critere => {
          ciblesParClasse[classe][critere] = {};
          const globalCritere = distribution[critere] || {};
          ['1', '2', '3', '4'].forEach(score => {
            const totalScore = Number(globalCritere[score]) || 0;
            const cible = total > 0 ? Math.round((effectifClasse / total) * totalScore) : 0;
            ciblesParClasse[classe][critere][score] = cible;
          });
        });
      });

      return ciblesParClasse;
    }

    function prepareContext({
      config,
      students,
      optionPools = {},
      structure,
      colIndexes = {},
      criteria = DEFAULT_CRITERIA
    } = {}) {
      const elevesValides = ensureArray(students);
      const { classesState, effectifsClasses } = buildClassesState(elevesValides);
      const totalEleves = elevesValides.length;
      const nbClasses = Object.keys(classesState).length;

      const distribution =
        typeof computeDistribution === 'function'
          ? computeDistribution(elevesValides, criteria)
          : createEmptyDistribution(criteria);

      const ciblesParClasse = buildTargets({
        classesState,
        distributionGlobale: distribution,
        totalEleves,
        criteria
      });

      const dissocMap =
        typeof buildDissocCountMap === 'function'
          ? buildDissocCountMap(classesState)
          : {};

      const dataContext = {
        config,
        elevesValides,
        classesState,
        effectifsClasses,
        optionPools: optionPools || {},
        structureData: structure,
        colIndexes: colIndexes || {},
        dissocMap,
        distributionGlobale: distribution,
        ciblesParClasse,
        totalEleves,
        nbClasses
      };

      dataContext.classeCaches = createEmptyCaches(classesState, criteria);
      dataContext.scoreGlobal = 0;

      return dataContext;
    }

    return {
      prepareContext,
      buildClassesState,
      buildTargets
    };
  }

  function createService({
    domain = createDomain(),
    getConfig = global.getConfig,
    determineActiveLevel = global.determinerNiveauActifCache,
    loadStructure = global.chargerStructureEtOptions,
    buildOptionPools = global.buildOptionPools,
    loadStudents = global.chargerElevesEtClasses_AvecSEXE || global.chargerElevesEtClasses,
    sanitizeStudents = global.sanitizeStudents,
    classifyStudents = global.classifierEleves,
    criteria = ['COM', 'TRA', 'PART', 'ABS'],
    logger = defaultLogger
  } = {}) {
    function prepareData({ config: configArg, criteresUI } = {}) {
      let config = configArg;
      if (!config && typeof getConfig === 'function') {
        config = getConfig(criteresUI);
      }
      if (!config || typeof config !== 'object') {
        throw new Error('Configuration Nirvana invalide');
      }

      const activeLevel =
        typeof determineActiveLevel === 'function' ? determineActiveLevel(config) : null;

      const structureResult =
        typeof loadStructure === 'function' ? loadStructure(activeLevel, config) : null;
      if (!structureResult || structureResult.success === false || !structureResult.structure) {
        throw new Error('Échec chargement structure Nirvana');
      }

      const optionPools =
        typeof buildOptionPools === 'function'
          ? buildOptionPools(structureResult.structure, config)
          : {};

      const loader = typeof loadStudents === 'function' ? loadStudents : () => ({ success: false });
      const loadResult = loader(config, config && config.MOBILITE_FIELD ? config.MOBILITE_FIELD : 'MOBILITE');
      if (!loadResult || loadResult.success === false) {
        throw new Error('Échec chargement élèves Nirvana');
      }

      const sanitized =
        typeof sanitizeStudents === 'function'
          ? sanitizeStudents(loadResult.students)
          : { clean: Array.isArray(loadResult.students) ? loadResult.students : [] };

      const elevesValides = ensureArray(sanitized.clean);

      if (typeof classifyStudents === 'function') {
        try {
          classifyStudents(elevesValides, criteria);
        } catch (err) {
          if (logger && typeof logger.warn === 'function') {
            logger.warn('NirvanaDataBackend: échec de la classification des élèves', err);
          }
        }
      }

      const distributionCriteria = criteria.filter(critere => critere !== 'ABS');

      return domain.prepareContext({
        config,
        students: elevesValides,
        optionPools,
        structure: structureResult.structure,
        colIndexes: loadResult.colIndexes,
        criteria: distributionCriteria
      });
    }

    return { prepareData };
  }

  const api = {
    createDomain,
    createService
  };

  global.NirvanaDataBackend = api;
  return api;
})(typeof globalThis !== 'undefined' ? globalThis : this);

const __nirvanaDataLogger =
  (typeof Logger !== 'undefined' && Logger) ||
  (typeof console !== 'undefined' ? console : { log: () => {}, warn: () => {}, error: () => {} });
const __nirvanaDataDomain = NirvanaDataBackend.createDomain({ logger: __nirvanaDataLogger });
const __nirvanaDataService = NirvanaDataBackend.createService({
  domain: __nirvanaDataDomain,
  logger: __nirvanaDataLogger
});

