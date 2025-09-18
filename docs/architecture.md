# Architecture interne – module BackendV2

Cette refonte isole les responsabilités majeures de la chaîne « backend élèves » en trois couches.

## 1. Domaine (`ElevesBackend.createDomain`)

* centralise la résolution des alias déclarés dans `ELEVES_ALIAS` ;
* expose `createStudent`, `createClassFromSheet` et `buildClassesData` pour transformer les lignes brutes en objets élèves normalisés ;
* fournit `parseStructureRules` pour interpréter l’onglet `_STRUCTURE` (capacités et quotas) ;
* garantit la sérialisation stable via `sanitize`.

## 2. Accès données (`ElevesBackend.createDataAccess`)

* encapsule la sélection des feuilles en fonction des suffixes (`TEST`, `CACHE`, `INT`) et filtre les onglets techniques (`level_`, `grp_`, …) ;
* lit de manière centralisée `_STRUCTURE` et expose `getClassSheetsForSuffix` / `getStructureSheetValues`.

## 3. Service (`ElevesBackend.createService`)

* associe domaine + accès données pour alimenter les fonctions publiques `getElevesData`, `getElevesDataForMode` et `getStructureRules` ;
* normalise la sélection du suffixe demandé (`resolveSuffix`) et journalise les modes inconnus ;
* fournit une surface testable facilement injectée dans les tests Node (`tests/elevesService.test.js`).

## Intégration Apps Script

`BackendV2.js` instancie le service avec `SpreadsheetApp` par défaut afin que les fonctions globales historiques restent inchangées pour l’Apps Script UI.

`getEleveById_`, `buildStudentIndex_` et `getAvailableClasses` réutilisent désormais le domaine pour éviter la duplication des règles d’alias.

## Tests

Les nouveaux tests valident :

* la construction du service (`tests/elevesService.test.js`) ;
* la compatibilité de `getEleveById_` avec les alias (`tests/getEleveById.test.js`).

Cette architecture permet de brancher progressivement d’autres moteurs (API REST, batch Node, etc.) en réutilisant le même cœur métier.

# Architecture interne – Nirvana_DataBackend

Le module `Nirvana_DataBackend.js` généralise la préparation du contexte de données utilisé par les moteurs Nirvana.

## 1. Domaine (`NirvanaDataBackend.createDomain`)

* agrège les élèves par classes, calcule les effectifs et construit les cibles par classe à partir de la distribution globale ;
* délègue le calcul de distribution à `V2_Ameliore_CalculerDistributionGlobale` (surchargé au besoin dans les tests) ;
* expose `prepareContext` pour fabriquer un `dataContext` complet (caches, dissociations, cibles) à partir d’entrées normalisées.

## 2. Service (`NirvanaDataBackend.createService`)

* centralise l’accès aux dépendances Apps Script (`determinerNiveauActifCache`, `chargerStructureEtOptions`, `chargerElevesEtClasses_AvecSEXE`, `sanitizeStudents`, `classifierEleves`) ;
* produit un contexte unique en journalisant les erreurs de classification et en acceptant des critères configurables ;
* alimente directement `V2_Ameliore_PreparerDonnees` et `V2_Ameliore_PreparerDonnees_AvecSEXE` qui deviennent de simples wrappers.

## 3. Tests

* `tests/nirvanaDataBackend.test.js` valide la construction du contexte et la tolérance aux erreurs de classification sans dépendre de Google Sheets.

# Architecture interne – Nirvana_Combined_Orchestrator

La refonte de l’orchestrateur combine les mêmes principes de séparation des responsabilités.

## 1. Domaine (`NirvanaEngine.createDomain`)

* valide l’entrée (configuration, `classesState`) avant de lancer les phases ;
* déclenche les hooks `beforePhase1` / `beforePhase2` pour permettre au service d’afficher toasts et journaux ;
* agrège les résultats des phases V2/Parité, calcule les durées, le score final et construit un résumé sérialisable (`totalOperations`, `startedAt`, `endedAt`).

## 2. Service (`NirvanaEngine.createService`)

* encapsule les dépendances Apps Script (LockService, `SpreadsheetApp`, `UI`) et gère le verrouillage, les toasts et les alertes ;
* injecte dynamiquement `getConfig`, `V2_Ameliore_PreparerDonnees`, `combinaisonNirvanaOptimale`, `correctionPariteFinale`, `V2_Ameliore_CalculerEtatGlobal` ;
* fournit une méthode `runCombination` réutilisée par `lancerCombinaisonNirvanaOptimale` et testable hors environnement Google.

## 3. Tests

* `tests/nirvanaEngine.test.js` vérifie le calcul du résumé (durées, score final, total d’opérations) et le pilotage du verrou/toasts/alertes.
* `tests/helpers/loadNirvanaSandbox.js` charge le script dans un `vm` Node avec des stubs de `SpreadsheetApp` et `LockService` pour les tests.

Ainsi, la logique métier (phases V2 + Parité) peut être instrumentée, testée et migrée vers d’autres environnements sans dépendre du runtime Apps Script.

# Architecture interne – ScoresEquilibrageEngine

La phase spécialisée « scores » dispose désormais d’un noyau dédié embarqué dans `Nirvana_Combined_Orchestrator.js`.

## 1. Domaine (`ScoresEquilibrageEngine.createDomain`)

* orchestre la séquence « stratégie spécialisée → fallback Nirvana V2 » ;
* normalise les sorties (`nbOperations`, stratégie retenue, cycles) et historise chaque tentative (`history`) pour faciliter l’analyse ;
* gère les erreurs en isolant les exceptions et en garantissant une réponse sérialisable pour l’orchestrateur.

## 2. Service (`ScoresEquilibrageEngine.createService`)

* prépare la configuration spécifique aux scores (colonnes actives, mode agressif, itérations maximales) sans muter l’entrée ;
* injecte les dépendances Apps Script (`executerEquilibrageScoresPersonnalise`, `V2_Ameliore_OptimisationEngine`) et les expose sous forme de callbacks testables ;
* enrichit le rapport final (config utilisée, scénarios) pour simplifier le débogage.

## 3. Tests

* `tests/scoresPhaseEngine.test.js` vérifie l’ordre des tentatives, la préparation de configuration et la remontée des métriques clés hors Apps Script.

