# SEPTEMBRE-25

## Architecture "backend élèves"

Le module `BackendV2.js` expose désormais un noyau `ElevesBackend` organisé par domaines :

* **Domain (`ElevesBackend.createDomain`)** – transforme les lignes des feuilles en objets élèves, résout les alias de colonnes et prépare les règles de structure (capacités, quotas).
* **Data access (`ElevesBackend.createDataAccess`)** – encapsule l’accès aux feuilles Google Sheets (sélection par suffixe, lecture de `_STRUCTURE`, filtrage des onglets techniques).
* **Service (`ElevesBackend.createService`)** – orchestre la lecture des données, applique le domaine et renvoie des objets sérialisables pour `getElevesData*` et `getStructureRules`.

Cette séparation permet d’injecter facilement des dépendances pour les tests tout en conservant les fonctions Apps Script historiques (`getElevesData`, `getElevesDataForMode`, `getStructureRules`, `getEleveById_`, etc.).

Une description détaillée figure dans [`docs/architecture.md`](docs/architecture.md).

## Architecture "Nirvana engine"

`Nirvana_Combined_Orchestrator.js` expose désormais un module `NirvanaEngine` structuré de façon similaire :

* **Domain (`NirvanaEngine.createDomain`)** – agrège les résultats des phases V2/Parité, calcule les durées et produit les messages standardisés (alertes, toasts).
* **Service (`NirvanaEngine.createService`)** – encapsule l’orchestrateur Apps Script : gestion du verrou `LockService`, interaction avec l’UI Google Sheets et injection des dépendances (préparation des données, phases métier).

La fonction `lancerCombinaisonNirvanaOptimale` se contente ainsi d’invoquer le service configuré, ce qui simplifie les tests et la réutilisation du moteur.

## Architecture "Nirvana data backend"

`Nirvana_DataBackend.js` factorise la préparation du contexte utilisé par Nirvana V2 et les variantes parité :

* **Domain (`NirvanaDataBackend.createDomain`)** – regroupe les élèves par classe, calcule les effectifs/cibles et génère les caches internes (`classeCaches`, `dissocMap`).
* **Service (`NirvanaDataBackend.createService`)** – orchestre la lecture des feuilles (`chargerElevesEtClasses_AvecSEXE`), la sanitation et la classification avant de remettre un `dataContext` unique aux moteurs historiques.

Les fonctions `V2_Ameliore_PreparerDonnees*` deviennent de simples wrappers vers ce service, ce qui garantit un comportement homogène entre Apps Script et les tests Node.

## Architecture "Scores phase engine"

`Nirvana_Combined_Orchestrator.js` embarque désormais `ScoresEquilibrageEngine`, un noyau léger dédié à la phase spécialisée « scores » :

* **Domain (`ScoresEquilibrageEngine.createDomain`)** – séquence l’exécution de la stratégie spécialisée et du fallback Nirvana V2, capitalise les tentatives dans un historique et homogénéise les métriques (`nbOperations`, stratégie retenue) pour l’orchestrateur.
* **Service (`ScoresEquilibrageEngine.createService`)** – prépare la configuration agressive (`COLONNES_SCORES_ACTIVES`, `MAX_ITERATIONS_SCORES`), injecte les dépendances Apps Script (`executerEquilibrageScoresPersonnalise`, `V2_Ameliore_OptimisationEngine`) et renvoie un rapport testable hors Google Sheets.

La fonction `executerPhaseScoresSpecialisee` délègue ainsi l’intégralité de la logique à ce moteur, ce qui simplifie l’instrumentation et le pilotage des différents scénarios d’équilibrage.

## Tests

Les tests unitaires Node utilisent des loaders sandbox (`tests/helpers/loadBackendSandbox.js`, `tests/helpers/loadNirvanaSandbox.js`) pour exécuter `BackendV2.js` et `Nirvana_Combined_Orchestrator.js` hors Apps Script :

```bash
node --test
```

La suite couvre désormais :

* le backend élèves (alias d’ID, sélection des suffixes) ;
* le data backend Nirvana (construction du `dataContext`, gestion des erreurs de classification) ;
* l’orchestrateur combiné (agrégation des phases, verrouillage, toasts/alertes) ;
* le moteur de phase scores (enchaînement stratégie spécialisée/fallback, préparation de configuration).
