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

## Tests

Les tests unitaires Node utilisent des loaders sandbox (`tests/helpers/loadBackendSandbox.js`, `tests/helpers/loadNirvanaSandbox.js`) pour exécuter `BackendV2.js` et `Nirvana_Combined_Orchestrator.js` hors Apps Script :

```bash
node --test
```

Les tests vérifient à la fois le backend élèves (alias d’ID) et l’orchestrateur Nirvana (agrégation des phases, gestion du verrou, interactions UI simulées).
