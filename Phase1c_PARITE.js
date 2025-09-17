// ================================================
// FONCTION D'ÉCRITURE SPÉCIFIQUE PHASE 3 (Inchangée - OK)
// ================================================
/**
 * Écrit les élèves placés par Phase 3, en mettant la raison
 * DANS LA COLONNE Q (index 17).
 * @param {Array<Object>} elevesAPlacerPhase3 Liste des élèves placés.
 * @return {Object} Résultat { success: boolean, nbAjoutes: number, message?: string }
 */
function ecrireElevesPhase3(elevesAPlacerPhase3) {
  // ... (Code inchangé - utilise déjà ID_ELEVE correctement) ...
  if (!elevesAPlacerPhase3 || !Array.isArray(elevesAPlacerPhase3)) {
    Logger.log("❌ ERREUR (ecrireElevesPhase3): Liste d'élèves invalide.");
    return { success: false, nbAjoutes: 0, message: "Liste d'élèves Phase 3 invalide." };
  }
  if (elevesAPlacerPhase3.length === 0) {
      Logger.log("ℹ️ (ecrireElevesPhase3): Aucun élève placé par Phase 3 à écrire.");
      return { success: true, nbAjoutes: 0 };
  }

  Logger.log(`Début écriture de ${elevesAPlacerPhase3.length} élèves (raison en Col Q)...`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalAjoutes = 0;
  const COLONNE_RAISON_INDEX = 16; // Index base 0 pour colonne Q (17 - 1)

  try {
    // 1. Lire CONSOLIDATION pour mapper ID -> Ligne complète
    const consolidationSheet = ss.getSheetByName('CONSOLIDATION');
    if (!consolidationSheet) throw new Error("Onglet CONSOLIDATION introuvable.");
    const consolidationData = consolidationSheet.getDataRange().getValues();
    const consolHeaders = consolidationData[0].map(h => String(h).trim());
    const consolIdIndex = consolHeaders.indexOf("ID_ELEVE"); // OK: Utilise ID_ELEVE
    if (consolIdIndex === -1) throw new Error("Colonne 'ID_ELEVE' introuvable dans CONSOLIDATION.");

    const ligneParId = {};
    for (let i = 1; i < consolidationData.length; i++) {
      const row = consolidationData[i];
      const id = String(row[consolIdIndex] || "").trim();
      if (id) ligneParId[id] = row;
    }

    // 2. Lire les IDs existants (pour sécurité) et les headers des feuilles TEST
    const idsExistantsParClasse = {};
    const headersParClasse = {};
    const testSheets = getTestSheetsNames();

    for (const classe of testSheets) {
        idsExistantsParClasse[classe] = new Set();
        const sheet = ss.getSheetByName(classe);
        if (!sheet) continue;
        const lastRow = sheet.getLastRow();
        if (lastRow > 0) {
            headersParClasse[classe] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
            const idIdx = headersParClasse[classe].indexOf("ID_ELEVE"); // OK: Vérifie ID_ELEVE pour la sécurité
            if (idIdx !== -1 && lastRow > 1) {
                const idsData = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();
                idsData.forEach(rowData => {
                    const id = String(rowData[0] || "").trim();
                    if (id) idsExistantsParClasse[classe].add(id);
                });
            }
        } else {
             headersParClasse[classe] = [...consolHeaders];
        }
    }


    // 3. Grouper les élèves à écrire par classe
    const elevesParClasse = {};
    elevesAPlacerPhase3.forEach(eleve => {
      if (!eleve.classe || !ligneParId[eleve.id]) {
          Logger.log(`⚠️ Données manquantes pour écrire élève ${eleve.id}`);
          return;
      }
      // Sécurité: Vérifie si l'ID existe DEJA dans la classe cible AVANT d'écrire
      if (idsExistantsParClasse[eleve.classe] && idsExistantsParClasse[eleve.classe].has(eleve.id)) {
           Logger.log(`ℹ️ Sécurité (ecrireElevesPhase3): Élève ${eleve.id} déjà détecté dans ${eleve.classe}. Écriture annulée pour cet élève.`);
           return; // Ne pas ajouter cet élève aux données à écrire
       }

      if (!elevesParClasse[eleve.classe]) elevesParClasse[eleve.classe] = [];

      const ligneComplete = [...ligneParId[eleve.id]];
      while (ligneComplete.length <= COLONNE_RAISON_INDEX) {
          ligneComplete.push("");
      }
      ligneComplete[COLONNE_RAISON_INDEX] = eleve.raison || "Phase 3 - Parité/Capacité";
      elevesParClasse[eleve.classe].push(ligneComplete);

      // Mettre à jour immédiatement l'ensemble des IDs pour la sécurité intra-boucle
      if(idsExistantsParClasse[eleve.classe]) {
        idsExistantsParClasse[eleve.classe].add(eleve.id);
      }

    });

    // 4. Écrire dans chaque onglet TEST concerné
    for (const classe in elevesParClasse) {
      const sheet = ss.getSheetByName(classe);
      if (!sheet) continue;

      const nouvellesDonnees = elevesParClasse[classe];
      if (nouvellesDonnees.length > 0) {
          let startRow = sheet.getLastRow() + 1;
          const numRows = nouvellesDonnees.length;
          let headersAEcrire = headersParClasse[classe];

          const dataNumCols = nouvellesDonnees[0].length;
          let targetNumCols = Math.max(sheet.getLastColumn() || 1, dataNumCols);

          if (sheet.getMaxColumns() < targetNumCols) {
              sheet.insertColumnsAfter(sheet.getMaxColumns(), targetNumCols - sheet.getMaxColumns());
          }

          if (startRow === 1) {
               while(headersAEcrire.length < targetNumCols) headersAEcrire.push("");
               sheet.getRange(1, 1, 1, targetNumCols).setValues([headersAEcrire]).setFontWeight('bold');
               startRow = 2;
          }

          try {
              sheet.getRange(startRow, 1, numRows, dataNumCols).setValues(nouvellesDonnees);
              Logger.log(`✅ ${numRows} élèves ajoutés à ${classe} (raison en Col Q).`);
              totalAjoutes += numRows;
          } catch (e) {
               Logger.log(`❌ Erreur écriture ${classe} L${startRow}: ${e.message}.`);
               throw new Error(`Erreur écriture ${classe}: ${e.message}`);
          }
      }
    }

    Logger.log(`✅ Écriture Phase 3 terminée: ${totalAjoutes} élèves ajoutés.`);
    return { success: true, nbAjoutes: totalAjoutes };

  } catch (error) {
    Logger.log(`❌ ERREUR Critique dans ecrireElevesPhase3: ${error.message} - ${error.stack}`);
    SpreadsheetApp.getUi().alert(`Erreur Critique Phase 3 Écriture: ${error.message}`);
    return { success: false, nbAjoutes: totalAjoutes, message: error.message };
  }
}


// ================================================
// FONCTIONS UTILITAIRES (Modifications mineures ou logs ajoutés)
// ================================================
function getTestSheetsNames() { /* ... code inchangé ... */
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const testPattern = /^[3-6]°\d+TEST$/;
  const testSheets = sheets
    .map(sheet => sheet.getName())
    .filter(name => testPattern.test(name));
  if (testSheets.length === 0) {
    Logger.log("⚠️ Aucun onglet TEST trouvé selon le pattern ^[3-6]°\\d+TEST$.");
    return [];
  }
  Logger.log("Onglets TEST trouvés: " + testSheets.join(", "));
  return testSheets;
 }
function lireStructureClassesAvecEffectifs() { /* ... code inchangé ... */
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("_STRUCTURE");
    const structure = {};
    const testSheets = getTestSheetsNames(); // Pour savoir quelles classes chercher

    // Initialisation très basique
    testSheets.forEach(c => {
        structure[c] = {
            classe: c.replace("TEST", ""),
            effectifCible: 26 // Valeur par défaut globale
        };
    });

    if (!sheet) {
        Logger.log("⚠️ Onglet _STRUCTURE non trouvé. Utilisation de l'effectif cible par défaut (26).");
        return structure;
    }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
        Logger.log("⚠️ Onglet _STRUCTURE vide. Utilisation de l'effectif cible par défaut (26).");
        return structure;
    }

    // Recherche dynamique des colonnes CLASSE_DEST et EFFECTIF
    let headerRowIndex = -1;
    let colClasseDest = -1, colEffectif = -1;
    for (let i = 0; i < Math.min(5, data.length); i++) {
        const headers = data[i].map(h => String(h).toUpperCase().trim());
        // Cherche explicitement CLASSE_DEST car c'est elle qui correspond aux onglets TEST
        const potentialColClasseDest = headers.indexOf("CLASSE_DEST");
        const potentialColEffectif = headers.indexOf("EFFECTIF");

        if (potentialColClasseDest !== -1 && potentialColEffectif !== -1) {
            headerRowIndex = i;
            colClasseDest = potentialColClasseDest;
            colEffectif = potentialColEffectif;
            break;
        }
    }

    if (headerRowIndex === -1) {
        Logger.log("⚠️ En-têtes 'CLASSE_DEST' et 'EFFECTIF' non trouvés dans _STRUCTURE. Utilisation des effectifs par défaut.");
        return structure;
    }

    // Lecture des données
    for (let i = headerRowIndex + 1; i < data.length; i++) {
        const row = data[i];
        const classeDest = String(row[colClasseDest] || "").trim(); // Ex: "5°1"
        const effectifCible = parseInt(row[colEffectif], 10);
        const classeTEST = classeDest + "TEST"; // Ex: "5°1TEST"

        // Met à jour seulement si la classe TEST existe et l'effectif est valide
        if (classeDest && structure[classeTEST] && !isNaN(effectifCible)) {
            structure[classeTEST].effectifCible = effectifCible;
        }
    }
    Logger.log("Structure des classes (Effectifs cibles) chargée: " + JSON.stringify(structure));
    return structure;
 }

/**
 * Lit TOUS les élèves de CONSOLIDATION et les élèves DÉJÀ PLACÉS dans les onglets TEST.
 * Fusionne les listes en marquant correctement `dejaPlace`.
 * **LOGS AJOUTÉS** pour tracer la fusion.
 * @return {Array<Object>} Liste de tous les élèves avec statut `dejaPlace`.
 */
function lireConsolidationAvecGenre() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Étape 1: Lire les élèves DÉJÀ dans les onglets TEST
  // Cette fonction DOIT maintenant utiliser ID_ELEVE correctement.
  const elevesExistants = lireElevesExistants(); // Appel à la version corrigée ci-dessous
  Logger.log(`[Fusion Step 1] ${elevesExistants.length} élèves trouvés dans les onglets TEST (marqués dejaPlace=true).`);

  // Créer une Map pour un accès rapide aux élèves existants par ID
  const elevesMap = {};
  elevesExistants.forEach(e => { if (e.id) { elevesMap[e.id] = e; } });
  Logger.log(`[Fusion Step 1b] ${Object.keys(elevesMap).length} élèves uniques trouvés dans les onglets TEST et mappés par ID.`);

  // Étape 2: Lire TOUS les élèves depuis CONSOLIDATION
  const sheet = ss.getSheetByName("CONSOLIDATION");
  const elevesConsolidation = [];
  const idsConsolidation = new Set(); // Pour éviter doublons de CONSO

  if (!sheet) {
    Logger.log("⚠️ Onglet CONSOLIDATION introuvable.");
    // On continue avec seulement les élèves des onglets TEST si CONSO manque
  } else {
      const data = sheet.getDataRange().getValues();
      if (data.length > 1) {
          const headers = data[0].map(h => String(h).trim());
          const idIndex = headers.indexOf("ID_ELEVE"); // Crucial: doit être ID_ELEVE
          const nomPrenomIndex = headers.indexOf("NOM & PRENOM");
          const genreIndex = headers.indexOf("SEXE"); // Crucial
          const optIndex = headers.indexOf("OPT");
          const lv2Index = headers.indexOf("LV2");
          const assoIndex = headers.indexOf("ASSO");
          const dissoIndex = headers.indexOf("DISSO");

          if (idIndex === -1) {
              Logger.log("❌ ERREUR: Colonne 'ID_ELEVE' requise manquante dans CONSOLIDATION.");
              throw new Error("Colonne 'ID_ELEVE' manquante dans CONSOLIDATION.");
          } else if (genreIndex === -1) {
               Logger.log("❌ ERREUR: Colonne 'SEXE' requise manquante dans CONSOLIDATION. La parité Phase 3 ne peut pas fonctionner.");
               throw new Error("Colonne 'SEXE' manquante dans CONSOLIDATION, Phase 3 impossible.");
          } else {
              for (let i = 1; i < data.length; i++) {
                  const row = data[i];
                  const id = String(row[idIndex] || "").trim();

                  // Ignore les lignes sans ID ou les IDs déjà lus de CONSO
                  if (!id || idsConsolidation.has(id)) {
                    if(idsConsolidation.has(id)) Logger.log(`   [CONSO Read] Skipping duplicate ID ${id} within CONSOLIDATION.`);
                    continue;
                  }

                  const genre = String(row[genreIndex] || "").toUpperCase().trim();
                  const options = (optIndex !== -1) ? String(row[optIndex] || "").toUpperCase().trim() : "";
                  const lv2 = (lv2Index !== -1) ? String(row[lv2Index] || "").toUpperCase().trim() : "";

                  // Crée l'objet élève depuis CONSO, avec dejaPlace=false PAR DEFAUT
                  const eleve = {
                      id: id,
                      nomPrenom: (nomPrenomIndex !== -1) ? String(row[nomPrenomIndex] || "") : `Élève ${id}`,
                      genre: genre,
                      isGarcon: genre === "M",
                      isFille: genre === "F",
                      options: options,
                      lv2: lv2,
                      assoCode: (assoIndex !== -1) ? String(row[assoIndex] || "").trim() : "",
                      dissoCode: (dissoIndex !== -1) ? String(row[dissoIndex] || "").trim() : "",
                      classe: null, // Sera défini si déjà placé ou par la répartition
                      raison: "",   // Sera défini si déjà placé ou par la répartition
                      dejaPlace: false // Important: Initialisé à false
                  };
                  elevesConsolidation.push(eleve);
                  idsConsolidation.add(id);
              }
          }
      }
      Logger.log(`[Fusion Step 2] ${elevesConsolidation.length} élèves lus depuis CONSOLIDATION (marqués initialement dejaPlace=false).`);
  }

  // Étape 3: Fusionner les listes
  Logger.log("[Fusion Step 3] Début de la fusion CONSOLIDATION + TEST...");
  const elevesFinaux = [];
  const idsTraitesFusion = new Set(); // Pour suivre les IDs traités lors de la fusion

  // Boucle 1: Parcourir les élèves de CONSOLIDATION
  elevesConsolidation.forEach(eConso => {
      if (elevesMap[eConso.id]) {
          // Cet élève de CONSO existe DÉJÀ dans un onglet TEST.
          // On prend la version de TEST (qui a dejaPlace=true et potentiellement une classe/raison).
          const eleveExistant = elevesMap[eConso.id];
          elevesFinaux.push(eleveExistant);
          Logger.log(`   [Fusion] ID ${eConso.id} (${eConso.nomPrenom}): Trouvé dans TEST (${eleveExistant.classe}), marqué dejaPlace=true.`);
      } else {
          // Cet élève de CONSO N'existe PAS dans les onglets TEST.
          // On prend la version de CONSO (qui a dejaPlace=false). Il est à placer.
          elevesFinaux.push(eConso);
           Logger.log(`   [Fusion] ID ${eConso.id} (${eConso.nomPrenom}): Non trouvé dans TEST, marqué dejaPlace=false (à placer).`);
      }
      idsTraitesFusion.add(eConso.id); // Marquer cet ID comme traité
  });

  // Boucle 2: Ajouter les élèves des onglets TEST qui n'étaient PAS dans CONSOLIDATION
  // (Cas où un élève est dans un TEST mais absent de la liste CONSO)
  elevesExistants.forEach(eTest => {
      if (eTest.id && !idsTraitesFusion.has(eTest.id)) {
          // Cet élève trouvé dans TEST n'était pas dans la liste CONSO.
          // On l'ajoute (il est déjà marqué dejaPlace=true).
          elevesFinaux.push(eTest);
          Logger.log(`   [Fusion] ID ${eTest.id} (${eTest.nomPrenom}): Trouvé dans TEST (${eTest.classe}) mais absent de CONSO. Ajouté (dejaPlace=true).`);
          idsTraitesFusion.add(eTest.id); // Marquer comme traité
      }
  });

  Logger.log(`[Fusion Step 4] Fusion terminée: ${elevesFinaux.length} élèves uniques au total.`);

  // Vérification finale: Compter combien sont marqués dejaPlace vs non
  const comptesFinaux = elevesFinaux.reduce((acc, e) => {
      acc[e.dejaPlace ? 'places' : 'aPlacer']++;
      return acc;
  }, { places: 0, aPlacer: 0 });
  Logger.log(`[Fusion Result] ${comptesFinaux.places} élèves marqués comme déjà placés, ${comptesFinaux.aPlacer} élèves marqués comme à placer.`);

  return elevesFinaux;
}

/**
 * Lit les élèves DÉJÀ PRÉSENTS dans les onglets TEST.
 * **CORRIGÉ** pour utiliser "ID_ELEVE".
 * @return {Array<Object>} Liste des élèves trouvés dans les onglets TEST, marqués `dejaPlace: true`.
 */
function lireElevesExistants() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheets = getTestSheetsNames();
  const elevesExistants = [];
  const idsTraites = new Set(); // Pour éviter les doublons si un élève est dans plusieurs TESTs (erreur)

  Logger.log(`--- Début lecture élèves existants dans ${testSheets.length} onglets TEST... ---`);

  for (const sheetName of testSheets) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`   Skipping non-existent sheet: ${sheetName}`);
      continue;
    }
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log(`   Skipping empty or header-only sheet: ${sheetName}`);
      continue;
    }

    Logger.log(`   Lecture de ${sheetName} (jusqu'à la ligne ${lastRow})...`);
    let countSheet = 0;
    try {
        const data = sheet.getRange(1, 1, lastRow, sheet.getMaxColumns()).getValues();
        const headers = data[0].map(h => String(h).trim());

        // ***** CORRECTION CRUCIALE ICI *****
        const idIdx = headers.indexOf("ID_ELEVE");
        // ***********************************

        const nomIdx = headers.indexOf("NOM & PRENOM");
        const genreIdx = headers.indexOf("SEXE");
        const optIdx = headers.indexOf("OPT");
        const lv2Idx = headers.indexOf("LV2");
        const raisonIdx = headers.indexOf("RAISON PLACEMENT");
        const classeFinaleIdx = headers.indexOf("CLASSE_FINALE"); // Colonne Q

        if (idIdx === -1) {
            Logger.log(`   ⚠️ Colonne 'ID_ELEVE' non trouvée dans ${sheetName}. Impossible de lire les élèves existants de cet onglet.`);
            continue; // Passer à l'onglet suivant si pas d'ID_ELEVE
        }
        if (genreIdx === -1) { Logger.log(`   ⚠️ Colonne 'SEXE' non trouvée dans ${sheetName} lors de la lecture des existants.`); }

        for (let i = 1; i < data.length; i++) { // Commence à la ligne 2 (index 1)
            const row = data[i];
            const id = String(row[idIdx] || "").trim();

            // Ignore les lignes sans ID ou les IDs déjà traités (venant d'un autre onglet TEST)
            if (!id) continue;
            if (idsTraites.has(id)) {
                Logger.log(`   ⚠️ ID_ELEVE ${id} déjà lu depuis un autre onglet TEST. Ignoré dans ${sheetName}.`);
                continue;
            }

            const genre = (genreIdx !== -1) ? String(row[genreIdx] || "").toUpperCase().trim() : "";
            const options = (optIdx !== -1) ? String(row[optIdx] || "").toUpperCase().trim() : "";
            const lv2 = (lv2Idx !== -1) ? String(row[lv2Idx] || "").toUpperCase().trim() : "";

            // Lire la raison: Priorité à RAISON_PLACEMENT, fallback sur CLASSE_FINALE (Q), puis défaut.
            let raison = (raisonIdx !== -1) ? String(row[raisonIdx] || "").trim() : "";
            if (!raison && classeFinaleIdx !== -1) {
                 raison = String(row[classeFinaleIdx] || "").trim(); // Prend le contenu de Q si l'autre est vide
                 if (raison) raison = `Col Q: ${raison}`; // Préfixer pour savoir d'où elle vient
            }
            if (!raison) raison = "Placé précedemment (Raison non spécifiée)"; // Défaut si toujours vide

            elevesExistants.push({
                id: id,
                nomPrenom: (nomIdx !== -1) ? String(row[nomIdx] || "") : `Élève ${id}`,
                genre: genre,
                isGarcon: genre === "M",
                isFille: genre === "F",
                options: options,
                lv2: lv2,
                classe: sheetName, // L'élève est DANS CETTE CLASSE
                raison: raison,    // Raison de sa présence ici
                dejaPlace: true    // MARQUEUR ESSENTIEL
            });
            idsTraites.add(id);
            countSheet++;
        }
        Logger.log(`   ${countSheet} élèves uniques lus depuis ${sheetName}.`);

    } catch(e) { Logger.log(`   ❌ Erreur lecture élèves existants ${sheetName}: ${e.message}`); }
  }
  Logger.log(`--- Fin lecture élèves existants. Total: ${elevesExistants.length} élèves uniques trouvés. ---`);
  return elevesExistants;
}

function calculerStatsClasses(elevesPlaces, structureClasses) { /* ... code inchangé ... */
    const stats = {};
    const testSheets = Object.keys(structureClasses); // Utiliser les clés de la structure chargée

    // Initialiser les statistiques
    testSheets.forEach(c => {
        stats[c] = {
            total: 0,
            garcons: 0,
            filles: 0,
            autres: 0, // Compte ceux sans 'M' ou 'F' ou si colonne SEXE absente
            effectifCible: structureClasses[c].effectifCible,
            placesRestantes: structureClasses[c].effectifCible
        };
    });

    // Calculer les statistiques actuelles
    let missingSexColumnWarning = false;
    elevesPlaces.forEach(e => {
        // Important: On ne compte QUE les élèves qui ont une classe assignée
        if (e.classe && stats[e.classe]) {
            stats[e.classe].total++;
            if (e.genre === "M") stats[e.classe].garcons++;
            else if (e.genre === "F") stats[e.classe].filles++;
            else {
                 stats[e.classe].autres++;
                 if (e.genre === "" && !missingSexColumnWarning) {
                     missingSexColumnWarning = true; // Loggué une seule fois si besoin
                 }
            }
        } else if (e.classe) {
            // Cas étrange : un élève a une classe assignée mais elle n'est pas dans la structure ?
             Logger.log(`⚠️ Élève ${e.id} assigné à la classe ${e.classe} qui n'est pas dans la structure chargée.`);
        }
        // Si e.classe est null, l'élève n'est pas placé, on ne le compte pas ici.
    });

     if (missingSexColumnWarning) {
       Logger.log("⚠️ Au moins un élève compté dans les stats a un genre vide ou non défini ('M'/'F'). Il est compté dans 'autres'.");
     }

    // Calculer les places restantes
    for (const c in stats) {
        stats[c].placesRestantes = stats[c].effectifCible - stats[c].total;
        // On ne met plus placesRestantes à 0 si négatif ici,
        // car la logique de placement doit pouvoir voir le sureffectif.
        if (stats[c].placesRestantes < 0) {
            Logger.log(`⚠️ Classe ${c} en sureffectif (${stats[c].total}/${stats[c].effectifCible}) AVANT Phase 3.`);
        }
    }

    Logger.log("Statistiques initiales des classes (Total/Genre/Places restantes): " + JSON.stringify(stats));
    return stats;
 }


// ================================================
// LOGIQUE DE RÉPARTITION PHASE 3 (Inchangée - OK)
// ================================================
/**
 * Répartit les élèves restants en respectant la parité et les contraintes d'options/langues
 * @param {Array} elevesAPlacerUniquement Ceux filtrés avec dejaPlace === false
 * @param {Object} statsClasses Statistiques actuelles des classes
 * @param {Object} structureClasses Structure des classes avec effectifs cibles
 * @return {Object} Résultat de répartition { elevesPlaces: Array<Object>, statsFinales: Object }
 */
/**
 * Répartit les élèves restants en respectant la parité, les contraintes d'options/langues
 * ET en essayant de maintenir les groupes ASSO ensemble.
 * @param {Array} elevesAPlacerUniquement Ceux filtrés avec dejaPlace === false
 * @param {Object} statsClasses Statistiques actuelles des classes
 * @param {Object} structureClasses Structure des classes avec effectifs cibles
 * @return {Object} Résultat de répartition { elevesPlaces: Array<Object>, statsFinales: Object }
 */
function repartirElevesRestantsAvecParite(elevesAPlacerUniquement, statsClasses, structureClasses) {
    Logger.log(`--- Début Répartition Phase 3 pour ${elevesAPlacerUniquement.length} élèves (Priorité Groupes ASSO puis Parité/Capacité) ---`);

    const elevesPlacesParCettePhase = [];
    let statsSimulees = JSON.parse(JSON.stringify(statsClasses));
    const elevesRestantsPourIndividuel = [...elevesAPlacerUniquement]; // Copie pour manipuler

    // --- Étape 1: Identifier et Grouper les élèves avec code ASSO ---
    const groupesAsso = {};
    const idsElevesDansGroupesAsso = new Set();

    elevesRestantsPourIndividuel.forEach(eleve => {
        const assoCode = eleve.assoCode ? String(eleve.assoCode).trim() : "";
        if (assoCode) {
            if (!groupesAsso[assoCode]) {
                groupesAsso[assoCode] = {
                    membres: [],
                    taille: 0,
                    filles: 0,
                    garcons: 0,
                    autres: 0,
                    code: assoCode
                };
            }
            groupesAsso[assoCode].membres.push(eleve);
            groupesAsso[assoCode].taille++;
            if (eleve.isFille) groupesAsso[assoCode].filles++;
            else if (eleve.isGarcon) groupesAsso[assoCode].garcons++;
            else groupesAsso[assoCode].autres++;
            idsElevesDansGroupesAsso.add(eleve.id); // Marquer cet élève comme faisant partie d'un groupe traité
        }
    });

    const nbGroupesAsso = Object.keys(groupesAsso).length;
    if (nbGroupesAsso > 0) {
       Logger.log(`   Identifié ${nbGroupesAsso} groupes ASSO à traiter en priorité.`);
    } else {
       Logger.log(`   Aucun groupe ASSO à traiter en priorité.`);
    }


    // --- Étape 2: Traiter les Groupes ASSO ---
    const groupesAssoPlaces = new Set(); // Pour savoir quels groupes ont été placés

    for (const codeAsso in groupesAsso) {
        const groupe = groupesAsso[codeAsso];
        Logger.log(`   Traitement Groupe ASSO: ${codeAsso} (Taille: ${groupe.taille}, ${groupe.filles}F/${groupe.garcons}G/${groupe.autres}X)`);

        let meilleureClassePourGroupe = null;
        let meilleurScorePourGroupe = -Infinity;

        // Trouver les classes candidates (assez de place pour le groupe entier)
        const classesCandidates = Object.keys(statsSimulees).filter(classeNom => {
           return statsSimulees[classeNom] && statsSimulees[classeNom].placesRestantes >= groupe.taille;
        });

        if (classesCandidates.length === 0) {
            Logger.log(`      ⚠️ Aucune classe avec suffisamment de place (${groupe.taille}) pour le groupe ASSO ${codeAsso}.`);
            continue; // Passer au groupe suivant
        }

        // Évaluer les classes candidates pour ce groupe
        for (const classeNom of classesCandidates) {
            const classeStats = statsSimulees[classeNom];
            let score = 0;

            // Critère 1: Parité (impact de l'ajout du groupe entier)
            const totalApres = classeStats.total + groupe.taille;
            const fillesApres = classeStats.filles + groupe.filles;
            const garconsApres = classeStats.garcons + groupe.garcons;
            const ratioIdeal = 0.5;

            let ratioGarconsAvant = classeStats.total > 0 ? classeStats.garcons / classeStats.total : ratioIdeal;
            let ratioFillesAvant = classeStats.total > 0 ? classeStats.filles / classeStats.total : ratioIdeal;
            // Nouveaux ratios G/F après ajout du groupe (ignorant 'autres' pour le ratio)
            let denomRatioApres = fillesApres + garconsApres; // Dénominateur pour ratio G/F
             let ratioGarconsApres = denomRatioApres > 0 ? garconsApres / denomRatioApres : ratioIdeal;
             let ratioFillesApres = denomRatioApres > 0 ? fillesApres / denomRatioApres : ratioIdeal;


            const ecartAvant = Math.abs(ratioGarconsAvant - ratioFillesAvant);
            const ecartApres = Math.abs(ratioGarconsApres - ratioFillesApres);
            score += (ecartAvant - ecartApres) * 100; // Amélioration parité

            // Critère 2: Capacité restante (favorise classes plus vides)
            score += classeStats.placesRestantes * 10;

            // Logger.log(`      [Eval Groupe ${codeAsso}] Classe ${classeNom}: Score Parité=${(ecartAvant - ecartApres) * 100}, Score Capacité=${classeStats.placesRestantes * 10}, Total=${score}`);

            if (score > meilleurScorePourGroupe) {
                meilleurScorePourGroupe = score;
                meilleureClassePourGroupe = classeNom;
            }
        }

        // Placer le groupe si une classe a été trouvée
        if (meilleureClassePourGroupe) {
            Logger.log(`      ✅ Placement Groupe ASSO ${codeAsso} dans ${meilleureClassePourGroupe} (Score: ${meilleurScorePourGroupe}).`);
            const raisonPlacement = `Phase 3 - Groupe ASSO ${codeAsso}`;
            groupe.membres.forEach(membre => {
                membre.classe = meilleureClassePourGroupe;
                membre.raison = raisonPlacement;
                elevesPlacesParCettePhase.push(membre);
            });
            // Mettre à jour les stats simulées pour la classe choisie
            statsSimulees[meilleureClassePourGroupe].total += groupe.taille;
            statsSimulees[meilleureClassePourGroupe].filles += groupe.filles;
            statsSimulees[meilleureClassePourGroupe].garcons += groupe.garcons;
            statsSimulees[meilleureClassePourGroupe].autres += groupe.autres;
            statsSimulees[meilleureClassePourGroupe].placesRestantes -= groupe.taille;
            groupesAssoPlaces.add(codeAsso); // Marquer ce groupe comme placé

             Logger.log(`         Stats ${meilleureClassePourGroupe} après ajout groupe: ${statsSimulees[meilleureClassePourGroupe].total}/${statsSimulees[meilleureClassePourGroupe].effectifCible}, ${statsSimulees[meilleureClassePourGroupe].filles}F/${statsSimulees[meilleureClassePourGroupe].garcons}G/${statsSimulees[meilleureClassePourGroupe].autres}X, reste ${statsSimulees[meilleureClassePourGroupe].placesRestantes}`);

        } else {
             Logger.log(`      ⚠️ Le groupe ASSO ${codeAsso} n'a pas pu être placé (pas de classe optimale trouvée parmi les candidates).`);
        }
    }


    // --- Étape 3: Traiter les élèves individuels restants ---
    // Filtrer pour ne garder que ceux qui n'ont PAS de code ASSO OU dont le groupe ASSO n'a PAS été placé
    const elevesIndividuelsAPlacer = elevesRestantsPourIndividuel.filter(eleve => {
         const assoCode = eleve.assoCode ? String(eleve.assoCode).trim() : "";
         // Garder si pas de code ASSO OU si le code ASSO existe mais n'est PAS dans groupesAssoPlaces
         return !assoCode || (assoCode && !groupesAssoPlaces.has(assoCode));
    });


    if (elevesIndividuelsAPlacer.length > 0) {
       Logger.log(`--- Traitement de ${elevesIndividuelsAPlacer.length} élèves individuels restants (sans ASSO ou groupe non placé) ---`);
    } else if (nbGroupesAsso > 0) {
        Logger.log(`--- Tous les élèves restants faisaient partie de groupes ASSO (placés ou non). Aucun placement individuel nécessaire. ---`);
    } else {
        Logger.log(`--- Aucun élève individuel à placer. ---`);
    }


    const fillesIndiv = elevesIndividuelsAPlacer.filter(e => e.isFille);
    const garconsIndiv = elevesIndividuelsAPlacer.filter(e => e.isGarcon);
    const autresIndiv = elevesIndividuelsAPlacer.filter(e => !e.isGarcon && !e.isFille);
    Logger.log(`   Individuels à placer: ${garconsIndiv.length} G / ${fillesIndiv.length} F / ${autresIndiv.length} Autres`);


    /* Copier ici la fonction interne trouverClasseOptimale de la version précédente
       (celle qui évalue pour UN SEUL élève) */
    function trouverClasseOptimaleIndividuel(eleve, currentStats) {
       let meilleureClasse = null;
       let meilleurScore = -Infinity;

       // Lister les classes où il y a de la place
        const classesPossibles = Object.keys(currentStats).filter(c => currentStats[c] && currentStats[c].placesRestantes > 0);


       if (classesPossibles.length === 0) return null;

       for (const classeNom of classesPossibles) {
           const classeStats = currentStats[classeNom];
           let score = 0;
           const totalApres = classeStats.total + 1;
           const ratioIdeal = 0.5;

           let ratioGarconsAvant = classeStats.total > 0 ? classeStats.garcons / classeStats.total : ratioIdeal;
           let ratioFillesAvant = classeStats.total > 0 ? classeStats.filles / classeStats.total : ratioIdeal;
           let ratioGarconsApres = classeStats.garcons / totalApres;
           let ratioFillesApres = classeStats.filles / totalApres;

           if (eleve.isGarcon) ratioGarconsApres = (classeStats.garcons + 1) / totalApres;
           else if (eleve.isFille) ratioFillesApres = (classeStats.filles + 1) / totalApres;
           else { /* autres ne changent pas G/F pour calcul écart */ }

           const ecartAvant = Math.abs(ratioGarconsAvant - ratioFillesAvant);
           const ecartApres = Math.abs(ratioGarconsApres - ratioFillesApres);
           score += (ecartAvant - ecartApres) * 100; // Poids fort parité

           score += classeStats.placesRestantes * 10; // Poids faible capacité

           if (score > meilleurScore) {
               meilleurScore = score;
               meilleureClasse = classeNom;
           }
       }
       return meilleureClasse;
    }

    // Placement Alterné Filles / Garçons (Individuels)
    let i = 0, j = 0;
    let placedInIterationIndiv = true;

    while (placedInIterationIndiv && (i < fillesIndiv.length || j < garconsIndiv.length)) {
        placedInIterationIndiv = false;

        if (i < fillesIndiv.length) {
            const fille = fillesIndiv[i];
             const classeChoisie = trouverClasseOptimaleIndividuel(fille, statsSimulees);
             if (classeChoisie) {
                 fille.classe = classeChoisie;
                 fille.raison = "Phase 3 - Parité/Capacité (Indiv)";
                 elevesPlacesParCettePhase.push(fille);
                 statsSimulees[classeChoisie].total++; statsSimulees[classeChoisie].filles++;
                 statsSimulees[classeChoisie].placesRestantes--;
                 placedInIterationIndiv = true;
                 Logger.log(`   [Placement Indiv OK] ${fille.id} (F) placée dans ${classeChoisie}.`);
             }
             i++;
        }
        if (j < garconsIndiv.length) {
            const garcon = garconsIndiv[j];
            const classeChoisie = trouverClasseOptimaleIndividuel(garcon, statsSimulees);
            if (classeChoisie) {
                garcon.classe = classeChoisie;
                garcon.raison = "Phase 3 - Parité/Capacité (Indiv)";
                elevesPlacesParCettePhase.push(garcon);
                statsSimulees[classeChoisie].total++; statsSimulees[classeChoisie].garcons++;
                statsSimulees[classeChoisie].placesRestantes--;
                placedInIterationIndiv = true;
                 Logger.log(`   [Placement Indiv OK] ${garcon.id} (G) placé dans ${classeChoisie}.`);
            }
            j++;
        }
    }

    // Placement des 'Autres' (Individuels)
     Logger.log(`   Traitement des ${autresIndiv.length} élèves 'Autres' individuels...`);
     autresIndiv.forEach(eleve => {
         const classeChoisie = trouverClasseOptimaleIndividuel(eleve, statsSimulees);
         if (classeChoisie) {
             eleve.classe = classeChoisie;
             eleve.raison = "Phase 3 - Remplissage (Indiv)";
             elevesPlacesParCettePhase.push(eleve);
             statsSimulees[classeChoisie].total++; statsSimulees[classeChoisie].autres++;
             statsSimulees[classeChoisie].placesRestantes--;
             placedInIterationIndiv = true; // Marquer qu'on a potentiellement placé
             Logger.log(`   [Placement Indiv OK] ${eleve.id} (Autre) placé dans ${classeChoisie}.`);
         } else {
             Logger.log(`   [Placement Indiv Echec] Élève ${eleve.id} (${eleve.nomPrenom}) sans genre spécifié n'a pas pu être placé.`);
         }
     });


    // --- Bilan Final ---
    const tousLesElevesInitialement = elevesAPlacerUniquement;
    const elevesNonPlacesFinal = tousLesElevesInitialement.filter(e => !e.classe); // Ceux qui n'ont pas de classe assignée à la fin

    if (elevesNonPlacesFinal.length > 0) {
        Logger.log(`⚠️ ATTENTION: ${elevesNonPlacesFinal.length} élèves sur ${tousLesElevesInitialement.length} initialement à placer n'ont pas pu être affectés :`);
        elevesNonPlacesFinal.forEach(e => {
             const assoCode = e.assoCode ? ` (ASSO: ${e.assoCode})` : "";
             Logger.log(`   - ${e.id} (${e.nomPrenom})${assoCode}`);
        });
         Logger.log(`   Raisons possibles: Groupes ASSO trop grands pour les places restantes, ou manque de place général même pour les individus.`);
    }

    Logger.log(`--- Fin Répartition Phase 3: ${elevesPlacesParCettePhase.length} élèves placés au total par cette phase ---`);
    Logger.log("Statistiques simulées APRES Phase 3:");
    for (const c in statsSimulees) {
        Logger.log(`   ${c}: ${statsSimulees[c].total}/${statsSimulees[c].effectifCible}, ${statsSimulees[c].filles}F/${statsSimulees[c].garcons}G/${statsSimulees[c].autres}X, reste ${statsSimulees[c].placesRestantes}`);
    }

    return { elevesPlaces: elevesPlacesParCettePhase, statsFinales: statsSimulees };
}


// ================================================
// FONCTION D'ORCHESTRATION PHASE 3 (Logique) (Adaptée)
// ================================================
function repartirPhase3() {
    Logger.log("===== DÉMARRAGE LOGIQUE PHASE 3 (Strict Capacité/Parité) =====");
    SpreadsheetApp.getActiveSpreadsheet().toast("Phase 3 : Lecture et identification des élèves...", "En cours...", 5);

    try {
        // 1. Lire la structure des classes (effectifs cibles)
        const structureClasses = lireStructureClassesAvecEffectifs();
        if (Object.keys(structureClasses).length === 0) {
          throw new Error("Aucune classe TEST valide trouvée/configurée dans _STRUCTURE.");
        }

        // 2. Lire TOUS les élèves (CONSO + TEST) et déterminer qui est 'dejaPlace'
        // C'est ici que la correction sur ID_ELEVE est cruciale.
        const tousEleves = lireConsolidationAvecGenre(); // Utilise la version avec logs améliorés

        // 3. Filtrer pour obtenir UNIQUEMENT les élèves à placer (ceux marqués dejaPlace = false)
        const elevesAPlacer = tousEleves.filter(e => !e.dejaPlace);
        Logger.log(`---> ${elevesAPlacer.length} élèves identifiés comme NON placés et à traiter par Phase 3.`);

        if (elevesAPlacer.length === 0) {
            Logger.log("Aucun élève non placé trouvé. Phase 3 logique terminée sans action.");
            SpreadsheetApp.getActiveSpreadsheet().toast("Aucun nouvel élève à placer.", "Phase 3 Info", 5);
            // Important: retourner un succès même si rien n'est fait.
            return { success: true, nbAjoutes: 0, message: "Aucun nouvel élève à placer." };
        }

        // 4. Filtrer pour obtenir les élèves DÉJÀ placés (pour calculer les stats initiales)
        const elevesDejaPlaces = tousEleves.filter(e => e.dejaPlace);
        Logger.log(`---> ${elevesDejaPlaces.length} élèves identifiés comme DÉJÀ placés.`);

        // 5. Calculer les statistiques initiales basées sur les élèves déjà placés
        SpreadsheetApp.getActiveSpreadsheet().toast("Phase 3 : Calcul stats initiales...", "En cours...", 5);
        const statsInitiales = calculerStatsClasses(elevesDejaPlaces, structureClasses);

        // 6. Lancer la répartition pour les élèves NON PLACÉS
        SpreadsheetApp.getActiveSpreadsheet().toast(`Phase 3 : Répartition de ${elevesAPlacer.length} élèves...`, "En cours...", 10);
        const resultatRepartition = repartirElevesRestantsAvecParite(elevesAPlacer, statsInitiales, structureClasses);
        const elevesPlacesParPhase3 = resultatRepartition.elevesPlaces; // Contient SEULEMENT les élèves placés par cette phase

        Logger.log(`---> ${elevesPlacesParPhase3.length} élèves ont été placés par la logique de Phase 3.`);

        // 7. Écrire UNIQUEMENT les élèves placés par cette Phase 3
        if (elevesPlacesParPhase3.length > 0) {
           SpreadsheetApp.getActiveSpreadsheet().toast(`Phase 3 : Écriture de ${elevesPlacesParPhase3.length} élèves...`, "En cours...", 5);
           const resultatEcriture = ecrireElevesPhase3(elevesPlacesParPhase3); // Appelle la version qui écrit en Q

           if (!resultatEcriture.success) {
               // Si l'écriture échoue, logguer et remonter l'erreur
               throw new Error(resultatEcriture.message || "Erreur lors de l'écriture des élèves de Phase 3.");
           }
           Logger.log(`Phase 3 logique terminée. ${resultatEcriture.nbAjoutes} élèves ajoutés avec succès.`);
           SpreadsheetApp.getActiveSpreadsheet().toast(`${resultatEcriture.nbAjoutes} élèves ajoutés par Phase 3.`, "Phase 3 Réussie", 7);
           return { success: true, nbAjoutes: resultatEcriture.nbAjoutes };

        } else {
            Logger.log("Phase 3 logique terminée. Aucun élève n'a pu être placé lors de la répartition (ou aucun à placer initialement).");
            SpreadsheetApp.getActiveSpreadsheet().toast("Aucun élève placé par Phase 3 (manque de place ou déjà complet).", "Phase 3 Info", 7);
            return { success: true, nbAjoutes: 0, message: "Aucun élève placé par cette phase." };
        }

    } catch (error) {
        Logger.log(`❌ ERREUR Critique pendant la logique de Phase 3: ${error.message} - ${error.stack}`);
        SpreadsheetApp.getUi().alert(`Erreur Critique Phase 3 Logique: ${error.message}`);
        // Renvoyer un échec
        return { success: false, nbAjoutes: 0, message: error.message };
    } finally {
        Logger.log("===== FIN LOGIQUE PHASE 3 =====");
    }
}

/**
 * Cache une colonne (ou plusieurs) uniquement si :
 * - la feuille n'est pas null
 * - la colonne de départ existe
 * - le nombre total de colonnes couvre toute la plage
 */
function safeHideColumns(sheet, startCol, howMany = 1) {
  if (!sheet) return;
  const max = sheet.getMaxColumns();
  if (startCol >= 1 && startCol + howMany - 1 <= max) {
    try {
      sheet.hideColumns(startCol, howMany);
    } catch(e) {
      Logger.log(`warn safeHideColumns(${startCol},${howMany}): ${e.message}`);
    }
  }
}

/**
 * Sécurise les appels aux fonctions de Sheet (getRange, setWidth, etc.)
 */
function safeSetWidth(sheet, colIdx, width) {
  if (!sheet || colIdx < 1) return;
  try {
    if (colIdx <= sheet.getMaxColumns()) {
      sheet.setColumnWidth(colIdx, width);
    }
  } catch(e) {
    Logger.log(`warn safeSetWidth(${colIdx},${width}): ${e.message}`);
  }
}

/*  PATCH 10‑05‑2025 — palette EXACTE maquette + en‑tête vert clair uniforme  */

/**
 * SCRIPT DE MISE EN FORME AUTOMATISÉE DES ONGLETS « TEST »
 * Version : 3.3 (maquette fidèle)
 * Auteur  : [Votre Nom/Équipe]
 *
 * Description :
 *   Formate automatiquement tous les onglets dont le nom se termine par « TEST »
 *   – Forçage de « ID_ELEVE » en colonne A.
 *   – Masquage/gel de colonnes et ajustement des largeurs.
 *   – Application de formats conditionnels (sexe, LV2, options, scores).
 *   – Zébrures, ligne de séparation et filtres.
 *   – Validation dynamique « CLASSE DEF ».
 *   – Calcul et écriture de statistiques.
 *   – Nettoyage des lignes/colonnes inutilisées.
 *
 * Palette conforme à la capture fournie :
 *   • En‑têtes : vert clair #C6E0B4
 *   • SEXE :   M #4F81BD | F #F28EA8
 *   • LV2 ESP #E59838 (autres variantes ci‑dessous)
 *   • Scores : 4 #006400 | 3 #3CB371 | 2 #FFD966 | 1 #FF0000
 *   • Colonne NOM_PRENOM en gras pour toutes les lignes élèves
 */

/************************** CONSTANTES **************************/
const STATS_ROW_COUNT     = 6;
const CLEANUP_BUFFER_ROWS = 3;
const CLEANUP_BUFFER_COLS = 1;

const HEADER_BACKGROUND_COLOR       = "#C6E0B4"; // vert clair uniforme
const HEADER_FONT_WEIGHT            = "bold";
const HEADER_HORIZONTAL_ALIGNMENT   = "center";
const HEADER_VERTICAL_ALIGNMENT     = "middle";

const EVEN_ROW_BACKGROUND           = "#f8f9fa";
const STATS_SEPARATOR_ROW_BACKGROUND = "#dcdcdc";

// Palette maquette
const SEXE_COLORS  = { M: "#4F81BD", F: "#F28EA8" };
const LV2_COLORS   = { ESP: "#E59838", ITA: "#73C6B6", ALL: "#F4B084", AUTRE: "#D9D9D9" };
const OPTION_FORMATS = [
  { text:"GREC",  bgColor:"#C0392B", fgColor:"#FFFFFF" },
  { text:"LATIN", bgColor:"#641E16", fgColor:"#FFFFFF" },
  { text:"LLCA",  bgColor:"#F4B084", fgColor:"#000000" },
  { text:"CHAV",  bgColor:"#6C3483", fgColor:"#FFFFFF" },
  { text:"UPE2A", bgColor:"#D5D8DC", fgColor:null         }
];
const SCORE_COLORS = { 1:"#FF0000", 2:"#FFD966", 3:"#3CB371", 4:"#006400" };
const STATS_AVERAGE_STYLE = { bgColor:"#34495e", fgColor:"#ffffff", fmt:"0.00", align:"center", bold:true };

const DEFAULT_REQUIRED_HEADERS_CONFIG = {
  ID_ELEVE   : ["ID_ELEVE","ID","IDENTIFIANT","IDELEVE"],
  NOM_PRENOM : ["NOM & PRENOM","NOM_PRENOM","NOM PRENOM","ELEVE","NOM ET PRENOM"],
  SEXE       : ["SEXE","GENRE"],
  LV2        : ["LV2","LANGUE","LANGUE VIVANTE 2"],
  OPT        : ["OPT","OPTION","OPTIONS"],
  COM        : ["COM","COMPORTEMENT","H"],
  TRA        : ["TRA","TRAVAIL","I"],
  PART       : ["PART","PARTICIPATION","J"],
  ABS        : ["ABS","ABSENCES","ASSIDUITE","K"],
  INDICATEUR : ["INDICATEUR","IND","L"],
  ASSO       : ["ASSO","ASSOCIATION","ASSOC"],
  DISSO      : ["DISSO","DISSOCIATION","DISSOC"],
  CLASSE_DEF : ["CLASSE DEF","CLASSE DEFINITIVE","CLASSE_DEF"],
  SOURCE     : ["SOURCE","ORIGINE"]
};
const CANONICAL_ID_ELEVE_HEADER = "ID_ELEVE";

/************************** UTILITAIRES **************************/
// Au début du fichier Phase1c_PARITE.gs (par exemple, ligne 1)
// const normalizeHeader = s => String(s||"")
//   .replace(/[\s\u00A0\u200B-\u200D]/g, "")
//   .toUpperCase();

function setColumnWidths(sheet,cfgs){
  cfgs.forEach(c=>{if(c.col>0&&c.col<=sheet.getMaxColumns())try{sheet.setColumnWidth(c.col,c.width);}catch(e){Logger.log(`Width col ${c.col}: ${e}`);}});
}

/**
 * Détermine la couleur de texte appropriée pour un score donné
 * @param {number} score Le score (1-4)
 * @return {string} Code couleur hexadécimal pour le texte
 */
function getScoreFontColor(score) {
  // Retourne une couleur de texte appropriée basée sur le score
  switch(score) {
    case 1: return "#FFFFFF"; // Blanc sur fond rouge
    case 2: return "#000000"; // Noir sur fond jaune
    case 3: return "#FFFFFF"; // Blanc sur fond vert clair
    case 4: return "#FFFFFF"; // Blanc sur fond vert foncé
    default: return "#000000"; // Noir par défaut
  }
}

/************************** MAIN **************************/
function miseEnFormeOngletsTest(options={}){
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const cfg  = {
    applyColumnFormatting : options.applyColumnFormatting !== false,
    applyConditionalRules : options.applyConditionalRules !== false,
    calculateAndWriteStats: options.calculateAndWriteStats !== false,
    applyFiltersAndZebra  : options.applyFiltersAndZebra   !== false,
    cleanUnused           : options.cleanUnused           !== false
  };

  const sheets = getTestSheetsNames();
  if(!sheets.length) return {success:true,message:"Aucun onglet TEST trouvé.",bilan:{errors:[]}};
  ss.toast("Mise en forme des onglets TEST…","En cours…",-1);

  const classeDefList = sheets.map(n=>n.replace(/TEST$/i,"")+"DEF").sort();
  const validationRuleDEF = classeDefList.length ? SpreadsheetApp.newDataValidation()
     .requireValueInList(classeDefList,true).setAllowInvalid(false)
     .setHelpText("Sélectionnez la classe DEF").build() : null;

  const bilan = {classes:[],genreCounts:[],criteresMoyennes:[],lv2Counts:[],optionsCounts:[],errors:[]};

  /* ————————————  BOUCLE SUR LES FEUILLES  ———————————— */
  sheets.forEach((name,idx)=>{
    ss.toast(`Formatage ${name} (${idx+1}/${sheets.length})`,`Progression`,-1);
    const sheet = ss.getSheetByName(name);
    if(!sheet){bilan.errors.push(`${name}: non trouvé`);return;}

    /* A. Forçage ID_ELEVE colonne A */
    const rawHdr = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
    const hdrNorm = rawHdr.map(normalizeHeader);
    if(hdrNorm[0]!==normalizeHeader(CANONICAL_ID_ELEVE_HEADER)){
      const idxFound = hdrNorm.indexOf(normalizeHeader(CANONICAL_ID_ELEVE_HEADER));
      if(idxFound>-1){ sheet.moveColumns(sheet.getRange(1,idxFound+1,sheet.getMaxRows()),1); }
      else { sheet.insertColumnBefore(1); sheet.getRange(1,1).setValue(CANONICAL_ID_ELEVE_HEADER); }
    }

    /* Style en‑tête uniforme */
    sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn()))
         .setFontWeight(HEADER_FONT_WEIGHT)
         .setBackground(HEADER_BACKGROUND_COLOR)
         .setHorizontalAlignment(HEADER_HORIZONTAL_ALIGNMENT)
         .setVerticalAlignment(HEADER_VERTICAL_ALIGNMENT);

    /* A2. Ajout en-tête CLASSE DEF si manquant */
    const CLASSE_DEF_COL = 18; // Colonne R est la 18ème colonne
    // S'assurer que la colonne R existe
    if (sheet.getMaxColumns() < CLASSE_DEF_COL) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), CLASSE_DEF_COL - sheet.getMaxColumns());
    }
    
    // Vérifier si l'en-tête CLASSE DEF est vide ou absent
    if (sheet.getLastColumn() < CLASSE_DEF_COL || 
        String(sheet.getRange(1, CLASSE_DEF_COL).getValue()).trim() === "") {
      // Ajouter l'en-tête
      sheet.getRange(1, CLASSE_DEF_COL)
        .setValue("CLASSE DEF")
        .setFontWeight(HEADER_FONT_WEIGHT)
        .setBackground(HEADER_BACKGROUND_COLOR)
        .setHorizontalAlignment(HEADER_HORIZONTAL_ALIGNMENT)
        .setVerticalAlignment(HEADER_VERTICAL_ALIGNMENT);
      Logger.log(`Ajout en-tête "CLASSE DEF" en colonne R pour ${name}`);
    }

    /* A3. cacher colonne */
    cacherColonnesABCDansOngletsTEST();

    /* B. Index des colonnes */
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(normalizeHeader);
    const idxFn = (aliases,def=-1)=>{
      const list=[].concat(aliases).map(normalizeHeader);
      const i=headers.findIndex(h=>list.includes(h));
      return i>=0?i+1:def;};
    const colMap={}; Object.keys(DEFAULT_REQUIRED_HEADERS_CONFIG).forEach(k=>{colMap[k]=idxFn(DEFAULT_REQUIRED_HEADERS_CONFIG[k]);});
    if(colMap.ID_ELEVE!==1){bilan.errors.push(`${name}: ID_ELEVE pas en A`);return;}

    /* B2. S'assurer que CLASSE_DEF est défini dans colMap */
    if(colMap.CLASSE_DEF <= 0) {
      // Si CLASSE_DEF n'est pas défini, forcer sa définition à la colonne R
      colMap.CLASSE_DEF = CLASSE_DEF_COL;
      Logger.log(`Forcé colMap.CLASSE_DEF = ${CLASSE_DEF_COL} (colonne R) pour ${name}`);
    }

    /* C. Dernière ligne */
    const lastRow=sheet.getLastRow(); let lastDataRow=1;
    if(lastRow>1){const ids=sheet.getRange(2,1,lastRow-1,1).getValues().flat();
      for(let i=ids.length-1;i>=0;i--){if(String(ids[i]).trim()){lastDataRow=i+2;break;}}}
    const nRows= lastDataRow>1 ? lastDataRow-1 : 0;

    /* 1) Column formatting */
    if(cfg.applyColumnFormatting){
      if(colMap.SOURCE>0) safeHideColumns(sheet,colMap.SOURCE);
      safeHideColumns(sheet,16,2); safeHideColumns(sheet,19,5);

      const widths=[];
      if(colMap.NOM_PRENOM>0) widths.push({col:colMap.NOM_PRENOM,width:200});
      [colMap.SEXE,colMap.LV2,colMap.OPT,colMap.INDICATEUR,colMap.ASSO,colMap.DISSO].forEach(c=>c>0&&widths.push({col:c,width:70}));
      [colMap.COM,colMap.TRA,colMap.PART,colMap.ABS].forEach(c=>c>0&&widths.push({col:c,width:60}));
      if(colMap.CLASSE_DEF>0) widths.push({col:colMap.CLASSE_DEF,width:100});
      setColumnWidths(sheet,widths);
      if(sheet.getFrozenRows()===0&&sheet.getLastRow()>0) sheet.setFrozenRows(1);

      // Col NOM_PRENOM en gras
      if(colMap.NOM_PRENOM>0&&nRows>0) sheet.getRange(2,colMap.NOM_PRENOM,nRows,1).setFontWeight("bold");
    }

    /* 2) Conditional rules */
    if(cfg.applyConditionalRules&&nRows>0) applyConditionalFormattingRules_Maquette(sheet,nRows,colMap);
    
    /* 3) Zebra + filtre */
    if(cfg.applyFiltersAndZebra){
      if(nRows>0){ for(let r=2;r<=lastDataRow;r+=2) sheet.getRange(r,1,1,sheet.getMaxColumns()).setBackground(EVEN_ROW_BACKGROUND); }
      if(sheet.getMaxRows()>lastDataRow){ sheet.insertRowAfter(lastDataRow);
        sheet.getRange(lastDataRow+1,1,1,sheet.getMaxColumns()).setBackground(STATS_SEPARATOR_ROW_BACKGROUND).clearContent(); }
      if(sheet.getFilter()) sheet.getFilter().remove();
      sheet.getRange(1,1,Math.max(1,lastDataRow),sheet.getMaxColumns()).createFilter();
    }

    /* 4) Validation DEF */
    if(validationRuleDEF && nRows>0) {
      // Toujours appliquer la validation à la colonne R (CLASSE_DEF), qu'elle soit détectée ou forcée
      sheet.getRange(2, CLASSE_DEF_COL, nRows, 1).setDataValidation(validationRuleDEF);
      Logger.log(`Liste déroulante CLASSE DEF appliquée en colonne R pour ${nRows} lignes dans ${name}`);
    }

    /* 5) Stats */
    if(cfg.calculateAndWriteStats){
      const stats=nRows>0?calculateSheetStats(sheet,nRows,colMap,1):null;
      const statsRow=(cfg.applyFiltersAndZebra&&nRows>0)?lastDataRow+2:lastDataRow+1;
      writeSheetStats(sheet,statsRow,colMap,stats);
      if(stats){
        bilan.classes.push(name);
        bilan.genreCounts.push(stats.genreCounts);
        bilan.criteresMoyennes.push(stats.criteresMoyennes);
        bilan.lv2Counts.push(stats.lv2Counts);
        bilan.optionsCounts.push(stats.optionsCounts);
      }
    }

    /* 6) Cleanup - Modifié pour ne pas supprimer la colonne R */
    if(cfg.cleanUnused){
      const finalRow=lastDataRow+(cfg.applyFiltersAndZebra&&nRows>0?1:0)+(cfg.calculateAndWriteStats?STATS_ROW_COUNT:0);
      // Assurons-nous que maxUsedCol est au moins 18 (colonne R)
      const maxUsedCol=Math.max(...Object.values(colMap).filter(c=>c>0), CLASSE_DEF_COL);
      cleanUnusedRowsAndColumns(sheet,finalRow,maxUsedCol,CLEANUP_BUFFER_ROWS,CLEANUP_BUFFER_COLS);
    }
  });

  ss.toast("","",-1);
  const msg=bilan.errors.length?`Terminé avec ${bilan.errors.length} erreur(s). Voir logs.`:"Mise en forme terminée avec succès.";
  ss.toast(msg,"Terminé",5);
  return {success:!bilan.errors.length,message:msg,bilan};
}

/*********************** FONCTIONS AIDE (identiques à v3.2 sauf palette) ***********************/

function applyConditionalFormattingRules_Maquette(sheet, nRows, colMap) {
  try {
    sheet.clearConditionalFormatRules();
    const rules = []; 
    const R = SpreadsheetApp.newConditionalFormatRule; 
    const start = 2;
    const add = (rng, bld) => { 
      try { 
        rules.push(bld.setRanges([rng]).build()); 
      } catch(e) { 
        Logger.log(`Warn: Échec création règle pour ${rng.getA1Notation()}: ${e}`); 
      }
    };
    
    // --- Format SEXE (Maquette) avec NOUVELLES COULEURS ---
    if(colMap.SEXE > 0 && nRows > 0) {
      const r = sheet.getRange(start, colMap.SEXE, nRows, 1);
      
      // Nouvelles couleurs plus douces
      const sexeColorM = "#B3CFEC"; // Bleu pâle superbe
      const sexeColorF = "#FFD1DC"; // Rose pâle superbe
      
      add(r, R().whenTextEqualTo('M').setBackground(sexeColorM));
      add(r, R().whenTextEqualTo('F').setBackground(sexeColorF));
      
      // Application directe du gras et centrage
      r.setFontWeight("bold").setHorizontalAlignment("center");
    }
    
    // --- Format LV2 (Maquette ESP et AUTRES) ---
    if(colMap.LV2 > 0 && nRows > 0) {
      const r = sheet.getRange(start, colMap.LV2, nRows, 1);
      
      // Garder l'orange pour ESP et ajouter bleu-vert clair pour autres
      const lv2ColorESP = "#E59838";   // Orange original pour ESP
      const lv2ColorAUTRE = "#A3E4D7"; // Bleu-vert clair magnifique
      
      add(r, R().whenTextEqualTo('ESP').setBackground(lv2ColorESP));
      
      // Ajouter règle pour AUTRES LV2 (tout ce qui n'est pas ESP et pas vide)
      add(r, R().whenFormulaSatisfied('AND(ISTEXT(F2),F2<>"ESP",F2<>"")').setBackground(lv2ColorAUTRE));
      
      // Application directe du gras et centrage
      r.setFontWeight("bold").setHorizontalAlignment("center");
    }
    
    // --- Format OPTIONS (Inchangé) ---
    if(colMap.OPT > 0 && nRows > 0) {
      const r = sheet.getRange(start, colMap.OPT, nRows, 1);
      for(const opt of OPTION_FORMATS) {
        const bld = R().whenTextEqualTo(opt.text).setBackground(opt.bgColor);
        if(opt.fgColor) bld.setFontColor(opt.fgColor);
        add(r, bld);
      }
      
      // Application directe du gras et centrage
      r.setFontWeight("bold").setHorizontalAlignment("center");
    }
    
    // --- Format SCORES (Maquette COM, TRA, PART, ABS) avec NOUVELLES COULEURS ---
    ['COM', 'TRA', 'PART', 'ABS'].forEach(cKey => {
      if(colMap[cKey] > 0 && nRows > 0) {
        const range = sheet.getRange(start, colMap[cKey], nRows, 1);
        
        // Nouvelles couleurs plus douces pour les scores avec jaune plus foncé
        const scoreColors = {
          "1": "#FF0000",     // Rouge (inchangé)
          "2": "#FFD966",     // Jaune moins pâle
          "3": "#A8E4BC",     // Vert plus pâle, plus léger
          "4": "#006400"      // Vert foncé (inchangé)
        };
        
        // Appliquer les règles de couleur pour chaque score
        for(let score = 1; score <= 4; score++) {
          const scoreStr = score.toString();
          if(scoreStr in scoreColors) {
            // Couleur de texte différente selon le score
            // Score 2 et 3 en noir, score 1 et 4 en blanc
            const fontColor = (score === 2 || score === 3) ? "#000000" : "#FFFFFF";
            
            add(range, R().whenNumberEqualTo(score)
                .setBackground(scoreColors[scoreStr])
                .setFontColor(fontColor));
          }
        }
        
        // Application directe du gras et centrage
        range.setFontWeight("bold").setHorizontalAlignment("center");
      }
    });
    
    // --- Mettre en gras et centrer les colonnes L, M, N ---
    ['INDICATEUR', 'ASSO', 'DISSO'].forEach(cKey => {
      if(colMap[cKey] > 0 && nRows > 0) {
        const range = sheet.getRange(start, colMap[cKey], nRows, 1);
        range.setFontWeight("bold").setHorizontalAlignment("center");
      }
    });
    
    // Appliquer les règles
    if(rules.length) sheet.setConditionalFormatRules(rules);
    
    return true;
  } catch(e) {
    Logger.log(`Erreur Format Cond. (Maquette) ${sheet.getName()}: ${e}`);
    return false;
  }
}

// ================================================
// --- MISE EN PAGE ---
// ================================================

function writeSheetStats(sheet, row, colMap, stats) {
  // row est déjà +2 après le dernier élève selon votre code principal
  
  const maxCol = Math.max(
    colMap.NOM_PRENOM || 1,
    colMap.SEXE || 1,
    colMap.LV2 || 1,
    colMap.OPT || 1,
    colMap.COM || 1,
    colMap.TRA || 1,
    colMap.PART || 1,
    colMap.ABS || 1
  );
  
  if (sheet.getMaxRows() < row + 7) return; // +7 pour couvrir toutes les lignes de statistiques
  
  // Nettoyer la zone des statistiques
  sheet.getRange(row, 1, 7, maxCol).clearContent().clearFormat();
  
  // Fonction d'aide pour définir la valeur et le formatage des cellules
  const set = (r, c, v, s = {}) => {
    if (c > 0 && c <= sheet.getMaxColumns() && r <= sheet.getMaxRows()) {
      const cell = sheet.getRange(r, c);
      cell.setValue(v);
      if (s.bold) cell.setFontWeight('bold');
      if (s.align) cell.setHorizontalAlignment(s.align);
      if (s.bg) cell.setBackground(s.bg);
      if (s.fg) cell.setFontColor(s.fg);
      if (s.fmt) cell.setNumberFormat(s.fmt);
      if (s.italic) cell.setFontStyle('italic');
    }
  };
  
  // Gérer le cas où il n'y a pas de données
  if (!stats) {
    set(row, 1, 'Pas de données', { italic: true });
    return;
  }
  
  // Statistiques de genre/langues sur le côté
  // 1. Nombre de filles (F) dans la colonne E (rose)
  set(row, 5, stats.genreCounts[0], { align: 'center', bg: SEXE_COLORS.F, bold: true });
  
  // 2. Nombre de garçons (M) dans la colonne E (bleu, police blanche)
  set(row + 1, 5, stats.genreCounts[1], { align: 'center', bg: SEXE_COLORS.M, fg: 'white', bold: true });
  
  // 3. Nombre d'élèves ESP dans la colonne F
  set(row, 6, stats.lv2Counts[0], { align: 'center', bg: LV2_COLORS.ESP, bold: true });
  
  // 4. Nombre d'élèves non-ESP dans la colonne F
  set(row + 1, 6, stats.lv2Counts[1], { align: 'center', bg: LV2_COLORS.AUTRE, bold: true });
  
  // 5. Nombre d'élèves avec options dans la colonne G
  set(row, 7, stats.optionsCounts[0], { align: 'center', bold: true });
  
  // SCORES: exactement comme dans l'image

  // En bas: scores et moyennes dans les colonnes H, I, J, K
  // On commence par placer les scores dans l'ordre 4->1 comme sur l'image
  // Pas de libellés de scores, juste les valeurs

  ['COM', 'TRA', 'PART', 'ABS'].forEach((k, i) => {
    const columnIndex = colMap[k];
    if (columnIndex > 0) {
      // Score 4 (vert foncé) - utiliser une police blanche pour le contraste
      set(row, columnIndex, stats.criteresScores[k][4], { align: 'center', bg: SCORE_COLORS[4], bold: true, fg: 'white' });
      
      // Score 3 (vert clair)
      set(row + 1, columnIndex, stats.criteresScores[k][3], { align: 'center', bg: SCORE_COLORS[3], bold: true });
      
      // Score 2 (jaune)
      set(row + 2, columnIndex, stats.criteresScores[k][2], { align: 'center', bg: SCORE_COLORS[2], bold: true });
      
      // Score 1 (rouge)
      set(row + 3, columnIndex, stats.criteresScores[k][1], { align: 'center', bg: SCORE_COLORS[1], bold: true, fg: 'white' });
      
      // Moyenne (pas de couleur de fond)
      set(row + 4, columnIndex, stats.criteresMoyennes[i], { align: 'center', bold: true, fmt: '#,##0.00' });
    }
  });
  
  // Optionnellement: ajouter des libellés dans la colonne D si nécessaire
  // set(row + 4, 4, 'Moyenne :', { bold: true, align: 'right' });
}

// --- FONCTIONS UTILITAIRES POUR LE FORMATAGE ET LES STATS ---

function setColumnWidths(sheet, widthsConfig) {
    widthsConfig.forEach(conf => {
        if (conf.col > 0 && conf.col <= sheet.getMaxColumns()) {
            try { sheet.setColumnWidth(conf.col, conf.width); }
            catch (e) { Logger.log(`Erreur largeur col ${conf.col} pour ${sheet.getName()}: ${e}`); }
        }
    });
}

function applyConditionalFormattingRules(sheet, numDataRows, colMap) {
    try {
        sheet.clearConditionalFormatRules();
        const rules = [];
        const dataRangeStartRow = 2;

        const addRuleToRange = (range, builderFunction) => {
            try {
                const ruleBuilder = SpreadsheetApp.newConditionalFormatRule();
                builderFunction(ruleBuilder);
                rules.push(ruleBuilder.setRanges([range]).build());
            } catch (e) { Logger.log(`Erreur création règle pour ${range.getA1Notation()} sur ${sheet.getName()}: ${e}`);}
        };
        
        // Sexe
        if (colMap.SEXE > 0) {
            const rangeSexe = sheet.getRange(dataRangeStartRow, colMap.SEXE, numDataRows, 1);
            addRuleToRange(rangeSexe, b => b.whenTextEqualTo('M').setBackground(SEXE_COLORS.M));
            addRuleToRange(rangeSexe, b => b.whenTextEqualTo('F').setBackground(SEXE_COLORS.F));
        }
        // LV2
        if (colMap.LV2 > 0) {
            const rangeLV2 = sheet.getRange(dataRangeStartRow, colMap.LV2, numDataRows, 1);
            for (const lang in LV2_COLORS) {
                 addRuleToRange(rangeLV2, b => b.whenTextEqualTo(lang).setBackground(LV2_COLORS[lang]));
            }
        }
        // Options
        if (colMap.OPT > 0) {
            const rangeOpt = sheet.getRange(dataRangeStartRow, colMap.OPT, numDataRows, 1);
            OPTION_FORMATS.forEach(optFmt => {
                addRuleToRange(rangeOpt, b => {
                    b.whenTextContains(optFmt.text).setBackground(optFmt.bgColor);
                    if (optFmt.fgColor) b.setFontColor(optFmt.fgColor);
                });
            });
        }
        // Critères (COM, TRA, PART, ABS)
        const criteresCols = [colMap.COM, colMap.TRA, colMap.PART, colMap.ABS];
        criteresCols.forEach(critColIdx => {
            if (critColIdx > 0) {
                const rangeCrit = sheet.getRange(dataRangeStartRow, critColIdx, numDataRows, 1);
                for (const score in SCORE_COLORS) {
                    addRuleToRange(rangeCrit, b => b.whenTextEqualTo(score).setBackground(SCORE_COLORS[score]));
                }
            }
        });
        if (rules.length > 0) sheet.setConditionalFormatRules(rules);
    } catch (e) { Logger.log(`Erreur globale format conditionnel pour ${sheet.getName()}: ${e.message} \n ${e.stack}`); }
}

function calculateSheetStats(sheet, numDataRows, colMap, iIdCol) {
    const data = sheet.getRange(2, 1, numDataRows, sheet.getMaxColumns()).getValues();
    const validRows = data.filter(r => String(r[iIdCol - 1]).trim() !== "");
    if (validRows.length === 0) return null;

    const getColData = (colKey) => colMap[colKey] > 0 ? validRows.map(r => r[colMap[colKey] - 1]) : [];
    
    const countValues = (colData, value) => colData.filter(cell => String(cell).trim().toUpperCase() === String(value).toUpperCase()).length;
    const countNonEmpty = (colData) => colData.filter(cell => String(cell).trim() !== "").length;
    
    const toNumericArray = (colData) => colData.map(cell => {
        const num = Number(String(cell).replace(',', '.'));
        return isNaN(num) ? 0 : num;
    });
    const calculateAverage = (numArray) => {
        const filtered = numArray.filter(n => n > 0);
        return filtered.length > 0 ? filtered.reduce((s, v) => s + v, 0) / filtered.length : 0;
    };

    const sexeData = getColData('SEXE');
    const lv2Data  = getColData('LV2');
    const optData  = getColData('OPT');

    const stats = {
        genreCounts: [countValues(sexeData, 'F'), countValues(sexeData, 'M')],
        lv2Counts: [
            countValues(lv2Data, 'ESP'), 
            lv2Data.length - countValues(lv2Data, 'ESP') - countValues(lv2Data, '') // Autres (non-ESP et non-vide)
        ],
        optionsCounts: [countNonEmpty(optData)],
        criteresScores: {},
        criteresMoyennes: []
    };

    ['COM', 'TRA', 'PART', 'ABS'].forEach(critKey => {
        const critDataNum = toNumericArray(getColData(critKey));
        stats.criteresScores[critKey] = {};
        for (let score = 1; score <= 4; score++) {
            stats.criteresScores[critKey][score] = critDataNum.filter(val => val === score).length;
        }
        stats.criteresMoyennes.push(calculateAverage(critDataNum));
    });
     while(stats.criteresMoyennes.length < 4) stats.criteresMoyennes.push(0); // Assurer 4 moyennes

    return stats;
}

// --- FONCTIONS UTILITAIRES GÉNÉRIQUES ---
function safeHideColumns(sheet, startColumn, numColumns = 1) {
    if (startColumn > 0 && startColumn <= sheet.getMaxColumns()) {
        const actualNumColumns = Math.min(numColumns, sheet.getMaxColumns() - startColumn + 1);
        if (actualNumColumns > 0) {
            try { sheet.hideColumns(startColumn, actualNumColumns); }
            catch (e) { Logger.log(`Erreur masquage cols ${startColumn}-${startColumn+actualNumColumns-1} pour ${sheet.getName()}: ${e}`); }
        }
    }
}

function cleanUnusedRowsAndColumns(sheet, lastUsedRow, lastUsedCol, bufferRows = 0, bufferCols = 0) {
  try {
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();

    const rowToDeleteStart = lastUsedRow + 1 + bufferRows;
    if (rowToDeleteStart < maxRows && rowToDeleteStart > 0) { // S'assurer que rowToDeleteStart est valide
      sheet.deleteRows(rowToDeleteStart, maxRows - rowToDeleteStart + 1);
    }

    const colToDeleteStart = lastUsedCol + 1 + bufferCols;
    if (colToDeleteStart < maxCols && colToDeleteStart > 0) { // S'assurer que colToDeleteStart est valide
      sheet.deleteColumns(colToDeleteStart, maxCols - colToDeleteStart + 1);
    }
  } catch (e) {
    Logger.log(`Erreur nettoyage ${sheet.getName()}: ${e.message} \n ${e.stack}`);
  }
}

// --- FONCTION PRINCIPALE D'ENTRÉE ---
function entryPointMiseEnFormeTest() {
    const options = {
        applyColumnFormatting: true,
        applyConditionalRules: true,
        calculateAndWriteStats: true,
        applyFiltersAndZebra: true,
        cleanUnused: true,
    };
    const result = miseEnFormeOngletsTest(options);
    SpreadsheetApp.getUi().alert(`Mise en forme "Maquette" : ${result.message}`);
}

/**
 * Récupère la liste des onglets dont le nom se termine par "TEST" (insensible à la casse).
 * @return {string[]} Tableau des noms d'onglets à formater.
 */
function getTestSheetsNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testPattern = /TEST$/i;
  return ss.getSheets()
           .map(sheet => sheet.getName())
           .filter(name => testPattern.test(name));
}

// --- FONCTIONS UTILITAIRES POUR LE FORMATAGE ET LES STATS ---

function setColumnWidths(sheet, widthsConfig) {
    widthsConfig.forEach(conf => {
        if (conf.col > 0 && conf.col <= sheet.getMaxColumns()) {
            try { sheet.setColumnWidth(conf.col, conf.width); }
            catch (e) { Logger.log(`Erreur largeur col ${conf.col} pour ${sheet.getName()}: ${e}`); }
        }
    });
}

function applyConditionalFormattingRules(sheet, numDataRows, colMap) {
    try {
        sheet.clearConditionalFormatRules();
        const rules = [];
        const dataRangeStartRow = 2;

        const addRuleToRange = (range, builderFunction) => {
            try {
                const ruleBuilder = SpreadsheetApp.newConditionalFormatRule();
                builderFunction(ruleBuilder);
                rules.push(ruleBuilder.setRanges([range]).build());
            } catch (e) { Logger.log(`Erreur création règle pour ${range.getA1Notation()} sur ${sheet.getName()}: ${e}`);}
        };
        
        // Sexe
        if (colMap.SEXE > 0) {
            const rangeSexe = sheet.getRange(dataRangeStartRow, colMap.SEXE, numDataRows, 1);
            addRuleToRange(rangeSexe, b => b.whenTextEqualTo('M').setBackground(SEXE_COLORS.M));
            addRuleToRange(rangeSexe, b => b.whenTextEqualTo('F').setBackground(SEXE_COLORS.F));
        }
        // LV2
        if (colMap.LV2 > 0) {
            const rangeLV2 = sheet.getRange(dataRangeStartRow, colMap.LV2, numDataRows, 1);
            for (const lang in LV2_COLORS) {
                 addRuleToRange(rangeLV2, b => b.whenTextEqualTo(lang).setBackground(LV2_COLORS[lang]));
            }
        }
        // Options
        if (colMap.OPT > 0) {
            const rangeOpt = sheet.getRange(dataRangeStartRow, colMap.OPT, numDataRows, 1);
            OPTION_FORMATS.forEach(optFmt => {
                addRuleToRange(rangeOpt, b => {
                    b.whenTextContains(optFmt.text).setBackground(optFmt.bgColor);
                    if (optFmt.fgColor) b.setFontColor(optFmt.fgColor);
                });
            });
        }
        // Critères (COM, TRA, PART, ABS)
        const criteresCols = [colMap.COM, colMap.TRA, colMap.PART, colMap.ABS];
        criteresCols.forEach(critColIdx => {
            if (critColIdx > 0) {
                const rangeCrit = sheet.getRange(dataRangeStartRow, critColIdx, numDataRows, 1);
                for (const score in SCORE_COLORS) {
                    addRuleToRange(rangeCrit, b => b.whenTextEqualTo(score).setBackground(SCORE_COLORS[score]));
                }
            }
        });
        if (rules.length > 0) sheet.setConditionalFormatRules(rules);
    } catch (e) { Logger.log(`Erreur globale format conditionnel pour ${sheet.getName()}: ${e.message} \n ${e.stack}`); }
}

function calculateSheetStats(sheet, numDataRows, colMap, iIdCol) {
    const data = sheet.getRange(2, 1, numDataRows, sheet.getMaxColumns()).getValues();
    const validRows = data.filter(r => String(r[iIdCol - 1]).trim() !== "");
    if (validRows.length === 0) return null;

    const getColData = (colKey) => colMap[colKey] > 0 ? validRows.map(r => r[colMap[colKey] - 1]) : [];
    
    const countValues = (colData, value) => colData.filter(cell => String(cell).trim().toUpperCase() === String(value).toUpperCase()).length;
    const countNonEmpty = (colData) => colData.filter(cell => String(cell).trim() !== "").length;
    
    const toNumericArray = (colData) => colData.map(cell => {
        const num = Number(String(cell).replace(',', '.'));
        return isNaN(num) ? 0 : num;
    });
    const calculateAverage = (numArray) => {
        const filtered = numArray.filter(n => n > 0);
        return filtered.length > 0 ? filtered.reduce((s, v) => s + v, 0) / filtered.length : 0;
    };

    const sexeData = getColData('SEXE');
    const lv2Data  = getColData('LV2');
    const optData  = getColData('OPT');

    const stats = {
        genreCounts: [countValues(sexeData, 'F'), countValues(sexeData, 'M')],
        lv2Counts: [
            countValues(lv2Data, 'ESP'), 
            lv2Data.length - countValues(lv2Data, 'ESP') - countValues(lv2Data, '') // Autres (non-ESP et non-vide)
        ],
        optionsCounts: [countNonEmpty(optData)],
        criteresScores: {},
        criteresMoyennes: []
    };

    ['COM', 'TRA', 'PART', 'ABS'].forEach(critKey => {
        const critDataNum = toNumericArray(getColData(critKey));
        stats.criteresScores[critKey] = {};
        for (let score = 1; score <= 4; score++) {
            stats.criteresScores[critKey][score] = critDataNum.filter(val => val === score).length;
        }
        stats.criteresMoyennes.push(calculateAverage(critDataNum));
    });
     while(stats.criteresMoyennes.length < 4) stats.criteresMoyennes.push(0); // Ensure 4 averages

    return stats;
}


// --- FONCTIONS UTILITAIRES GÉNÉRIQUES ---
function safeHideColumns(sheet, startColumn, numColumns = 1) {
    if (startColumn > 0 && startColumn <= sheet.getMaxColumns()) {
        const actualNumColumns = Math.min(numColumns, sheet.getMaxColumns() - startColumn + 1);
        if (actualNumColumns > 0) {
            try { sheet.hideColumns(startColumn, actualNumColumns); }
            catch (e) { Logger.log(`Erreur masquage cols ${startColumn}-${startColumn+actualNumColumns-1} pour ${sheet.getName()}: ${e}`); }
        }
    }
}

function cleanUnusedRowsAndColumns(sheet, lastUsedRow, lastUsedCol, bufferRows = 0, bufferCols = 0) {
  try {
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();

    const rowToDeleteStart = lastUsedRow + 1 + bufferRows;
    if (rowToDeleteStart < maxRows && rowToDeleteStart > 0) { // S'assurer que rowToDeleteStart est valide
      sheet.deleteRows(rowToDeleteStart, maxRows - rowToDeleteStart + 1);
    }

    const colToDeleteStart = lastUsedCol + 1 + bufferCols;
    if (colToDeleteStart < maxCols && colToDeleteStart > 0) { // S'assurer que colToDeleteStart est valide
      sheet.deleteColumns(colToDeleteStart, maxCols - colToDeleteStart + 1);
    }
  } catch (e) {
    Logger.log(`Erreur nettoyage ${sheet.getName()}: ${e.message} \n ${e.stack}`);
  }
}

// --- FONCTION PRINCIPALE D'ENTRÉE ---
function entryPointMiseEnFormeTest() {
    const options = {
        applyColumnFormatting: true,
        applyConditionalRules: true,
        calculateAndWriteStats: true,
        applyFiltersAndZebra: true,
        cleanUnused: true,
    };
    const result = miseEnFormeOngletsTest(options);
    SpreadsheetApp.getUi().alert(`Mise en forme "Maquette" : ${result.message}`);
}

/**
 * Récupère la liste des onglets dont le nom se termine par "TEST" (insensible à la casse).
 * @return {string[]} Tableau des noms d’onglets à formater.
 */
function getTestSheetsNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testPattern = /TEST$/i;
  return ss.getSheets()
           .map(sheet => sheet.getName())
           .filter(name => testPattern.test(name));
}


// ================================================
// FONCTION BILAN (Supposée Exister - VOTRE CODE ORIGINAL) - (Inchangée)
// ================================================
// ... (Votre fonction createBilanTest originale complète ici) ...
function createBilanTest(ss, bilanData) {
    // ... METTRE VOTRE CODE ORIGINAL COMPLET de createBilanTest ici ...
    // Ce code utilisera l'objet bilanData préparé par miseEnFormeOngletsTest
     Logger.log("Création/Mise à jour de l'onglet BILAN TEST...");
     try {
        let bilanSheet = ss.getSheetByName("BILAN TEST");
        if (bilanSheet) {
            bilanSheet.clear({contentsOnly: true, formatOnly: false}); // Clear content and formats, keep sheet
            bilanSheet.getCharts().forEach(chart => bilanSheet.removeChart(chart));
        } else {
            bilanSheet = ss.insertSheet("BILAN TEST");
        }
        bilanSheet.setTabColor("#800080"); // Purple tab color

        // Titre
        bilanSheet.getRange("A1:J1").merge().setValue("BILAN COMPARATIF DES CLASSES")
          .setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center")
          .setBackground("#34495e").setFontColor("#ffffff").setVerticalAlignment("middle");
        bilanSheet.setRowHeight(1, 40);

        let currentRow = 3; // Start content below title

        // Tableau Effectifs et Parité
        if (bilanData.classes.length > 0) {
            bilanSheet.getRange(currentRow, 1).setValue("EFFECTIFS ET PARITÉ").setFontSize(14).setFontWeight("bold").setBackground("#eceff1"); // Light grey background for section title
             bilanSheet.getRange(currentRow, 1, 1, 6).merge(); // Merge section title cells
            currentRow++;
            const headersParite = ["CLASSE", "TOTAL", "FILLES", "GARÇONS", "% FILLES", "% GARÇONS"];
            bilanSheet.getRange(currentRow, 1, 1, headersParite.length).setValues([headersParite])
                .setFontWeight("bold").setBackground("#b0bec5").setHorizontalAlignment("center"); // Blue-grey header
                 bilanSheet.getRange(currentRow, 1, 1, headersParite.length).setBorder(true, true, true, true, false, false, "#607d8b", SpreadsheetApp.BorderStyle.SOLID_MEDIUM); // Header border
            currentRow++;
            const startPariteRow = currentRow;
            for (let i = 0; i < bilanData.classes.length; i++) {
                const row = currentRow + i;
                const classeCell = bilanSheet.getRange(row, 1);
                classeCell.setValue(bilanData.classes[i]); // Class name

                bilanSheet.getRange(row, 2).setFormula(`=SUM(C${row}:D${row})`).setNumberFormat("0"); // Total
                const filleCell = bilanSheet.getRange(row, 3);
                filleCell.setFormula(bilanData.genreCounts[i][0]).setNumberFormat("0"); // Filles Count
                if(filleCell.getFormula() !== "=0") filleCell.setBackground("#fce4ec"); // Light Pink

                const garconCell = bilanSheet.getRange(row, 4);
                garconCell.setFormula(bilanData.genreCounts[i][1]).setNumberFormat("0"); // Garçons Count
                 if(garconCell.getFormula() !== "=0") garconCell.setBackground("#e3f2fd"); // Light Blue

                const fillePercentCell = bilanSheet.getRange(row, 5);
                fillePercentCell.setFormula(`=IFERROR(C${row}/B${row};0)`).setNumberFormat("0.0%"); // % Filles
                 if(filleCell.getFormula() !== "=0") fillePercentCell.setBackground("#fce4ec").setFontColor("#880e4f"); // Dark Pink Font

                const garconPercentCell = bilanSheet.getRange(row, 6);
                garconPercentCell.setFormula(`=IFERROR(D${row}/B${row};0)`).setNumberFormat("0.0%"); // % Garçons
                 if(garconCell.getFormula() !== "=0") garconPercentCell.setBackground("#e3f2fd").setFontColor("#0d47a1"); // Dark Blue Font
            }
            currentRow += bilanData.classes.length;
            const endPariteRow = currentRow - 1;

            // Total Row
            bilanSheet.getRange(currentRow, 1).setValue("TOTAL").setFontWeight("bold");
            bilanSheet.getRange(currentRow, 2).setFormula(`=SUM(B${startPariteRow}:B${endPariteRow})`).setNumberFormat("0");
            bilanSheet.getRange(currentRow, 3).setFormula(`=SUM(C${startPariteRow}:C${endPariteRow})`).setNumberFormat("0");
            bilanSheet.getRange(currentRow, 4).setFormula(`=SUM(D${startPariteRow}:D${endPariteRow})`).setNumberFormat("0");
            bilanSheet.getRange(currentRow, 5).setFormula(`=IFERROR(C${currentRow}/B${currentRow};0)`).setNumberFormat("0.0%");
            bilanSheet.getRange(currentRow, 6).setFormula(`=IFERROR(D${currentRow}/B${currentRow};0)`).setNumberFormat("0.0%");
            bilanSheet.getRange(currentRow, 1, 1, 6).setBackground("#78909c").setFontWeight("bold").setFontColor("#ffffff"); // Darker Blue-grey Total row

            // Apply borders to the data table
            bilanSheet.getRange(startPariteRow, 1, bilanData.classes.length + 1, 6) // +1 for Total row
               .setBorder(true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID); // All borders

            const totalPariteRow = currentRow;

             // Chart: Effectifs par classe et par genre (Stacked Column)
             try {
                 // Select Classe, Filles, Garçons columns including headers and data rows (NOT total row)
                 const chartRange = bilanSheet.getRange(startPariteRow - 1, 1, bilanData.classes.length + 1, 4); // A:D range, Header + Data
                 const chartEffectifs = bilanSheet.newChart().setChartType(Charts.ChartType.COLUMN)
                    .addRange(chartRange)
                    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS) // Use Col A as category label
                    .setTransposeRowsAndColumns(false)
                    .setNumHeaders(1) // First row is header
                    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
                    .setOption('title', 'Effectifs par classe et par genre')
                    .setOption('titleTextStyle', {color: '#424242', fontSize: 14, bold: true})
                    .setOption('legend', { position: 'top', textStyle: { color: '#616161' } })
                    .setOption('colors', ['#fce4ec', '#e3f2fd']) // Filles (Pink), Garçons (Blue)
                    .setOption('isStacked', true) // Stack Filles/Garçons
                    .setOption('width', 600)
                    .setOption('height', 400)
                    .setOption('hAxis', { title: 'Classes', textStyle: { color: '#616161' } })
                    .setOption('vAxis', { title: 'Nombre d\'élèves', textStyle: { color: '#616161' }, minValue: 0, gridlines: { count: -1 }})
                    .setPosition(3, 8, 10, 10) // Position: row 3, col 8 (H), offset 10,10
                    .build();
                 bilanSheet.insertChart(chartEffectifs);
             } catch (e) { Logger.log(`Erreur création graphique effectifs: ${e.message}`); }

             currentRow = totalPariteRow + 3; // Add more space after the table/chart
        } else {
             bilanSheet.getRange(currentRow, 1).setValue("Aucune donnée de classe à afficher.").setFontStyle("italic");
             currentRow+=2;
        }

        // Tableau Moyennes des Critères
        if (bilanData.classes.length > 0 && bilanData.criteresMoyennes[0].length > 0) { // Check if criteres data exists
            bilanSheet.getRange(currentRow, 1).setValue("MOYENNES DES CRITÈRES").setFontSize(14).setFontWeight("bold").setBackground("#eceff1");
             bilanSheet.getRange(currentRow, 1, 1, 5).merge();
            currentRow++;
            const headersCriteres = ["CLASSE", "COM", "TRA", "PART", "ABS"]; // Assuming H=COM, I=TRA, J=PART, K=ABS
            bilanSheet.getRange(currentRow, 1, 1, headersCriteres.length).setValues([headersCriteres])
                .setFontWeight("bold").setBackground("#b0bec5").setHorizontalAlignment("center");
                 bilanSheet.getRange(currentRow, 1, 1, headersCriteres.length).setBorder(true, true, true, true, false, false, "#607d8b", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
            currentRow++;
            const startCritereRow = currentRow;
            for (let i = 0; i < bilanData.classes.length; i++) {
                const row = currentRow + i;
                bilanSheet.getRange(row, 1).setValue(bilanData.classes[i]);
                // Fill in the 4 criteria averages
                for (let j = 0; j < 4; j++) {
                     // Check if formula exists for this criterion, otherwise put empty
                     const formula = bilanData.criteresMoyennes[i][j] || "";
                     bilanSheet.getRange(row, 2 + j).setFormula(formula)
                         .setNumberFormat("0.00")
                         .setHorizontalAlignment("center");
                }
            }
            currentRow += bilanData.classes.length;
            const endCritereRow = currentRow - 1;

            // Apply borders to the data table
            bilanSheet.getRange(startCritereRow, 1, bilanData.classes.length, 5)
               .setBorder(true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);

            // Conditional Formatting for the averages (B to E)
            try {
                const criteresRange = bilanSheet.getRange(startCritereRow, 2, bilanData.classes.length, 4); // B:E range for data rows
                const rules = bilanSheet.getConditionalFormatRules(); // Get existing rules first
                rules.push(SpreadsheetApp.newConditionalFormatRule()
                    .setGradientMinpointWithValue("#e57373", SpreadsheetApp.InterpolationType.NUMBER, "1") // Red for 1
                    .setGradientMidpointWithValue("#fff59d", SpreadsheetApp.InterpolationType.NUMBER, "2.5") // Yellow for 2.5
                    .setGradientMaxpointWithValue("#81c784", SpreadsheetApp.InterpolationType.NUMBER, "4") // Green for 4
                    .setRanges([criteresRange]).build());
                bilanSheet.setConditionalFormatRules(rules); // Set combined rules
            } catch (e) { Logger.log(`Erreur formatage conditionnel moyennes: ${e.message}`); }

             // Chart: Radar Chart for Criteria Averages
             try {
                 const radarRange = bilanSheet.getRange(startCritereRow - 1, 1, bilanData.classes.length + 1, 5); // A:E range, Header + Data
                 const radarChart = bilanSheet.newChart().setChartType(Charts.ChartType.RADAR)
                    .addRange(radarRange)
                    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS) // Use Col A as series label
                    .setTransposeRowsAndColumns(false) // Check if transpose needed based on data structure
                    .setNumHeaders(1)
                    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
                    .setOption('title', 'Profil moyen des classes par critère')
                    .setOption('titleTextStyle', { color: '#424242', fontSize: 14, bold: true })
                    .setOption('legend', { position: 'right', textStyle: { color: '#616161' } })
                     .setOption('vAxis', {minValue: 1, maxValue: 4}) // Set scale 1 to 4 if appropriate
                    .setOption('width', 600)
                    .setOption('height', 400)
                    // Position below the Parity chart
                    .setPosition(startCritereRow + bilanData.classes.length + 1, 8, 10, 10)
                    .build();
                 bilanSheet.insertChart(radarChart);
             } catch (e) { Logger.log(`Erreur création graphique radar: ${e.message}`); }

             currentRow = endCritereRow + 3; // Add more space
        } else {
             Logger.log("Pas de données de critères à afficher dans le bilan.");
             currentRow += 1;
        }


         // Tableau Répartition LV2 / Options (Adapté au Patch)
        if (bilanData.classes.length > 0) {
            bilanSheet.getRange(currentRow, 1).setValue("RÉPARTITION LV2 / OPTIONS").setFontSize(14).setFontWeight("bold").setBackground("#eceff1");
            bilanSheet.getRange(currentRow, 1, 1, 4).merge(); // Merge for section title
            currentRow++;
            // PATCH: Headers adaptés: Classe, ESP, Autres LV2, Options Total
            const headersOptions = ["CLASSE", "ESP", "AUTRES LV2", "OPTIONS"];
            bilanSheet.getRange(currentRow, 1, 1, headersOptions.length).setValues([headersOptions])
                .setFontWeight("bold").setBackground("#b0bec5").setHorizontalAlignment("center");
            bilanSheet.getRange(currentRow, 1, 1, headersOptions.length).setBorder(true, true, true, true, false, false, "#607d8b", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
            currentRow++;
            const startOptionRow = currentRow;
            for (let i = 0; i < bilanData.classes.length; i++) {
                const row = currentRow + i;
                bilanSheet.getRange(row, 1).setValue(bilanData.classes[i]); // Classe

                // ESP Count (from bilanData.lv2Counts[i][0])
                const espCell = bilanSheet.getRange(row, 2);
                espCell.setFormula(bilanData.lv2Counts[i][0]).setNumberFormat("0");
                 if(espCell.getFormula() !== "=0") espCell.setBackground("#ffcc80"); // Light Orange

                // Autres LV2 Count (from bilanData.lv2Counts[i][1])
                const autreLv2Cell = bilanSheet.getRange(row, 3);
                autreLv2Cell.setFormula(bilanData.lv2Counts[i][1]).setNumberFormat("0");
                 if(autreLv2Cell.getFormula() !== "=0") autreLv2Cell.setBackground("#80cbc4"); // Light Teal

                // Options Total Count (from bilanData.optionsCounts[i][0]) - PATCH
                const optionCell = bilanSheet.getRange(row, 4);
                optionCell.setFormula(bilanData.optionsCounts[i][0]).setNumberFormat("0");
                 if(optionCell.getFormula() !== "=0") optionCell.setBackground("#bcaaa4"); // Light Brown
            }
            currentRow += bilanData.classes.length;
            const endOptionRow = currentRow - 1;

            // Total Row for LV2/Options
            bilanSheet.getRange(currentRow, 1).setValue("TOTAL").setFontWeight("bold");
            bilanSheet.getRange(currentRow, 2).setFormula(`=SUM(B${startOptionRow}:B${endOptionRow})`).setNumberFormat("0"); // Total ESP
            bilanSheet.getRange(currentRow, 3).setFormula(`=SUM(C${startOptionRow}:C${endOptionRow})`).setNumberFormat("0"); // Total Autres LV2
            bilanSheet.getRange(currentRow, 4).setFormula(`=SUM(D${startOptionRow}:D${endOptionRow})`).setNumberFormat("0"); // Total Options
            bilanSheet.getRange(currentRow, 1, 1, 4).setBackground("#78909c").setFontWeight("bold").setFontColor("#ffffff"); // Darker Total row

             // Apply borders to the data table
            bilanSheet.getRange(startOptionRow, 1, bilanData.classes.length + 1, 4) // +1 for Total row
               .setBorder(true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);

            // Chart: Répartition LV2 / Options par classe (PATCHED - Simple Column Chart)
            try {
                const chartOptionsRange = bilanSheet.getRange(startOptionRow - 1, 1, bilanData.classes.length + 1, 4); // A:D range, Header + Data
                const chartOptions = bilanSheet.newChart().setChartType(Charts.ChartType.COLUMN)
                    .addRange(chartOptionsRange)
                    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
                    .setTransposeRowsAndColumns(false)
                    .setNumHeaders(1)
                    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
                    .setOption('title', 'Répartition LV2 / Options par classe')
                    .setOption('titleTextStyle', { color: '#424242', fontSize: 14, bold: true })
                    .setOption('legend', { position: 'top', textStyle: { color: '#616161' } })
                    .setOption('colors', ['#ffcc80', '#80cbc4', '#bcaaa4']) // ESP, Autres LV2, Options
                    .setOption('isStacked', false) // Not stacked is clearer here
                    .setOption('width', 600).setOption('height', 400)
                    .setOption('hAxis', { title: 'Classes', textStyle: { color: '#616161' } })
                    .setOption('vAxis', { title: 'Nombre d\'élèves', textStyle: { color: '#616161' }, minValue: 0, gridlines: { count: -1 }})
                    // Position below the Criteria radar chart
                    .setPosition(endOptionRow + 1, 8, 10, 10)
                    .build();
                bilanSheet.insertChart(chartOptions);
            } catch (e) { Logger.log(`Erreur création graphique options/LV2: ${e.message}`); }
            currentRow = currentRow + 1; // Move down after table
        } else {
             Logger.log("Pas de données LV2/Options à afficher dans le bilan.");
             currentRow += 1;
        }


        // Final Cleanup and Formatting
        bilanSheet.autoResizeColumns(1, 7); // Auto-resize columns A-G
        bilanSheet.setColumnWidth(1, 120); // Set Classe column wider

        // Clean unused rows/cols (use the calculated currentRow)
        const lastMeaningfulRow = currentRow + 20; // Add buffer after last content
        const lastMeaningfulCol = 15; // Up to column O (adjust if charts go further right)
        cleanUnusedRowsAndColumns(bilanSheet, lastMeaningfulRow, lastMeaningfulCol, 10);

        // Activate the sheet
        bilanSheet.activate();
        Logger.log("Onglet BILAN TEST créé/mis à jour avec succès.");

    } catch (e) {
        Logger.log(`❌ ERREUR lors de la création du BILAN TEST: ${e.message} - ${e.stack}`);
        SpreadsheetApp.getUi().alert(`Erreur lors de la création du Bilan: ${e.message}`);
    }
}

// ================================================
// FONCTION PRINCIPALE D'EXÉCUTION (Lance tout) - (Inchangée)
// ================================================
// ... (Votre fonction executerParite originale complète ici) ...
function executerParite() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    Logger.log(">>> DÉBUT Exécution Complète Phase 3 (Répartition Capacité/Parité + Mise en forme) <<<");
    ss.toast("Début Phase 3 : Répartition...", "Progression", 5);

    // Étape 1: Logique répartition + écriture Phase 3
    Logger.log("Appel de repartirPhase3()...");
    const resultatRepartition = repartirPhase3();

    if (!resultatRepartition || !resultatRepartition.success) {
        const errorMessage = resultatRepartition.message || 'Erreur inconnue lors de la répartition Phase 3.';
        Logger.log(`Échec de la logique de répartition/écriture Phase 3: ${errorMessage}`);
        ss.toast(`Erreur Répartition: ${errorMessage}`, "Erreur Phase 3", 10);
        throw new Error(errorMessage);
    }

    ss.toast(`Répartition terminée (${resultatRepartition.nbAjoutes} élèves ajoutés). Début Mise en forme...`, "Progression", 10);

    // Étape 2: Mise en forme des onglets TEST avec collecte de stats
    Logger.log("Appel de miseEnFormeOngletsTest()...");
    const resultatMiseEnForme = miseEnFormeOngletsTest();

    if (!resultatMiseEnForme || !resultatMiseEnForme.success) {
       const errorMessage = resultatMiseEnForme.message || 'Erreur inconnue lors de la mise en forme.';
        Logger.log(`Échec de la mise en forme: ${errorMessage}`);
        ss.toast(`Erreur Mise en forme: ${errorMessage}`, "Erreur Phase 3", 10);
        throw new Error(errorMessage);
    }
    
    // Étape 3: Création du bilan avec les stats
    if (resultatMiseEnForme.classes && resultatMiseEnForme.classes.length > 0) {
      try {
        Logger.log("Création du bilan TEST...");
        ss.toast("Création du bilan statistique...", "Progression", 5);
        
        const bilanData = {
          classes: resultatMiseEnForme.classes,
          genreCounts: resultatMiseEnForme.genreCounts,
          criteresMoyennes: resultatMiseEnForme.criteresMoyennes,
          lv2Counts: resultatMiseEnForme.lv2Counts,
          optionsCounts: resultatMiseEnForme.optionsCounts
        };
        
        createBilanTest(ss, bilanData);
        Logger.log("Bilan TEST créé avec succès.");
      } catch (e) {
        Logger.log(`Erreur lors de la création du bilan: ${e.message}`);
        // On continue malgré l'erreur de bilan car la mise en forme est déjà faite
      }
    }

    ss.toast("Phase 3 (Répartition, Mise en forme et Bilan) terminée avec succès !", "Succès Phase 3", 10);

    Logger.log(">>> FIN Exécution Complète Phase 3 <<<");
    return { 
      success: true, 
      message: `Phase 3 complète exécutée. ${resultatRepartition.nbAjoutes} élèves ajoutés.` 
    };

  } catch (error) {
    Logger.log("❌ ERREUR Globale dans executerParite: " + error.message + " - " + error.stack);
    ss.toast(`Erreur Critique Phase 3: ${error.message}`, "Erreur Critique", 15);
    throw new Error(`Erreur Globale Phase 3: ${error.message}`);
  }
}

// ================================================
// FONCTION UTILITAIRE NETTOYAGE (Inchangée)
// ================================================
// ... (Votre fonction cleanUnusedRowsAndColumns originale complète ici) ...
/**
 * Supprime les colonnes et lignes inutilisées dans une feuille.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La feuille
 * @param {number} lastUsedRow Dernière ligne avec contenu/formule/format à garder
 * @param {number} lastUsedCol Dernière colonne avec contenu/formule/format à garder
 * @param {number} buffer Marge de lignes/colonnes vides à conserver
 */
function cleanUnusedRowsAndColumns(sheet, lastUsedRow, lastUsedCol, buffer = 5) {
  try {
    // Assurer que les paramètres sont des nombres valides et >= 1
    lastUsedRow = Math.max(1, parseInt(lastUsedRow) || 1);
    lastUsedCol = Math.max(1, parseInt(lastUsedCol) || 1);
    buffer = Math.max(0, parseInt(buffer) || 0);

    const sheetName = sheet.getName(); // Pour les logs

    // Supprimer les colonnes inutilisées
    const totalColumns = sheet.getMaxColumns();
    const targetCol = lastUsedCol + buffer; // Colonne limite à garder
    if (totalColumns > targetCol) {
      // Vérifier s'il y a des colonnes à supprimer (targetCol+1 <= totalColumns)
      if (targetCol + 1 <= totalColumns) {
          const numColsToDelete = totalColumns - targetCol;
          sheet.deleteColumns(targetCol + 1, numColsToDelete);
          // Logger.log(`      ${sheetName}: Deleted ${numColsToDelete} unused columns starting from ${targetCol + 1}.`);
      }
    } else if (totalColumns < lastUsedCol) {
        // Cas étrange : moins de colonnes physiques que la dernière colonne utilisée ?
        Logger.log(`      Warn (${sheetName}): Max columns (${totalColumns}) is less than last used column (${lastUsedCol}).`);
    }

    // Supprimer les lignes inutilisées
    const totalRows = sheet.getMaxRows();
    const targetRow = lastUsedRow + buffer; // Ligne limite à garder
    if (totalRows > targetRow) {
        // Vérifier s'il y a des lignes à supprimer
        if (targetRow + 1 <= totalRows) {
            const numRowsToDelete = totalRows - targetRow;
            sheet.deleteRows(targetRow + 1, numRowsToDelete);
            // Logger.log(`      ${sheetName}: Deleted ${numRowsToDelete} unused rows starting from ${targetRow + 1}.`);
        }
    } else if (totalRows < lastUsedRow) {
         Logger.log(`      Warn (${sheetName}): Max rows (${totalRows}) is less than last used row (${lastUsedRow}).`);
    }

  } catch (e) {
    // Logguer l'erreur mais ne pas arrêter le script pour ça
    Logger.log(`      Warning: Error cleaning unused rows/cols in ${sheet.getName()} (lastRow: ${lastUsedRow}, lastCol: ${lastUsedCol}): ${e.message}`);
  }
}
function debugHeadersTEST() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getTestSheetsNames(); // ta fonction existante
  sheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { Logger.log(`→ ${name} : feuille introuvable`); return; }
    const raw = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    Logger.log(`→ ${name}: [${ raw.map(x=>String(x)).join(" | ") }]`);
  });
}
function cacherColonnesABCDansOngletsTEST() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var toutesLesFeuilles = spreadsheet.getSheets();
  
  for (var i = 0; i < toutesLesFeuilles.length; i++) {
    var feuille = toutesLesFeuilles[i];
    var nomFeuille = feuille.getName();
    
    if (nomFeuille.endsWith("TEST")) {
      feuille.hideColumns(1, 3);
    }
  }
}