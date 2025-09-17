/**
 * Initialisation.gs
 * Gère le processus d'initialisation et de configuration initiale du classeur.
 * Version universelle pour tous les niveaux (sources ≠ destinations)
 */

// --- Fonctions Utilitaires ---

/**
 * Nettoie l'input utilisateur pour les listes (virgules, espaces, majuscules, doublons).
 * @param {string} inputString - La chaîne entrée par l'utilisateur.
 * @return {string[]} Un tableau de chaînes nettoyées.
 */
function nettoyerListeInput(inputString) {
  if (!inputString || inputString.trim() === "") {
    return [];
  }
  const items = inputString.split(',')
                       .map(item => item.trim().toUpperCase())
                       .filter(item => item !== "");
  return [...new Set(items)].sort(); // Enlève doublons et trie
}

/**
 * Vérifie le mot de passe administrateur (version simple, dépend de getConfig).
 * @param {string} action - Description de l'action à protéger.
 * @return {boolean} Vrai si le mot de passe est correct.
 */
function verifierMotDePasse(action) {
  const config = getConfig(); // Appel la version centrale
  const ui = SpreadsheetApp.getUi();

  const mdpResponse = ui.prompt(
    `Accès ${action}`,
    "Veuillez entrer le mot de passe administrateur :",
    ui.ButtonSet.OK_CANCEL
  );

  if (mdpResponse.getSelectedButton() !== ui.Button.OK) {
    return false;
  }

  const motDePasse = mdpResponse.getResponseText();
  // Utilise le mot de passe de la config
  const mdpCorrect = config.ADMIN_PASSWORD;

  if (motDePasse !== mdpCorrect) {
    ui.alert(
      "Accès refusé",
      "Mot de passe incorrect.",
      ui.ButtonSet.OK
    );
    return false;
  }

  // Journaliser la tentative d'accès réussie (si logAction est défini)
  try {
      logAction(`Accès ${action} validé`);
  } catch(e) { Logger.log(`Accès ${action} validé (logAction non disponible)`); }
  return true;
}

// --- Fonction Principale d'Initialisation via Menu ---

/**
 * Ouvre la boîte de dialogue d'initialisation du système.
 * Version UNIVERSELLE pour tous les niveaux.
 */
function ouvrirInitialisation() {
  // 1. Vérifier mot de passe
  if (!verifierMotDePasse("Administration (Initialisation)")) return;

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 2. Confirmer réinitialisation si nécessaire
  try {
      const structureSheet = ss.getSheetByName(CONFIG.SHEETS.STRUCTURE);
      if (structureSheet) {
        const response = ui.alert(
          "Réinitialisation du système",
          "Ce classeur semble déjà initialisé. La réinitialisation effacera la configuration et les données existantes.\n\nVoulez-vous continuer ?",
          ui.ButtonSet.YES_NO
        );
        if (response !== ui.Button.YES) {
            ui.alert("Initialisation annulée.");
            return;
        }
         Logger.log("Confirmation réinitialisation OK.");
      }
  } catch (e) {
      Logger.log(`Erreur vérification feuille _STRUCTURE: ${e}`);
  }

  // 3. Demander le niveau
  const niveaux = ["6°", "5°", "4°", "3°"];
  const niveauResponse = ui.prompt(
    "Initialisation - Étape 1/5 : Niveau",
    `Niveau scolaire (${niveaux.join(", ")}) :`,
    ui.ButtonSet.OK_CANCEL
  );
  if (niveauResponse.getSelectedButton() !== ui.Button.OK) return;
  const niveau = niveauResponse.getResponseText().trim();
  if (!niveaux.includes(niveau)) {
    ui.alert("Niveau invalide", `Veuillez entrer un niveau valide: ${niveaux.join(", ")}.`, ui.ButtonSet.OK);
    return;
  }
  Logger.log(`Niveau : ${niveau}`);

  // 4. Demander le nombre de classes/écoles SOURCES
  const sourceLabel = niveau === "6°" ? "écoles primaires" : `classes de ${determinerNiveauSource(niveau)}`;
  const nbSourcesResponse = ui.prompt(
    "Initialisation - Étape 2/5 : Sources",
    `Nombre de ${sourceLabel} sources (ex: ${niveau === "6°" ? "9" : "6"}) :`,
    ui.ButtonSet.OK_CANCEL
  );
  if (nbSourcesResponse.getSelectedButton() !== ui.Button.OK) return;
  const nbSources = parseInt(nbSourcesResponse.getResponseText(), 10);
  if (isNaN(nbSources) || nbSources <= 0 || nbSources > 20) {
    ui.alert("Nombre invalide", "Veuillez entrer un nombre valide (1-20).", ui.ButtonSet.OK);
    return;
  }
  Logger.log(`Nb Sources : ${nbSources}`);

  // 5. Demander le nombre de classes DESTINATIONS
  const nbDestResponse = ui.prompt(
    "Initialisation - Étape 3/5 : Destinations",
    `Nombre de classes de ${niveau} à créer (ex: ${nbSources === 9 ? "5" : nbSources}) :`,
    ui.ButtonSet.OK_CANCEL
  );
  if (nbDestResponse.getSelectedButton() !== ui.Button.OK) return;
  const nbDest = parseInt(nbDestResponse.getResponseText(), 10);
  if (isNaN(nbDest) || nbDest <= 0 || nbDest > 15) {
    ui.alert("Nombre invalide", "Veuillez entrer un nombre valide (1-15).", ui.ButtonSet.OK);
    return;
  }
  Logger.log(`Nb Destinations : ${nbDest}`);

  // 6. Demander les LV2
  const lv2Response = ui.prompt(
    "Initialisation - Étape 4/5 : LV2",
    "Entrez les SIGLES des LV2 proposées, séparés par des virgules (Ex: ESP,ITA,ALL). Laissez vide si aucune.",
    ui.ButtonSet.OK_CANCEL
  );
  if (lv2Response.getSelectedButton() !== ui.Button.OK) return;
  const lv2Options = nettoyerListeInput(lv2Response.getResponseText());
  Logger.log(`Options LV2 saisies : ${lv2Options.join(', ') || 'Aucune'}`);

  // 7. Demander les Options
  const optResponse = ui.prompt(
    "Initialisation - Étape 5/5 : Options Spécifiques",
    "Entrez les SIGLES des options spécifiques, séparés par des virgules (Ex: LATIN,GREC,CHAV). Laissez vide si aucune.",
    ui.ButtonSet.OK_CANCEL
  );
  if (optResponse.getSelectedButton() !== ui.Button.OK) return;
  const optOptions = nettoyerListeInput(optResponse.getResponseText());
  Logger.log(`Options spécifiques saisies : ${optOptions.join(', ') || 'Aucune'}`);

  // 8. Lancer l'initialisation complète
  initialiserSysteme(niveau, nbSources, nbDest, lv2Options, optOptions);
}

/**
 * Détermine le niveau source en fonction du niveau destination
 * @param {string} niveau - Niveau de destination
 * @return {string} Le niveau source
 */
function determinerNiveauSource(niveau) {
  switch (niveau) {
    case "6°": return "CM2";
    case "5°": return "6°";
    case "4°": return "5°";
    case "3°": return "4°";
    default: return "niveau précédent";
  }
}

// --- Processus d'Initialisation Principal ---

/**
 * Initialise le système complet avec tous les onglets et configurations nécessaires.
 * Version UNIVERSELLE supportant nbSources ≠ nbDestinations.
 * @param {string} niveau - Niveau scolaire ("6°", "5°", "4°", "3°").
 * @param {number} nbSources - Nombre de classes/écoles sources.
 * @param {number} nbDest - Nombre de classes destinations.
 * @param {string[]} lv2Options - Tableau des sigles LV2.
 * @param {string[]} optOptions - Tableau des sigles Options.
 */
function initialiserSysteme(niveau, nbSources, nbDest, lv2Options, optOptions) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  Logger.log(`--- Début Initialisation Système ---`);
  Logger.log(`Niveau: ${niveau}, Sources: ${nbSources}, Destinations: ${nbDest}`);
  Logger.log(`LV2: ${lv2Options.join(',')}, OPT: ${optOptions.join(',')}`);
  
  // Configuration du spinner progressif
  const etapes = [
    { nom: "Préparation", pourcentage: 0 },
    { nom: "Configuration", pourcentage: 15 },
    { nom: "Onglets système", pourcentage: 25 },
    { nom: "Structure", pourcentage: 40 },
    { nom: "Accueil", pourcentage: 50 },
    { nom: "Sources", pourcentage: 70 },
    { nom: "Listes déroulantes", pourcentage: 85 },
    { nom: "Mise en forme", pourcentage: 95 },
    { nom: "Finalisation", pourcentage: 100 }
  ];
  
  // Fonction pour afficher le spinner avec l'étape en cours
  function afficherSpinner(etapeIndex) {
    const etape = etapes[etapeIndex];
    const pourcentage = etape.pourcentage;
    
    // Création de la barre visuelle
    const longueurBarre = 20;
    const barreRemplie = Math.floor((pourcentage / 100) * longueurBarre);
    const barre = '▓'.repeat(barreRemplie) + '░'.repeat(longueurBarre - barreRemplie);
    
    // Affichage du message
    const message = `${etape.nom} : ${pourcentage}%\n${barre}`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Initialisation', -1);
  }

  try {
    // Début du spinner
    afficherSpinner(0);
    
    // Préparation: Nettoyer les anciens onglets 
    supprimerAnciensOngletsNonSysteme();
    afficherSpinner(1);

    // 1. Créer/Réinitialiser _CONFIG avec les nouvelles données
    Logger.log("Étape 1: Création/MàJ Onglet _CONFIG...");
    creerOuMajOngletConfig(niveau, lv2Options, optOptions);
    afficherSpinner(2);

    // 2. Créer/Réinitialiser les onglets système (_JOURNAL, _BACKUP)
    Logger.log("Étape 2: Création Onglets Système...");
    creerOngletsSysteme();
    afficherSpinner(3);

    // 3. Créer/Réinitialiser _STRUCTURE
    Logger.log("Étape 3: Création Onglet _STRUCTURE...");
    creerOngletStructure(niveau, nbSources, nbDest);
    afficherSpinner(4);

    // 4. Créer/Réinitialiser ACCUEIL
    Logger.log("Étape 4: Création Onglet ACCUEIL...");
    creerOngletPresentation();
    afficherSpinner(5);

    // 5. Créer/Réinitialiser les onglets sources et CONSOLIDATION
    Logger.log("Étape 5: Création Onglets Sources & CONSOLIDATION...");
    creerOngletsSourcesVides(niveau, nbSources);
    afficherSpinner(6);

    // 6. Ajouter/MàJ Listes Déroulantes & Formatage Associé
    Logger.log("Étape 6: Ajout Listes Déroulantes & Formatage...");
    ajouterListesDeroulantes();
    afficherSpinner(7);

    // 7. Appliquer Mise en Forme Générale (Filtre/Fige)
    Logger.log("Étape 7: Mise en Forme Générale (Filtre/Fige)...");
    miseEnFormeGeneraleSources();
    afficherSpinner(8);

    // --- Finalisation ---
    SpreadsheetApp.flush(); // Forcer l'application des changements
    
    // Activer l'onglet Accueil à la fin
    try {
        const accueilSheet = ss.getSheetByName("ACCUEIL");
        if (accueilSheet) ss.setActiveSheet(accueilSheet);
    } catch (e) { /* Ignorer si ACCUEIL n'existe pas */ }

    SpreadsheetApp.getActiveSpreadsheet().toast('Initialisation terminée !', 'Succès', 5);
    Logger.log("--- Initialisation Système Terminée avec Succès ---");

    // Afficher une confirmation finale
    const messageFinal = `Initialisation Réussie !\n\n` +
                         `Niveau: ${niveau}\n` +
                         `Sources: ${nbSources} ${niveau === "6°" ? "écoles" : "classes"}\n` +
                         `Destinations: ${nbDest} classes de ${niveau}\n` +
                         `LV2: ${lv2Options.join(', ') || 'Aucune'}\n` +
                         `Options: ${optOptions.join(', ') || 'Aucune'}\n\n` +
                         `Prochaines étapes:\n` +
                         `1. Vérifier _STRUCTURE\n` +
                         `2. Importer les données dans les onglets sources\n` +
                         `3. Utiliser la Console pour la répartition`;
    ui.alert("Initialisation Terminée", messageFinal, ui.ButtonSet.OK);

    // Journaliser l'action
    logAction(`Initialisation: ${niveau}, ${nbSources} sources → ${nbDest} destinations, LV2[${lv2Options.join(',')}], OPT[${optOptions.join(',')}]`);

    // Ouvrir la Console (si elle existe)
    try { if (typeof ouvrirConsole === "function") { ouvrirConsole(); } } catch(e) { Logger.log("Console non disponible"); }

  } catch (e) {
    Logger.log(`!!! ERREUR FATALE LORS DE L'INITIALISATION !!!\nErreur: ${e.toString()}\nStack: ${e.stack}`);
    SpreadsheetApp.getActiveSpreadsheet().toast('Erreur pendant l\'initialisation!', 'ERREUR', 10);
    ui.alert(
      "Erreur lors de l'initialisation",
      `Une erreur majeure s'est produite : ${e.message}\nConsultez les logs (Extensions > Apps Script > Exécutions) pour plus de détails.`,
      ui.ButtonSet.OK
    );
  }
}

// --- Sous-Fonctions d'Initialisation ---

/**
 * Supprime tous les onglets sauf ceux définis comme système/accueil.
 * À utiliser avec prudence lors d'une réinitialisation complète.
 */
function supprimerAnciensOngletsNonSysteme() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    // Liste des onglets à ne PAS supprimer
    const ongletsAPreserver = ["ACCUEIL", ...Object.values(CONFIG.SHEETS)];

    Logger.log("Nettoyage des anciens onglets non-système...");
    let deletedCount = 0;
    // Itérer en sens inverse pour éviter les problèmes d'index lors de la suppression
    for (let i = allSheets.length - 1; i >= 0; i--) {
        const sheet = allSheets[i];
        const sheetName = sheet.getName();
        // S'il reste plus d'une feuille ET que la feuille n'est pas dans la liste à préserver
        if (ss.getNumSheets() > 1 && !ongletsAPreserver.includes(sheetName)) {
            Logger.log(` - Suppression onglet: ${sheetName}`);
            try {
                ss.deleteSheet(sheet);
                deletedCount++;
            } catch (e) {
                Logger.log(`   Impossible de supprimer ${sheetName}: ${e}`);
            }
        } else if (ongletsAPreserver.includes(sheetName)) {
            // Logger.log(` - Conservation onglet: ${sheetName}`);
        } else {
            Logger.log(` - Impossible de supprimer ${sheetName} (dernière feuille?).`);
        }
    }
    if (deletedCount > 0) SpreadsheetApp.flush();
    Logger.log(`Nettoyage terminé. ${deletedCount} onglets supprimés.`);
}

/**
 * Crée ou Met à Jour l'onglet _CONFIG.
 * @param {string} niveau - Niveau scolaire.
 * @param {string[]} lv2Options - Tableau des sigles LV2.
 * @param {string[]} optOptions - Tableau des sigles Options.
 */
function creerOuMajOngletConfig(niveau, lv2Options, optOptions) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheetName = CONFIG.SHEETS.CONFIG;
  let configSheet = ss.getSheetByName(configSheetName);

  const lv2String = lv2Options.join(',');
  const optString = optOptions.join(',');

  // Structure attendue des données dans _CONFIG
  const configMap = {
    "NIVEAU": { value: niveau, description: "Niveau scolaire principal du classeur" },
    "ADMIN_PASSWORD": { value: CONFIG.ADMIN_PASSWORD_DEFAULT, description: "Mot de passe pour fonctions admin" },
    "MAX_SWAPS": { value: CONFIG.MAX_SWAPS, description: "Limite d'échanges pour l'optimisation" },
    "PARITY_TOLERANCE": { value: CONFIG.PARITY_TOLERANCE, description: "Écart max autorisé F/G par classe" },
    "AUTO_RENAME": { value: "NON", description: "Renommer automatiquement onglets DEF (OUI/NON)" },
    "LV2": { value: lv2String, description: "Liste des LV2 disponibles (séparées par virgule)" },
    "OPT": { value: optString, description: "Liste des Options spécifiques (séparées par virgule)" },
  };

  if (!configSheet) {
    // Créer l'onglet s'il n'existe pas
    Logger.log(`Création onglet ${configSheetName}...`);
    configSheet = ss.insertSheet(configSheetName);
    // Appliquer le formatage des en-têtes
    configSheet.getRange("A1:C1").setValues([["PARAMETRE", "VALEUR", "DESCRIPTION"]])
               .setFontWeight("bold").setBackground("#d5dbdb");
    configSheet.setFrozenRows(1);
  } else {
     Logger.log(`Mise à jour onglet ${configSheetName}...`);
  }

  // Lire les paramètres existants pour ne mettre à jour que les valeurs
  const existingData = configSheet.getDataRange().getValues();
  const existingParams = {};
  for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0]) {
          existingParams[existingData[i][0]] = { rowIndex: i + 1, value: existingData[i][1] };
      }
  }

  // Préparer les données à écrire (mise à jour ou ajout)
  const dataToWrite = [];
  const paramsToWrite = Object.keys(configMap);

  paramsToWrite.forEach(param => {
      const newValue = configMap[param].value;
      const description = configMap[param].description;
      if (existingParams[param]) {
          // Mettre à jour la valeur existante si différente
          if (existingParams[param].value !== newValue) {
              configSheet.getRange(existingParams[param].rowIndex, 2).setValue(newValue);
          }
          // Mettre à jour la description
           configSheet.getRange(existingParams[param].rowIndex, 3).setValue(description);
      } else {
          // Ajouter la ligne si le paramètre n'existe pas
          dataToWrite.push([param, newValue, description]);
      }
  });

  // Ajouter les nouvelles lignes si nécessaire
  if (dataToWrite.length > 0) {
      configSheet.getRange(configSheet.getLastRow() + 1, 1, dataToWrite.length, 3).setValues(dataToWrite);
  }

  // Mise en forme globale et protection
  configSheet.getRange(1, 1, configSheet.getLastRow(), 3).setBorder(true,true,true,true,true,true);
  configSheet.setColumnWidth(1, 180);
  configSheet.setColumnWidth(2, 150);
  configSheet.setColumnWidth(3, 400);
  configSheet.hideSheet();
  protegerFeuille(configSheet, "Configuration système");
  Logger.log(`Onglet ${configSheetName} configuré.`);
}

/**
 * Crée les onglets système cachés (_JOURNAL, _BACKUP)
 */
function creerOngletsSysteme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(" - Création/MàJ _JOURNAL et _BACKUP...");

  // _JOURNAL
  let journalSheet = ss.getSheetByName(CONFIG.SHEETS.JOURNAL);
  if (journalSheet) ss.deleteSheet(journalSheet);
  journalSheet = ss.insertSheet(CONFIG.SHEETS.JOURNAL);
  journalSheet.getRange(1, 1, 1, 3).setValues([["TIMESTAMP", "ACTION", "UTILISATEUR"]]);
  journalSheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#d5dbdb");
  journalSheet.setColumnWidth(1, 180).setColumnWidth(2, 400).setColumnWidth(3, 150);
  journalSheet.setFrozenRows(1);
  journalSheet.hideSheet();

  // _BACKUP
  let backupSheet = ss.getSheetByName(CONFIG.SHEETS.BACKUP);
  if (backupSheet) ss.deleteSheet(backupSheet);
  backupSheet = ss.insertSheet(CONFIG.SHEETS.BACKUP);
  backupSheet.getRange(1, 1, 1, 3).setValues([["TIMESTAMP", "DESCRIPTION", "DONNEES_JSON"]]);
  backupSheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#d5dbdb");
  backupSheet.setColumnWidth(1, 180).setColumnWidth(2, 250).setColumnWidth(3, 800);
  backupSheet.setFrozenRows(1);
  backupSheet.hideSheet();

  logAction("Création/Réinitialisation des onglets système (_JOURNAL, _BACKUP)");
}

/**
 * Crée l'onglet _STRUCTURE - Version UNIVERSELLE
 * @param {string} niveau - Niveau de destination
 * @param {number} nbSources - Nombre de classes/écoles sources
 * @param {number} nbDest - Nombre de classes destinations
 */
function creerOngletStructure(niveau, nbSources, nbDest) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const structureSheetName = CONFIG.SHEETS.STRUCTURE;
  Logger.log(` - Création/MàJ ${structureSheetName}...`);
  Logger.log(`   Sources: ${nbSources}, Destinations: ${nbDest}`);

  let structureSheet = ss.getSheetByName(structureSheetName);
  if (structureSheet) ss.deleteSheet(structureSheet);
  structureSheet = ss.insertSheet(structureSheetName);

  structureSheet.getRange("A:E").setNumberFormat('@');

  const headers = ["Type", "Nom Classe", "Capacité Max", "Prof Principal", "Observations"];
  structureSheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold").setBackground("#d3d3d3");
  structureSheet.setFrozenRows(1);

  let rowIndex = 2;
  const prefixeSource = determinerPrefixeSource(niveau);

  // Créer les lignes SOURCE
  for (let i = 1; i <= nbSources; i++) {
    const nomClasseSource = `${prefixeSource}${i}`;
    const observation = niveau === "6°" ? "École primaire" : "Classe source";
    structureSheet.getRange(rowIndex++, 1, 1, headers.length).setValues([
      ["SOURCE", nomClasseSource, "", "", observation]
    ]);
  }

  // Créer les lignes TEST et DEF
  for (let i = 1; i <= nbDest; i++) {
    const nomClasseTest = `${niveau}${i} TEST`;
    structureSheet.getRange(rowIndex++, 1, 1, headers.length).setValues([
      ["TEST", nomClasseTest, "30", "", "Classe cible temporaire"]
    ]);
    const nomClasseDef = `${niveau}${i} DEF`;
    structureSheet.getRange(rowIndex++, 1, 1, headers.length).setValues([
      ["DEF", nomClasseDef, "30", "", "Classe cible définitive"]
    ]);
  }

  // Formatage final
  structureSheet.getRange("C:C").setHorizontalAlignment("center");
  structureSheet.autoResizeColumns(1, headers.length);
  structureSheet.getRange(1, 1, structureSheet.getLastRow(), headers.length).setBorder(true,true,true,true,true,true);
  protegerFeuille(structureSheet, "Structure des classes - Modifier avec soin", true);
  
  Logger.log(`_STRUCTURE créée avec ${nbSources} sources et ${nbDest} destinations`);
}

/**
 * Détermine le préfixe des onglets sources en fonction du niveau.
 * @param {string} niveau - Niveau scolaire (6°, 5°, 4°, 3°).
 * @return {string} Le préfixe approprié.
 */
function determinerPrefixeSource(niveau) {
  switch (niveau) {
    case "6°": return "ECOLE";
    case "5°": return "6°";
    case "4°": return "5°";
    case "3°": return "4°";
    default: Logger.log(`WARN: Préfixe source inconnu pour niveau ${niveau}`); return "SOURCE";
  }
}

/**
 * Crée les onglets sources vides et CONSOLIDATION avec formatage optimisé et limites de taille
 * @param {string} niveau - Niveau scolaire
 * @param {number} nbClasses - Nombre de classes sources
 */
function creerOngletsSourcesVides(niveau, nbClasses) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prefixeSource = determinerPrefixeSource(niveau);
  Logger.log(` - Création/MàJ CONSOLIDATION et ${nbClasses} onglets sources (${prefixeSource})...`);

  // En-tête standardisé
  const entete = [
    "ID_ELEVE", "NOM", "PRENOM", "NOM_PRENOM", "SEXE", "LV2", "OPT",
    "COM", "TRA", "PART", "ABS", "DISPO", "ASSO", "DISSO", "SOURCE"
  ];
  
  // Récupération des index de colonnes
  const sourceColIndex = entete.indexOf("SOURCE") + 1;
  const nomPrenomColIndex = entete.indexOf("NOM_PRENOM") + 1;
  const nomColIndex = entete.indexOf("NOM") + 1;
  const prenomColIndex = entete.indexOf("PRENOM") + 1;
  const sexeColIndex = entete.indexOf("SEXE") + 1;
  const lv2ColIndex = entete.indexOf("LV2") + 1;
  const optColIndex = entete.indexOf("OPT") + 1;
  const assoColIndex = entete.indexOf("ASSO") + 1;
  const dissoColIndex = entete.indexOf("DISSO") + 1;
  const dispoColIndex = entete.indexOf("DISPO") + 1;
  const critColsIndexes = ["COM", "TRA", "PART", "ABS"].map(c => entete.indexOf(c) + 1);
  
  // Définition des largeurs optimisées pour chaque colonne
  const columnWidths = {
    [nomColIndex]: 120,
    [prenomColIndex]: 120,
    [nomPrenomColIndex]: 180,
    [sexeColIndex]: 50,
    [lv2ColIndex]: 60,
    [optColIndex]: 65,
    [assoColIndex]: 80,
    [dissoColIndex]: 80,
    [sourceColIndex]: 60,
    [dispoColIndex]: 90
  };
  
  // Ajouter les largeurs pour les critères
  critColsIndexes.forEach(idx => {
    columnWidths[idx] = 50;
  });
  
  // Nombre initial de lignes limité
  const initialRows = 35;
  
  // Nombre maximal de colonnes nécessaires
  const maxNeededCol = Math.max(...entete.map((_, i) => i + 1));

  // ----- CRÉATION ONGLET CONSOLIDATION -----
  const consolidationSheetName = CONFIG.SHEETS.CONSOLIDATION;
  let consolidationSheet = ss.getSheetByName(consolidationSheetName);
  if (consolidationSheet) ss.deleteSheet(consolidationSheet);
  consolidationSheet = ss.insertSheet(consolidationSheetName);
  
  // Limiter nombre de lignes et colonnes
  ajusterTailleOnglet(consolidationSheet, initialRows, maxNeededCol);
  
  // Formatage en-tête et données
  formaterOnglet(consolidationSheet, entete, initialRows, columnWidths, "#c9daf8");

  // ----- CRÉATION ONGLETS SOURCES -----
  for (let i = 1; i <= nbClasses; i++) {
    const nomOnglet = `${prefixeSource}${i}`;
    let sheet = ss.getSheetByName(nomOnglet);
    if (sheet) ss.deleteSheet(sheet);
    sheet = ss.insertSheet(nomOnglet);
    
    // Limiter nombre de lignes et colonnes
    ajusterTailleOnglet(sheet, initialRows, maxNeededCol);
    
    // Formatage en-tête et données
    formaterOnglet(sheet, entete, initialRows, columnWidths, "#d9ead3");
    
    // Formule NOM_PRENOM améliorée avec SIERREUR
    ajouterFormules(sheet, initialRows);
    
    // Préremplir SOURCE
    if (sourceColIndex > 0) {
      sheet.getRange(2, sourceColIndex, initialRows - 1, 1).setValue(nomOnglet);
    }
  }

  // Ajouter menu pour gestion des lignes
  ajouterMenuGestionLignes();

  logAction(`Création optimisée de ${nbClasses} onglets sources (${prefixeSource}) et CONSOLIDATION avec ${initialRows} lignes`);
}

/**
 * Ajuste la taille d'un onglet pour optimiser l'espace
 * @param {Sheet} sheet - L'onglet à ajuster
 * @param {number} rows - Nombre de lignes souhaitées
 * @param {number} cols - Nombre de colonnes souhaitées
 */
function ajusterTailleOnglet(sheet, rows, cols) {
  // Ajuster les lignes
  const currentRows = sheet.getMaxRows();
  if (currentRows > rows) {
    sheet.deleteRows(rows + 1, currentRows - rows);
  } else if (currentRows < rows) {
    sheet.insertRowsAfter(currentRows, rows - currentRows);
  }
  
  // Ajuster les colonnes
  const currentCols = sheet.getMaxColumns();
  if (currentCols > cols) {
    sheet.deleteColumns(cols + 1, currentCols - cols);
  } else if (currentCols < cols) {
    sheet.insertColumnsAfter(currentCols, cols - currentCols);
  }
}

/**
 * Formater un onglet avec les en-têtes et largeurs optimisées
 * @param {Sheet} sheet - L'onglet à formater
 * @param {string[]} entete - Tableau des en-têtes
 * @param {number} initialRows - Nombre de lignes
 * @param {Object} columnWidths - Objet avec les largeurs de colonnes
 * @param {string} headerColor - Couleur d'arrière-plan pour l'en-tête
 */
function formaterOnglet(sheet, entete, initialRows, columnWidths, headerColor) {
  // En-tête avec couleur et style gras
  sheet.getRange(1, 1, 1, entete.length).setValues([entete])
      .setFontWeight("bold").setBackground(headerColor);
  sheet.setFrozenRows(1);
  
  // Format texte pour les données
  sheet.getRange(1, 1, initialRows, entete.length).setNumberFormat('@');
  
  // Appliquer les largeurs de colonnes optimisées
  for (const [colIdx, width] of Object.entries(columnWidths)) {
    if (colIdx > 0) {
      sheet.setColumnWidth(parseInt(colIdx), width);
    }
  }
  
  // Alignement des colonnes de critères au centre
  entete.forEach((header, idx) => {
    if (['COM', 'TRA', 'PART', 'ABS'].includes(header)) {
      sheet.getRange(1, idx + 1, initialRows, 1).setHorizontalAlignment('center');
    }
  });
}

/**
 * Ajoute un menu pour la gestion des lignes des onglets
 */
function ajouterMenuGestionLignes() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Gestion Lignes')
      .addItem('Ajouter 30 lignes', 'ajouterLignesSupplementaires')
      .addItem('Ajouter 100 lignes', 'ajouterBeaucoupLignes')
      .addItem('Optimiser espace (supprimer lignes vides)', 'optimiserEspace')
      .addToUi();
  } catch (e) {
    Logger.log(`Erreur lors de l'ajout du menu: ${e}`);
  }
}

/**
 * Ajoute 30 lignes à la feuille active en préservant les formules et formatages
 */
function ajouterLignesSupplementaires() {
  ajouterLignes(30);
}

/**
 * Ajoute 100 lignes à la feuille active en préservant les formules et formatages
 */
function ajouterBeaucoupLignes() {
  ajouterLignes(100);
}

/**
 * Fonction générique pour ajouter des lignes
 * @param {number} nombreLignes - Nombre de lignes à ajouter
 */
function ajouterLignes(nombreLignes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Veuillez sélectionner une feuille avant d'ajouter des lignes.");
    return;
  }
  
  try {
    // Récupérer l'en-tête
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Trouver le dernier index de ligne contenant des données
    const lastRow = Math.max(2, sheet.getLastRow());
    
    // Ajouter les lignes
    sheet.insertRowsAfter(lastRow, nombreLignes);
    
    // Appliquer le formatage texte
    sheet.getRange(lastRow + 1, 1, nombreLignes, sheet.getLastColumn()).setNumberFormat('@');
    
    // Propager les formules et valeurs spéciales
    propagerFormulesEtValeurs(sheet, headers, lastRow, nombreLignes);
    
    // Propager les règles de validation
    propagerValidations(sheet, lastRow, nombreLignes);
    
    SpreadsheetApp.getUi().alert(`${nombreLignes} lignes ont été ajoutées à la feuille "${sheet.getName()}".`);
  } catch (e) {
    Logger.log(`Erreur lors de l'ajout de lignes: ${e}`);
    SpreadsheetApp.getUi().alert(`Erreur lors de l'ajout de lignes: ${e.message}`);
  }
}

/**
 * Propage les formules et valeurs spéciales aux nouvelles lignes
 * @param {Sheet} sheet - Feuille active
 * @param {Array} headers - Tableau des en-têtes
 * @param {number} lastRow - Dernière ligne avec données
 * @param {number} nombreLignes - Nombre de lignes ajoutées
 */
function propagerFormulesEtValeurs(sheet, headers, lastRow, nombreLignes) {
  // Rechercher les colonnes spéciales
  const nomPrenomIndex = headers.indexOf("NOM_PRENOM") + 1;
  const sourceIndex = headers.indexOf("SOURCE") + 1;
  
  // Propager formule NOM_PRENOM
  if (nomPrenomIndex > 0) {
    const referenceFormule = sheet.getRange(2, nomPrenomIndex).getFormula();
    if (referenceFormule) {
      for (let i = 0; i < nombreLignes; i++) {
        const rowNum = lastRow + 1 + i;
        // Adapter le numéro de ligne dans la formule
        const formuleAjustee = referenceFormule.replace(/2/g, rowNum.toString());
        sheet.getRange(rowNum, nomPrenomIndex).setFormula(formuleAjustee);
      }
    }
  }
  
  // Propager valeur SOURCE
  if (sourceIndex > 0) {
    const sourceValue = sheet.getRange(2, sourceIndex).getValue();
    if (sourceValue) {
      sheet.getRange(lastRow + 1, sourceIndex, nombreLignes, 1).setValue(sourceValue);
    }
  }
}

/**
 * Propage les règles de validation aux nouvelles lignes
 * @param {Sheet} sheet - Feuille active
 * @param {number} lastRow - Dernière ligne avec données
 * @param {number} nombreLignes - Nombre de lignes ajoutées
 */
function propagerValidations(sheet, lastRow, nombreLignes) {
  // Parcourir toutes les colonnes
  for (let col = 1; col <= sheet.getLastColumn(); col++) {
    // Vérifier si la deuxième ligne a une validation pour cette colonne
    const validation = sheet.getRange(2, col).getDataValidation();
    if (validation) {
      sheet.getRange(lastRow + 1, col, nombreLignes, 1).setDataValidation(validation);
    }
  }
}

/**
 * Optimise l'espace en supprimant les lignes vides en fin de feuille
 * mais en conservant un minimum de lignes
 */
function optimiserEspace() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Veuillez sélectionner une feuille à optimiser.");
    return;
  }
  
  try {
    // Trouver la dernière ligne avec des données
    let lastDataRow = sheet.getLastRow();
    
    // Conserver un minimum de 35 lignes, même si vides
    const minRows = 35;
    const currentRows = sheet.getMaxRows();
    
    if (lastDataRow < minRows) {
      lastDataRow = minRows;
    }
    
    // S'il y a des lignes en excès, les supprimer
    if (currentRows > lastDataRow) {
      const excessRows = currentRows - lastDataRow;
      sheet.deleteRows(lastDataRow + 1, excessRows);
      SpreadsheetApp.getUi().alert(`Optimisation effectuée : ${excessRows} lignes vides supprimées.\nLa feuille contient maintenant ${lastDataRow} lignes.`);
    } else {
      SpreadsheetApp.getUi().alert("Aucune optimisation nécessaire. Toutes les lignes contiennent des données ou sont dans la limite minimale.");
    }
  } catch (e) {
    Logger.log(`Erreur lors de l'optimisation: ${e}`);
    SpreadsheetApp.getUi().alert(`Erreur lors de l'optimisation: ${e.message}`);
  }
}

/**
 * Applique la mise en forme générale (Fige ligne, Filtre) aux onglets sources.
 * Le formatage conditionnel est géré par ajouterListesDeroulantesEtFormatage.
 */
function miseEnFormeGeneraleSources() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const config = getConfig();
  const niveau = config.NIVEAU;
  Logger.log(`[MiseEnForme Générale] Démarrage pour niveau ${niveau}...`);

  const ongletsSpeciauxAIgnorer = ["ACCUEIL", ...Object.values(CONFIG.SHEETS)];
  const prefixeSource = determinerPrefixeSource(niveau);

  for (const sheet of allSheets) {
    const nomOnglet = sheet.getName();
    const lcNomOnglet = nomOnglet.toLowerCase();

    if (ongletsSpeciauxAIgnorer.includes(nomOnglet) ||
        lcNomOnglet.includes("test") || lcNomOnglet.includes("def") ||
        lcNomOnglet.includes("consolidation") || lcNomOnglet.includes("bilan") ||
        lcNomOnglet.includes("stats")) {
      continue;
    }

    if (nomOnglet.startsWith(prefixeSource)) {
      Logger.log(` - ${nomOnglet}: Application Fige/Filtre...`);
      try {
        const lastCol = sheet.getLastColumn();
        const lastRowData = sheet.getLastRow();

        if (lastCol === 0 || lastRowData <= 0) continue;

        // Figer ligne 1
        sheet.setFrozenRows(1);

        // Appliquer Filtre
        if (lastRowData > 1) {
          const filterRange = sheet.getRange(1, 1, lastRowData, lastCol);
          const existingFilter = sheet.getFilter();
          if (existingFilter) existingFilter.remove();
          filterRange.createFilter();
        }

      } catch (e) {
        Logger.log(`   - ERREUR MiseEnForme Générale ${nomOnglet}: ${e}`);
      }
    }
  }
  Logger.log("[MiseEnForme Générale] Terminé.");
}

// --- Fonctions potentiellement manquantes ---

/**
 * Protège une feuille
 * @param {Sheet} sheet - La feuille à protéger
 * @param {string} description - Description de la protection
 * @param {boolean} warningOnly - Si true, avertissement seulement
 */
function protegerFeuille(sheet, description, warningOnly = false) {
  if (!sheet) return;
  try {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(p => p.remove());
    const protection = sheet.protect().setDescription(description);
    protection.removeEditors(protection.getEditors());
    if (warningOnly) {
      protection.setWarningOnly(true);
    } else {
       protection.addEditor(Session.getEffectiveUser());
    }
     Logger.log(`Protection (${warningOnly ? 'Avert.' : 'Complète'}) appliquée: ${sheet.getName()}`);
  } catch(e) { Logger.log(`Erreur protection ${sheet.getName()}: ${e}`);}
}

/**
 * Met à jour un paramètre de configuration
 * @param {string} param - Le paramètre à mettre à jour
 * @param {any} value - La nouvelle valeur
 */
function updateConfig(param, value) {
    try {
        const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CONFIG || "_CONFIG");
        if (!configSheet) return;
        const data = configSheet.getDataRange().getValues();
        for (let i=1; i<data.length; i++) {
            if (data[i][0] === param) {
                configSheet.getRange(i+1, 2).setValue(value);
                return;
            }
        }
         Logger.log(`updateConfig: Paramètre ${param} non trouvé pour mise à jour.`);
    } catch(e) { Logger.log(`Erreur updateConfig simple: ${e}`);}
}

/**
 * Ajoute la formule NOM_PRENOM et ID_ELEVE
 * @param {Sheet} sheet - La feuille cible
 * @param {number} initialRows - Nombre de lignes initiales
 */
function ajouterFormules(sheet, initialRows) {
  if (!sheet) {
    Logger.log("ERREUR: ajouterFormules appelée avec une feuille undefined");
    return;
  }
  
  try {
    Logger.log(`AjouterFormules: Traitement de ${sheet.getName()}`);
    initialRows = initialRows || 100;
    
    // Fonction vide pour l'instant mais qui permet la continuité
    
  } catch (e) {
    Logger.log(`ERREUR dans ajouterFormules: ${e}`);
  }
}