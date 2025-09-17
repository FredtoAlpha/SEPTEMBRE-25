/**
 * Script_Reservation.gs
 * Fonctions pour l'interface d'analyse et réservation des élèves
 */

// Bandeau maintenance : tous les utilitaires sont centralisés dans Utils.js
// Utilisez uniquement Utils.idx(header, name) et Utils.logAction(...)

/**
 * Ouvre l'interface HTML de réservation
 * Cette fonction est appelée par le bouton "Analyser Réserver" de la console
 */
function ouvrirInterfaceReservation() {
  const html = HtmlService.createHtmlOutputFromFile('ReservationUI')
    .setWidth(1000)
    .setHeight(700)
    .setTitle('Analyse & Réservation');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Analyse & Réservation');
}

/**
 * Récupère toutes les données nécessaires pour l'interface
 * @return {Object} Données formatées pour l'interface
 */
function getDonneesReservation() {
  try {
    const donnees = extraireDonneesEleves();
    
    return {
      eleves: donnees.eleves,
      criteres: {
        com: donnees.com,
        tra: donnees.tra,
        abs: donnees.abs,
        part: donnees.part
      },
      nbClasses: donnees.nbClasses,
      stats: {
        totalEleves: donnees.eleves.length,
        codesDISSO: compteCodes(donnees.eleves, "codeDISSO"),
        codesASSO: compteCodes(donnees.eleves, "codeASSO"),
        pourcentage: calculPourcentageCodes(donnees.eleves)
      },
      alertes: verifierCoherenceCodes(donnees)
    };
  } catch (e) {
    Logger.log("Erreur dans getDonneesReservation: " + e.toString());
    throw e;
  }
}

/**
 * Enregistre les codes DISSO et ASSO dans les feuilles sources
 * @param {Object} codes - Objet contenant les associations ID/codes
 * @return {Object} Résultat de l'opération
 */
function enregistrerCodesReservation(codes) {
  try {
    let totalUpdates = 0;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Parcourir toutes les feuilles
    const sheets = ss.getSheets();
    
    for (const sheet of sheets) {
      const nomFeuille = sheet.getName();
      
      // Exclure les feuilles spéciales
      if (nomFeuille.startsWith("_") || 
          nomFeuille.endsWith("Crit") || 
          nomFeuille === "RESERVATION") {
        continue;
      }
      
      const data = sheet.getDataRange().getValues();
      
      // Vérifier si l'onglet a au moins une ligne d'en-tête
      if (data.length <= 0) {
        continue;
      }
      
      // Trouver la colonne ID
      let idColIndex = -1;
      for (let col = 0; col < data[0].length; col++) {
        const header = data[0][col] ? data[0][col].toString().toUpperCase() : "";
        if (header === "ID") {
          idColIndex = col;
          break;
        }
      }
      
      // Si on n'a pas trouvé la colonne ID, passer à la feuille suivante
      if (idColIndex === -1) {
        continue;
      }
      
      // Colonne M (index 12) pour ASSO et N (index 13) pour DISSO
      const assoColIndex = 12;
      const dissoColIndex = 13;
      
      // S'assurer que les colonnes existent
      while (sheet.getMaxColumns() <= Math.max(assoColIndex, dissoColIndex)) {
        sheet.insertColumnAfter(sheet.getMaxColumns());
      }
      
      // Mettre à jour les en-têtes si nécessaire
      if (!data[0][assoColIndex] || data[0][assoColIndex] === "") {
        sheet.getRange(1, assoColIndex + 1).setValue("ASSO");
      }
      
      if (!data[0][dissoColIndex] || data[0][dissoColIndex] === "") {
        sheet.getRange(1, dissoColIndex + 1).setValue("DISSO");
      }
      
      // Mettre à jour les codes pour chaque ligne
      let sheetUpdated = false;
      
      for (let row = 1; row < data.length; row++) {
        const id = data[row][idColIndex];
        
        if (id && (codes.disso[id] !== undefined || codes.asso[id] !== undefined)) {
          // Mise à jour ASSO
          const newAsso = codes.asso[id] || "";
          const currentAsso = row < data.length && assoColIndex < data[row].length ? data[row][assoColIndex] : "";
          
          if (currentAsso !== newAsso) {
            sheet.getRange(row + 1, assoColIndex + 1).setValue(newAsso);
            sheetUpdated = true;
            totalUpdates++;
          }
          
          // Mise à jour DISSO
          const newDisso = codes.disso[id] || "";
          const currentDisso = row < data.length && dissoColIndex < data[row].length ? data[row][dissoColIndex] : "";
          
          if (currentDisso !== newDisso) {
            sheet.getRange(row + 1, dissoColIndex + 1).setValue(newDisso);
            sheetUpdated = true;
            totalUpdates++;
          }
        }
      }
      
      // Actualiser si des mises à jour ont été effectuées
      if (sheetUpdated) {
        SpreadsheetApp.flush();
      }
    }
    
    // Enregistrer dans le journal
    ajouterEntreeJournal(`Codes enregistrés: ${totalUpdates} modifications`);
    
    return { 
      success: true, 
      message: `Codes enregistrés avec succès! ${totalUpdates} mises à jour effectuées.`
    };
  } catch (e) {
    Logger.log("Erreur dans enregistrerCodesReservation: " + e.toString());
    ajouterEntreeJournal(`ERREUR lors de l'enregistrement des codes: ${e.toString()}`);
    return { success: false, message: "Erreur: " + e.toString() };
  }
}

/**
 *  Nouveau extraireDonneesEleves :
 *  – si des onglets se terminent par “Crit” existent, on les utilise (ancien comportement)
 *  – sinon on lit les feuilles d’origine déclarées dans _STRUCTURE (col. A : CLASSE_ORIGINE)
 *  Retourne { eleves, com, tra, part, abs, nbClasses }
 */
function extraireDonneesEleves() {

  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets     = ss.getSheets();
  const critSheets    = allSheets.filter(sh => /Crit$/i.test(sh.getName()));
  const eleves        = [];

  /** Ajoute les élèves d’une feuille “données” + codes d’une feuille “source”      */
  function pushFromSheets(sheetData, sheetSource, classe) {

    const data        = sheetData.getDataRange().getValues();
    const source      = sheetSource ? sheetSource.getDataRange().getValues() : null;

    // Cache ID → ASSO / DISSO
    const codesASSO   = {};
    const codesDISSO  = {};

    if (source && source.length > 1) {
      for (let r = 1; r < source.length; r++) {
        const id = source[r][0];    // Col A
        if (!id) continue;
        codesASSO[id]  = source[r][12] || "";   // Col M
        codesDISSO[id] = source[r][13] || "";   // Col N
      }
    }

    // Ligne d’en-tête ?
    const start = (data[0][0] + "").toUpperCase() === "ID" ? 1 : 0;

    for (let r = start; r < data.length; r++) {
      const row = data[r];
      const id  = row[0];
      if (!id) continue;

      eleves.push({
        id          : id,
        nom_prenom  : row[1],
        com         : parseFloat(row[2]) || 0,
        tra         : parseFloat(row[3]) || 0,
        part        : parseFloat(row[4]) || 0,
        abs         : parseFloat(row[5]) || 0,
        classeSource: classe,
        codeASSO    : codesASSO[id]  || "",
        codeDISSO   : codesDISSO[id] || ""
      });
    }
  }

  //------------------------------------------------------------------//
  // 1) CAS A : on a au moins un onglet Crit ➜ ancien comportement
  //------------------------------------------------------------------//
  if (critSheets.length) {

    critSheets.forEach(sh => {
      const classe            = sh.getName().replace(/Crit$/i, "");
      const feuilleSource     = ss.getSheetByName(classe); // peut être null
      pushFromSheets(sh, feuilleSource, classe);
    });

  //------------------------------------------------------------------//
  // 2) CAS B : aucun onglet Crit ➜ on lit _STRUCTURE → CLASSE_ORIGINE
  //------------------------------------------------------------------//
  } else {

    const struct = ss.getSheetByName("_STRUCTURE");
    if (!struct) throw new Error("Aucun onglet Crit et feuille _STRUCTURE manquante.");

    const sData  = struct.getDataRange().getValues();
    if (sData.length < 2) throw new Error("_STRUCTURE vide.");

    const idxOrig = sData[0].indexOf("CLASSE_ORIGINE");
    if (idxOrig === -1) throw new Error("Colonne CLASSE_ORIGINE absente dans _STRUCTURE.");

    const origines = new Set();

    for (let r = 1; r < sData.length; r++) {
      String(sData[r][idxOrig] || "")
        .split(',')
        .map(x => x.trim())
        .filter(Boolean)
        .forEach(o => origines.add(o));
    }

    origines.forEach(classe => {
      const sh = ss.getSheetByName(classe);
      if (!sh) return;       // on ignore celles sans feuille
      pushFromSheets(sh, sh, classe);  // données et codes sur la même feuille
    });
  }

  //------------------------------------------------------------------//
  // 3) Tri et retour
  //------------------------------------------------------------------//
  const tri = key => [...eleves].sort((a, b) => a[key] - b[key]);

  return {
    eleves   : eleves,
    com      : tri("com"),
    tra      : tri("tra"),
    part     : tri("part"),
    abs      : tri("abs"),
    nbClasses: determinerNombreClasses(eleves)   // votre fonction existante
  };
}

/**
 * Détermine le nombre de classes cibles
 * @param {Object[]} eleves - Liste des élèves
 * @return {number} Nombre de classes cibles
 */
function determinerNombreClasses(eleves) {
  // Par défaut, on suppose 5 classes
  let nbClasses = 5;
  
  // Si possible, déterminer le nombre de classes à partir des configurations existantes
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("_STRUCTURE");
    
    if (configSheet) {
      const configData = configSheet.getDataRange().getValues();
      
      // Compter le nombre de classes destination uniques
      const classesDestination = new Set();
      
      for (let i = 1; i < configData.length; i++) {
        if (configData[i][1]) { // Colonne B = CLASSE_DEST
          classesDestination.add(configData[i][1].toString());
        }
      }
      
      if (classesDestination.size > 0) {
        nbClasses = classesDestination.size;
      }
    }
  } catch (e) {
    Logger.log("Erreur lors de la détermination du nombre de classes: " + e.toString());
  }
  
  return nbClasses;
}

/**
 * Compte les codes uniques et les élèves associés
 * @param {Object[]} eleves - Liste des élèves
 * @param {string} typePropriete - Propriété à analyser ("codeDISSO" ou "codeASSO")
 * @return {Object} Statistiques sur les codes
 */
function compteCodes(eleves, typePropriete) {
  const codes = {};
  let totalEleves = 0;
  
  eleves.forEach(eleve => {
    const code = eleve[typePropriete];
    if (code) {
      if (!codes[code]) {
        codes[code] = 0;
      }
      codes[code]++;
      totalEleves++;
    }
  });
  
  return {
    nbCodes: Object.keys(codes).length,
    nbEleves: totalEleves,
    details: codes
  };
}

/**
 * Calcule le pourcentage d'élèves ayant des codes
 * @param {Object[]} eleves - Liste des élèves
 * @return {Object} Pourcentage calculé
 */
function calculPourcentageCodes(eleves) {
  const totalEleves = eleves.length;
  let elevesAvecCodes = 0;
  
  eleves.forEach(eleve => {
    if (eleve.codeDISSO || eleve.codeASSO) {
      elevesAvecCodes++;
    }
  });
  
  const pourcentage = totalEleves > 0 ? (elevesAvecCodes / totalEleves) * 100 : 0;
  
  return {
    pourcentage: pourcentage,
    statut: pourcentage <= 30 ? "OK" : pourcentage <= 35 ? "ATTENTION" : "CRITIQUE"
  };
}

/**
 * Vérifie la cohérence des codes DISSO et ASSO
 * @param {Object} donnees - Données des élèves et nombre de classes
 * @return {Object} Alertes détectées
 */
function verifierCoherenceCodes(donnees) {
  const alertesDISSO = [];
  const alertesASSO = [];
  
  // Vérification des codes DISSO (pas plus d'élèves que de classes)
  const codesDISSO = {};
  donnees.eleves.forEach(eleve => {
    if (eleve.codeDISSO) {
      if (!codesDISSO[eleve.codeDISSO]) {
        codesDISSO[eleve.codeDISSO] = [];
      }
      codesDISSO[eleve.codeDISSO].push(eleve);
    }
  });
  
  Object.keys(codesDISSO).forEach(code => {
    const eleves = codesDISSO[code];
    if (eleves.length > donnees.nbClasses) {
      alertesDISSO.push({
        code: code,
        nbEleves: eleves.length,
        nbClasses: donnees.nbClasses,
        message: `Le code DISSO "${code}" est attribué à ${eleves.length} élèves pour ${donnees.nbClasses} classes disponibles.`
      });
    }
  });
  
  // Vérification des codes ASSO (au moins 2 élèves par code)
  const codesASSO = {};
  donnees.eleves.forEach(eleve => {
    if (eleve.codeASSO) {
      if (!codesASSO[eleve.codeASSO]) {
        codesASSO[eleve.codeASSO] = [];
      }
      codesASSO[eleve.codeASSO].push(eleve);
    }
  });
  
  Object.keys(codesASSO).forEach(code => {
    const eleves = codesASSO[code];
    if (eleves.length === 1) {
      alertesASSO.push({
        code: code,
        nbEleves: eleves.length,
        eleve: eleves[0].nom_prenom,
        message: `Le code ASSO "${code}" n'est attribué qu'à un seul élève (${eleves[0].nom_prenom}).`
      });
    } else if (eleves.length > 5) {
      alertesASSO.push({
        code: code,
        nbEleves: eleves.length,
        message: `Le code ASSO "${code}" est attribué à ${eleves.length} élèves (groupe potentiellement trop grand).`
      });
    }
  });
  
  return {
    alertesDISSO: alertesDISSO,
    alertesASSO: alertesASSO,
    nbAlertes: alertesDISSO.length + alertesASSO.length
  };
}

/**
 * Recherche un élève par nom ou partie du nom
 * @param {string} terme - Terme de recherche
 * @return {Object[]} Élèves correspondants
 */
function rechercherEleve(terme) {
  try {
    const donnees = extraireDonneesEleves();
    const termeRecherche = terme.toLowerCase();
    
    const resultats = donnees.eleves.filter(eleve => 
      eleve.nom_prenom.toLowerCase().includes(termeRecherche)
    );
    
    return resultats;
  } catch (e) {
    Logger.log("Erreur dans rechercherEleve: " + e.toString());
    return [];
  }
}

/**
 * Ajoute une entrée dans le journal des actions
 * @param {string} message - Message à journaliser
 */
function ajouterEntreeJournal(message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName("_LOG_RESERVATION");
    
    // Créer la feuille de log si elle n'existe pas
    if (!logSheet) {
      logSheet = ss.insertSheet("_LOG_RESERVATION");
      logSheet.appendRow(["Date", "Heure", "Action"]);
      logSheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#E8EAF6");
    }
    
    // Ajouter l'entrée de journal
    const now = new Date();
    const date = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const heure = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");
    
    logSheet.appendRow([date, heure, message]);
    
    // Limiter le nombre d'entrées (garder les 1000 dernières)
    const maxRows = 1000;
    const currentRows = logSheet.getLastRow();
    if (currentRows > maxRows + 1) { // +1 pour l'en-tête
      logSheet.deleteRows(2, currentRows - maxRows - 1);
    }
  } catch (e) {
    Logger.log("Erreur lors de l'ajout d'une entrée au journal: " + e.toString());
  }
}
