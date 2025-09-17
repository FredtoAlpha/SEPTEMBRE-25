/**
 * Vérifie l'intégrité des données consolidées
 * - Vérifie que chaque élève a un ID unique
 * - Vérifie que les champs obligatoires sont remplis
 * - Ignore les colonnes G (OPT), L, M et N lors de la vérification
 */
function verifierDonnees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Vérifier d'abord les onglets sources
  const sourceSheets = getSourceSheets();
  if (sourceSheets.length === 0) {
    ui.alert("Aucun onglet source trouvé. Veuillez vérifier votre structure.");
    return "Aucun onglet source introuvable";
  }
  
  // Liste des problèmes pour tous les onglets
  let problemesGlobaux = [];
  let totalEleves = 0;
  
  // Vérifier chaque onglet source
  for (const sheet of sourceSheets) {
    const sheetName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      // Onglet vide ou juste l'en-tête, passer au suivant
      continue;
    }
    
    const headerRow = data[0];
    
    // Colonnes à vérifier (A, B, C, D, E, F, H, I, J, K et O)
    // A=ID_ELEVE, B=NOM, C=PRENOM, D=NOM_PRENOM, E=SEXE, F=LV2, H=COM, I=TRA, J=PART, K=ABS, O=?
    const requiredColumns = ["ID_ELEVE", "NOM", "PRENOM", "NOM_PRENOM", "SEXE", "LV2"];
    const additionalColumns = ["COM", "TRA", "PART", "ABS"];
    
    const indexes = {};
    requiredColumns.forEach(col => {
      indexes[col] = headerRow.indexOf(col);
    });
    
    additionalColumns.forEach(col => {
      indexes[col] = headerRow.indexOf(col);
    });
    
    // Vérifier si les colonnes requises existent
    const missingColumns = [];
    requiredColumns.forEach(col => {
      if (indexes[col] === -1) {
        missingColumns.push(col);
      }
    });
    
    if (missingColumns.length > 0) {
      ui.alert(`Colonnes manquantes dans ${sheetName}: ${missingColumns.join(", ")}`);
      return `Colonnes manquantes dans ${sheetName}: ${missingColumns.join(", ")}`;
    }
    
    // Vérifier les données (ignorer l'en-tête)
    const problemes = [];
    const idsUtilises = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowNum = i + 1;
      
      // Ignorer les lignes vides
      if (!row[indexes.NOM] && !row[indexes.PRENOM]) {
        continue;
      }
      
      totalEleves++;
      
      // Vérifier ID
      if (!row[indexes.ID_ELEVE]) {
        problemes.push(`Ligne ${rowNum}: ID manquant pour ${row[indexes.NOM]} ${row[indexes.PRENOM]}`);
      } else if (idsUtilises[row[indexes.ID_ELEVE]]) {
        problemes.push(`Ligne ${rowNum}: ID en double "${row[indexes.ID_ELEVE]}" (déjà utilisé ligne ${idsUtilises[row[indexes.ID_ELEVE]]})`);
      } else {
        idsUtilises[row[indexes.ID_ELEVE]] = rowNum;
      }
      
      // Vérifier les champs obligatoires (NOM, PRENOM, SEXE, LV2)
      for (const col of ["NOM", "PRENOM", "SEXE", "LV2"]) {
        if (indexes[col] !== -1 && !row[indexes[col]]) {
          problemes.push(`Ligne ${rowNum}: "${col}" manquant pour ${row[indexes.NOM] || ""} ${row[indexes.PRENOM] || ""}`);
        }
      }
      
      // Vérifier que NOM_PRENOM est correctement formé (si présent)
      if (indexes.NOM_PRENOM !== -1 && row[indexes.NOM] && row[indexes.PRENOM]) {
        const expectedNomPrenom = `${row[indexes.NOM]} ${row[indexes.PRENOM]}`;
        if (row[indexes.NOM_PRENOM] !== expectedNomPrenom) {
          problemes.push(`Ligne ${rowNum}: NOM_PRENOM incorrect "${row[indexes.NOM_PRENOM]}" (devrait être "${expectedNomPrenom}")`);
        }
      }
      
      // Vérifier les critères (COM, TRA, PART, ABS) s'ils existent
      additionalColumns.forEach(col => {
        if (indexes[col] !== -1) {
          const valeur = row[indexes[col]];
          if (valeur === "") {
            problemes.push(`Ligne ${rowNum}: "${col}" manquant pour ${row[indexes.NOM]} ${row[indexes.PRENOM]}`);
          } else if (typeof valeur === 'number' && (valeur < 1 || valeur > 4)) {
            problemes.push(`Ligne ${rowNum}: "${col}" invalide (${valeur}) pour ${row[indexes.NOM]} ${row[indexes.PRENOM]}`);
          }
        }
      });
      
      // Nous n'effectuons pas de vérification sur la colonne G (OPT), L, M et N
    }
    
    // Ajouter les problèmes de cet onglet à la liste globale
    if (problemes.length > 0) {
      problemesGlobaux.push(`Onglet ${sheetName}:`);
      problemesGlobaux = problemesGlobaux.concat(problemes);
      problemesGlobaux.push(""); // Ligne vide entre les onglets
    }
  }
  
  // Vérifier également l'onglet CONSOLIDATION s'il existe
  const consolidationSheet = ss.getSheetByName("CONSOLIDATION");
  if (consolidationSheet) {
    const data = consolidationSheet.getDataRange().getValues();
    
    if (data.length > 1) {
      const headerRow = data[0];
      
      // Vérifier uniquement ID_ELEVE, NOM, PRENOM, SEXE, LV2 dans CONSOLIDATION
      const requiredColumns = ["ID_ELEVE", "NOM", "PRENOM", "SEXE", "LV2"];
      
      const indexes = {};
      requiredColumns.forEach(col => {
        indexes[col] = headerRow.indexOf(col);
      });
      
      // Vérifier si les colonnes requises existent
      const missingColumns = [];
      requiredColumns.forEach(col => {
        if (indexes[col] === -1) {
          missingColumns.push(col);
        }
      });
      
      if (missingColumns.length > 0) {
        ui.alert(`Colonnes manquantes dans CONSOLIDATION: ${missingColumns.join(", ")}`);
        return `Colonnes manquantes dans CONSOLIDATION: ${missingColumns.join(", ")}`;
      }
      
      // Vérifier les données (ignorer l'en-tête)
      const problemes = [];
      const idsUtilises = {};
      const idsSuffixesTraites = new Set(); // Pour traquer les IDs avec suffixes déjà traités
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowNum = i + 1;
        
        // Ignorer les lignes vides
        if (!row[indexes.NOM] && !row[indexes.PRENOM]) {
          continue;
        }
        
        // Vérifier ID
        if (!row[indexes.ID_ELEVE]) {
          problemes.push(`Ligne ${rowNum}: ID manquant pour ${row[indexes.NOM]} ${row[indexes.PRENOM]}`);
        } else {
          const id = String(row[indexes.ID_ELEVE]);
          
          // Ignorer les vérifications de doublons pour les IDs avec suffixes (_1, _2, etc.)
          if (id.includes('_')) {
            // Extraire la partie base de l'ID (avant le _)
            const idBase = id.split('_')[0];
            
            // Si c'est la première fois qu'on voit cet ID avec suffixe, on l'accepte
            if (!idsSuffixesTraites.has(id)) {
              idsSuffixesTraites.add(id);
            } else {
              // Sinon, c'est un doublon d'un ID déjà avec suffixe
              problemes.push(`Ligne ${rowNum}: ID en double "${id}" (déjà utilisé ligne ${idsUtilises[id]})`);
            }
          } else if (idsUtilises[id]) {
            // ID en double (sans suffixe)
            problemes.push(`Ligne ${rowNum}: ID en double "${id}" (déjà utilisé ligne ${idsUtilises[id]})`);
          }
          
          // Enregistrer l'ID utilisé
          idsUtilises[id] = rowNum;
        }
        
        // Vérifier les champs obligatoires (NOM, PRENOM, SEXE, LV2)
        for (const col of ["NOM", "PRENOM", "SEXE", "LV2"]) {
          if (!row[indexes[col]]) {
            problemes.push(`Ligne ${rowNum}: "${col}" manquant pour ${row[indexes.NOM] || ""} ${row[indexes.PRENOM] || ""}`);
          }
        }
      }
      
      // Ajouter les problèmes de CONSOLIDATION à la liste globale
      if (problemes.length > 0) {
        problemesGlobaux.push("Onglet CONSOLIDATION:");
        problemesGlobaux = problemesGlobaux.concat(problemes);
      }
    }
  }
  
  // Afficher le résultat
  if (problemesGlobaux.length === 0) {
    ui.alert(
      "Vérification terminée",
      `Aucun problème détecté.\nTotal d'élèves: ${totalEleves}`,
      ui.ButtonSet.OK
    );
    return "Aucun problème détecté";
  } else {
    // Limiter le nombre de problèmes affichés
    const maxProblemes = 20;
    const rapport = problemesGlobaux.slice(0, maxProblemes).join("\n") + 
                   (problemesGlobaux.length > maxProblemes ? `\n\n... et ${problemesGlobaux.length - maxProblemes} autres problèmes` : "");
    
    ui.alert(
      "Problèmes détectés",
      `${problemesGlobaux.length} problème(s) détecté(s):\n\n${rapport}\n\nVeuillez corriger ces problèmes avant de continuer.`,
      ui.ButtonSet.OK
    );
    return `${problemesGlobaux.length} problème(s) détecté(s)`;
  }
}
/**
 * Consolide les données des onglets sources vers l'onglet CONSOLIDATION
 */
function consoliderDonnees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Trouver les onglets sources
  const sourceSheets = getSourceSheets();
  if (sourceSheets.length === 0) {
    ui.alert("Aucun onglet source trouvé. Veuillez vérifier votre structure.");
    return "Aucun onglet source trouvé";
  }

  // Récupérer la liste des options valides depuis _CONFIG ou _STRUCTURE
  let optionsValides = [];
  try {
    // D'abord essayer _CONFIG
    const configSheet = ss.getSheetByName("_CONFIG");
    if (configSheet) {
      const data = configSheet.getDataRange().getValues();
      for (const row of data) {
        if (row[0] === "OPT" && row[1]) {
          optionsValides = String(row[1]).split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
          Logger.log(`Options récupérées depuis _CONFIG: ${optionsValides.join(',')}`);
          break;
        }
      }
    }
    
    // Si rien trouvé, essayer _STRUCTURE
    if (optionsValides.length === 0) {
      const structureSheet = ss.getSheetByName("_STRUCTURE");
      if (structureSheet) {
        const data = structureSheet.getDataRange().getValues();
        const optCol = data[0].indexOf("OPTIONS");
        if (optCol !== -1) {
          const optValues = data.slice(1)
                          .map(row => row[optCol])
                          .filter(val => val && typeof val === 'string')
                          .map(val => val.includes("=") ? val.split("=")[0].trim() : val.trim())
                          .filter(val => val);
          optionsValides = [...new Set(optValues)];
          Logger.log(`Options récupérées depuis _STRUCTURE: ${optionsValides.join(',')}`);
        }
      }
    }
    
    // Si toujours rien, utiliser les valeurs par défaut
    if (optionsValides.length === 0) {
      optionsValides = ["CHAV", "LATIN"];
      Logger.log("Utilisation des options par défaut: CHAV, LATIN");
    }
  } catch (e) {
    Logger.log(`Erreur récupération options: ${e.message}`);
    optionsValides = ["CHAV", "LATIN"]; // Valeurs par défaut en cas d'erreur
  }
  
  // Récupérer ou créer l'onglet CONSOLIDATION
  let consolidationSheet = ss.getSheetByName("CONSOLIDATION");
  if (!consolidationSheet) {
    consolidationSheet = ss.insertSheet("CONSOLIDATION");
    creerEnteteConsolidation(consolidationSheet);
  } else {
    // Effacer les données existantes (sauf l'en-tête)
    const lastRow = Math.max(consolidationSheet.getLastRow(), 2);
    const lastCol = consolidationSheet.getLastColumn();
    if (lastRow > 1) {
      consolidationSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    }
  }

  // Lire les en-têtes pour déterminer les indices des colonnes
  const headers = consolidationSheet.getRange(1, 1, 1, consolidationSheet.getLastColumn()).getValues()[0];
  const idIndex = headers.indexOf("ID_ELEVE");
  const sourceIndex = headers.indexOf("SOURCE");
  const optIndex = headers.indexOf("OPT"); // Récupérer l'index de la colonne OPT
  
  if (idIndex === -1 || sourceIndex === -1) {
    ui.alert("Structure de l'onglet CONSOLIDATION incorrecte. Veuillez réinitialiser l'onglet.");
    return "Structure incorrecte";
  }

  // Collecter d'abord toutes les données
  const toutesLesDonnees = [];
  const idsUtilises = new Set();
  
  for (const sheet of sourceSheets) {
    const sheetName = sheet.getName();
    const lastRowSource = Math.max(sheet.getLastRow(), 1);
    if (lastRowSource <= 1) continue; // Onglet vide
    
    // Lire toutes les données de la source
    const sourceData = sheet.getRange(2, 1, lastRowSource - 1, headers.length).getValues();
    
    // Filtrer les lignes vides (élèves sans nom ou prénom)
    const filteredData = sourceData.filter(row => row[1] && row[2]);
    
    // Ajouter la source et générer NOM_PRENOM si manquant
    filteredData.forEach(row => {
      // Si pas d'ID, en générer un
      if (!row[idIndex]) {
        row[idIndex] = `${sheetName}${(toutesLesDonnees.length + 1).toString().padStart(3, '0')}`;
      }
      
      // Assigner la source
      row[sourceIndex] = sheetName;
      
      // Générer NOM_PRENOM si manquant
      const nomIndex = headers.indexOf("NOM");
      const prenomIndex = headers.indexOf("PRENOM");
      const nomPrenomIndex = headers.indexOf("NOM_PRENOM");
      if (nomIndex !== -1 && prenomIndex !== -1 && nomPrenomIndex !== -1) {
        if (!row[nomPrenomIndex] && row[nomIndex] && row[prenomIndex]) {
          row[nomPrenomIndex] = `${row[nomIndex]} ${row[prenomIndex]}`;
        }
      }
      
      // IMPORTANT: Vérifier et nettoyer la valeur OPT si elle n'est pas valide
      if (optIndex !== -1 && row[optIndex] && !optionsValides.includes(row[optIndex])) {
        Logger.log(`Valeur OPT non valide trouvée: "${row[optIndex]}" - remplacée par ""`);
        row[optIndex] = ""; // Remplacer par une valeur vide si non valide
      }
      
      // Vérifier que l'ID est unique en ajoutant un suffixe si nécessaire
      let idOriginal = row[idIndex];
      let compteur = 1;
      while (idsUtilises.has(row[idIndex])) {
        row[idIndex] = `${idOriginal}_${compteur}`;
        compteur++;
      }
      
      idsUtilises.add(row[idIndex]);
      toutesLesDonnees.push(row);
    });
  }

  // Écrire toutes les données dans CONSOLIDATION
  if (toutesLesDonnees.length > 0) {
    consolidationSheet.getRange(2, 1, toutesLesDonnees.length, headers.length).setValues(toutesLesDonnees);
  }

  // Formater et trier
  if (toutesLesDonnees.length > 0) {
    // Créer un filtre
    consolidationSheet.getRange(1, 1, toutesLesDonnees.length + 1, headers.length).createFilter();
    
    // Trier par NOM, PRENOM
    const nomIndex = headers.indexOf("NOM") + 1; // +1 car getRange est 1-indexé
    const prenomIndex = headers.indexOf("PRENOM") + 1;
    if (nomIndex > 0 && prenomIndex > 0) {
      consolidationSheet.getRange(2, 1, toutesLesDonnees.length, headers.length)
      .sort([{column: nomIndex, ascending: true}, {column: prenomIndex, ascending: true}]);
    }
  }

  // Mettre en forme pour faciliter la lecture
  consolidationSheet.setFrozenRows(1);

  // Pour les listes déroulantes
  try {
    if (typeof ajouterListesDeroulantes === 'function') {
      ajouterListesDeroulantes();
    }
  } catch (e) {
    console.log("Fonction ajouterListesDeroulantes non disponible");
  }

  return `Consolidation terminée : ${toutesLesDonnees.length} élèves consolidés`;
}
