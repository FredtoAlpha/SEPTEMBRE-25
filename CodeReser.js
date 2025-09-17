/**
 * ================================================================
 * SCRIPT D'ANALYSE DES CODES DE RÉSERVATION
 * ================================================================
 * Analyse les codes D et A dans les classes sources et affiche les statistiques
 */

/**
 * Fonction principale pour analyser les codes de réservation
 * Appelée depuis le bouton "CODES RESERVATION" de la console
 */
function analyserCodesReservation() {
  try {
    Logger.log("=== DÉBUT ANALYSE CODES RÉSERVATION ===");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // 1. LIRE LA STRUCTURE POUR TROUVER LES CLASSES SOURCES
    const classeSources = lireClassesSources();
    
    if (classeSources.length === 0) {
      const message = "❌ Aucune classe source trouvée dans l'onglet _STRUCTURE.\n\nVérifiez que :\n- L'onglet _STRUCTURE existe\n- La colonne A contient 'SOURCE'\n- La colonne B contient les noms de classes (ex: 4°1, 5°3)";
      ui.alert("Aucune classe source", message, ui.ButtonSet.OK);
      return message;
    }
    
    Logger.log(`✅ Classes sources trouvées: ${classeSources.join(', ')}`);
    
    // 2. ANALYSER LES CODES DANS CHAQUE CLASSE SOURCE
    const statsGlobales = analyserCodesDansClasses(classeSources);
    
    // 3. AFFICHER LES RÉSULTATS
    const resultat = afficherResultatsCodesReservation(statsGlobales, classeSources);
    
    Logger.log("=== FIN ANALYSE CODES RÉSERVATION ===");
    return resultat;
    
  } catch (e) {
    Logger.log(`❌ ERREUR analyserCodesReservation: ${e.message}\n${e.stack}`);
    const messageErreur = `Erreur lors de l'analyse des codes de réservation :\n\n${e.message}`;
    SpreadsheetApp.getUi().alert("Erreur", messageErreur, SpreadsheetApp.getUi().ButtonSet.OK);
    return messageErreur;
  }
}

/**
 * Lit l'onglet _STRUCTURE pour trouver les classes sources
 * @return {string[]} Tableau des noms de classes sources
 */
function lireClassesSources() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const structureSheet = ss.getSheetByName("_STRUCTURE");
  
  if (!structureSheet) {
    Logger.log("❌ Onglet _STRUCTURE introuvable");
    return [];
  }
  
  const data = structureSheet.getDataRange().getValues();
  const classes = [];
  
  // Chercher les lignes avec Type = "SOURCE"
  for (let i = 1; i < data.length; i++) { // Ignorer la ligne d'en-tête
    const type = String(data[i][0] || '').trim().toUpperCase();
    const nomClasse = String(data[i][1] || '').trim();
    
    if (type === 'SOURCE' && nomClasse) {
      // Vérifier le format de classe (ex: 4°1, 5°3)
      if (/^[3-6]°[1-9]\d*$/.test(nomClasse)) {
        classes.push(nomClasse);
        Logger.log(`✅ Classe source trouvée: ${nomClasse}`);
      } else {
        Logger.log(`⚠️ Format de classe invalide ignoré: "${nomClasse}"`);
      }
    }
  }
  
  return classes;
}

/**
 * Analyse les codes D et A dans toutes les classes sources
 * @param {string[]} classeSources - Liste des noms de classes sources
 * @return {object} Statistiques globales des codes
 */
function analyserCodesDansClasses(classeSources) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statsGlobales = {
    codesD: {},  // Ex: { "D1": 6, "D2": 3 }
    codesA: {},  // Ex: { "A1": 2, "A3": 1 }
    details: {}  // Détails par classe
  };
  
  classeSources.forEach(nomClasse => {
    const sheet = ss.getSheetByName(nomClasse);
    
    if (!sheet) {
      Logger.log(`⚠️ Onglet "${nomClasse}" introuvable`);
      return;
    }
    
    Logger.log(`🔍 Analyse de la classe: ${nomClasse}`);
    
    // Analyser cette classe spécifique
    const statsClasse = analyserCodesUneClasse(sheet, nomClasse);
    
    // Ajouter aux statistiques globales
    statsClasse.codesD.forEach(code => {
      statsGlobales.codesD[code] = (statsGlobales.codesD[code] || 0) + 1;
    });
    
    statsClasse.codesA.forEach(code => {
      statsGlobales.codesA[code] = (statsGlobales.codesA[code] || 0) + 1;
    });
    
    statsGlobales.details[nomClasse] = statsClasse;
  });
  
  return statsGlobales;
}

/**
 * Analyse les codes D et A dans une classe spécifique - CORRECTION VRAIE
 * @param {Sheet} sheet - L'onglet de la classe
 * @param {string} nomClasse - Nom de la classe
 * @return {object} Statistiques détaillées de la classe
 */
function analyserCodesUneClasse(sheet, nomClasse) {
  const statsClasse = {
    codesD: [],           // Liste brute pour le total global (GARDÉ)
    codesA: [],           // Liste brute pour le total global (GARDÉ)
    detailCodesD: {},     // AJOUTÉ: Comptage détaillé {D1: 3, D2: 2}
    detailCodesA: {},     // AJOUTÉ: Comptage détaillé {A1: 1, A3: 2}
    nbEleves: 0           // GARDÉ: Total élèves
  };
  
  try {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < 2 || lastCol < 14) {
      Logger.log(`⚠️ ${nomClasse}: Pas assez de données (${lastRow} lignes, ${lastCol} colonnes)`);
      return statsClasse;
    }
    
    // Lire les colonnes M (13) et N (14)
    const colonneM = sheet.getRange(2, 13, lastRow - 1, 1).getValues().flat(); // Codes A
    const colonneN = sheet.getRange(2, 14, lastRow - 1, 1).getValues().flat(); // Codes D
    
    // Traiter chaque ligne d'élève
    for (let i = 0; i < colonneM.length; i++) {
      const codeA = String(colonneM[i] || '').trim().toUpperCase();
      const codeD = String(colonneN[i] || '').trim().toUpperCase();
      
      // Vérifier si la ligne contient des données d'élève
      if (codeA || codeD) {
        statsClasse.nbEleves++;
        
        // Traiter codes A
        if (codeA && /^A\d+$/.test(codeA)) {
          statsClasse.codesA.push(codeA);                                    // GARDÉ pour total global
          statsClasse.detailCodesA[codeA] = (statsClasse.detailCodesA[codeA] || 0) + 1; // AJOUTÉ pour détail
        }
        
        // Traiter codes D
        if (codeD && /^D\d+$/.test(codeD)) {
          statsClasse.codesD.push(codeD);                                    // GARDÉ pour total global
          statsClasse.detailCodesD[codeD] = (statsClasse.detailCodesD[codeD] || 0) + 1; // AJOUTÉ pour détail
        }
      }
    }
    
    // Log amélioré avec détail
    const detailsD = Object.keys(statsClasse.detailCodesD).map(code => `${code}:${statsClasse.detailCodesD[code]}`).join(', ');
    const detailsA = Object.keys(statsClasse.detailCodesA).map(code => `${code}:${statsClasse.detailCodesA[code]}`).join(', ');
    Logger.log(`📊 ${nomClasse}: ${statsClasse.nbEleves} élèves, ${statsClasse.codesA.length} codes A, ${statsClasse.codesD.length} codes D`);
    if (detailsD) Logger.log(`   Détail D: ${detailsD}`);
    if (detailsA) Logger.log(`   Détail A: ${detailsA}`);
    
  } catch (e) {
    Logger.log(`❌ Erreur analyse ${nomClasse}: ${e.message}`);
  }
  
  return statsClasse;
}

/**
 * Affiche les résultats COMPLETS avec tout gardé + détail ajouté
 * @param {object} statsGlobales - Statistiques globales
 * @param {string[]} classeSources - Liste des classes analysées
 * @return {string} Message de résultat
 */
function afficherResultatsCodesReservation(statsGlobales, classeSources) {
  const ui = SpreadsheetApp.getUi();
  
  // Construire le message de résultats
  let message = `📊 ANALYSE DES CODES DE RÉSERVATION\n`;
  message += `════════════════════════════════════\n\n`;
  message += `Classes analysées: ${classeSources.length}\n`;
  message += `${classeSources.join(', ')}\n\n`;
  
  // CODES D (Dissociation) - GARDÉ
  message += `🔴 CODES D (Dissociation):\n`;
  const codesD = Object.keys(statsGlobales.codesD).sort();
  if (codesD.length > 0) {
    codesD.forEach(code => {
      message += `   ${code} = ${statsGlobales.codesD[code]} occurrence${statsGlobales.codesD[code] > 1 ? 's' : ''}\n`;
    });
    message += `   Total codes D: ${codesD.reduce((sum, code) => sum + statsGlobales.codesD[code], 0)}\n`;
  } else {
    message += `   Aucun code D trouvé\n`;
  }
  
  message += `\n`;
  
  // CODES A (Association) - GARDÉ
  message += `🟢 CODES A (Association):\n`;
  const codesA = Object.keys(statsGlobales.codesA).sort();
  if (codesA.length > 0) {
    codesA.forEach(code => {
      message += `   ${code} = ${statsGlobales.codesA[code]} occurrence${statsGlobales.codesA[code] > 1 ? 's' : ''}\n`;
    });
    message += `   Total codes A: ${codesA.reduce((sum, code) => sum + statsGlobales.codesA[code], 0)}\n`;
  } else {
    message += `   Aucun code A trouvé\n`;
  }
  
  // DÉTAILS PAR CLASSE - GARDÉ le format original + AJOUTÉ le détail
  message += `\n📋 DÉTAIL PAR CLASSE:\n`;
  classeSources.forEach(classe => {
    const detail = statsGlobales.details[classe];
    if (detail) {
      // GARDÉ: Ligne originale avec comptes totaux
      message += `   ${classe}: ${detail.nbEleves} élèves, ${detail.codesA.length} codes A, ${detail.codesD.length} codes D\n`;
      
      // AJOUTÉ: Détail des codes
      const codesD = Object.keys(detail.detailCodesD);
      const codesA = Object.keys(detail.detailCodesA);
      
      if (codesD.length > 0 || codesA.length > 0) {
        let detailLigne = `      → `;
        
        if (codesD.length > 0) {
          const detailsD = codesD.sort().map(code => `${detail.detailCodesD[code]} ${code}`).join(', ');
          detailLigne += `D: ${detailsD}`;
        }
        
        if (codesA.length > 0) {
          const detailsA = codesA.sort().map(code => `${detail.detailCodesA[code]} ${code}`).join(', ');
          detailLigne += (codesD.length > 0 ? ' | ' : '') + `A: ${detailsA}`;
        }
        
        message += `${detailLigne}\n`;
      }
    }
  });
  
  // Afficher dans une boîte de dialogue - GARDÉ
  ui.alert("Codes de Réservation", message, ui.ButtonSet.OK);
  
  // Logger aussi pour les logs - GARDÉ
  Logger.log(message);
  
  return message;
}

/**
 * Version console CORRIGÉE pour notification toast avec détail des codes
 */
function analyserCodesReservation() {
  try {
    Logger.log("=== ANALYSE CODES RÉSERVATION - CONSOLE ===");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Lire les classes sources
    const classeSources = lireClassesSources();
    
    if (classeSources.length === 0) {
      return "❌ Aucune classe source trouvée.\n\nVérifiez l'onglet _STRUCTURE :\n• Colonne A doit contenir 'SOURCE'\n• Colonne B doit contenir les noms de classes (ex: 4°1, 5°3)";
    }
    
    // 2. Analyser les codes
    const statsGlobales = analyserCodesDansClasses(classeSources);
    
    // 3. Formater pour notification TOAST (format compact)
    let message = `📊 CODES RÉSERVATION - ${classeSources.length} classe${classeSources.length > 1 ? 's' : ''} analysée${classeSources.length > 1 ? 's' : ''}\n`;
    message += `════════════════════════════════════\n\n`;
    
    // Codes D compacts
    const codesD = Object.keys(statsGlobales.codesD).sort();
    if (codesD.length > 0) {
      message += `🔴 CODES D: ${codesD.map(code => `${code}=${statsGlobales.codesD[code]}`).join(' • ')}\n`;
    } else {
      message += `🔴 CODES D: Aucun\n`;
    }
    
    // Codes A compacts
    const codesA = Object.keys(statsGlobales.codesA).sort();
    if (codesA.length > 0) {
      message += `🟢 CODES A: ${codesA.map(code => `${code}=${statsGlobales.codesA[code]}`).join(' • ')}\n`;
    } else {
      message += `🟢 CODES A: Aucun\n`;
    }
    
    // Totaux
    const totalD = codesD.reduce((sum, code) => sum + statsGlobales.codesD[code], 0);
    const totalA = codesA.reduce((sum, code) => sum + statsGlobales.codesA[code], 0);
    message += `\n📈 TOTAUX: ${totalD} codes D • ${totalA} codes A\n`;
    
    // DÉTAIL PAR CLASSE avec format précis comme demandé
    message += `\n📋 DÉTAIL PAR CLASSE:\n`;
    classeSources.forEach(classe => {
      const detail = statsGlobales.details[classe];
      if (detail) {
        let ligneClasse = `${classe}: `;
        
        // Construire le détail précis des codes
        const codesD = Object.keys(detail.detailCodesD);
        const codesA = Object.keys(detail.detailCodesA);
        
        const detailsParts = [];
        
        // Ajouter les codes D avec format "3 D1, 2 D2"
        if (codesD.length > 0) {
          const detailsD = codesD.sort().map(code => `${detail.detailCodesD[code]} ${code}`).join(', ');
          detailsParts.push(detailsD);
        }
        
        // Ajouter les codes A avec format "1 A3"
        if (codesA.length > 0) {
          const detailsA = codesA.sort().map(code => `${detail.detailCodesA[code]} ${code}`).join(', ');
          detailsParts.push(detailsA);
        }
        
        // Si aucun code, afficher "0 codes"
        if (detailsParts.length === 0) {
          detailsParts.push("0 codes");
        }
        
        ligneClasse += detailsParts.join(' | ');
        message += `${ligneClasse}\n`;
      }
    });
    
    Logger.log("Analyse codes réservation console terminée avec succès");
    return message;
    
  } catch (e) {
    Logger.log(`❌ ERREUR analyserCodesReservation: ${e.message}`);
    return `❌ Erreur lors de l'analyse des codes :\n\n${e.message}`;
  }
}