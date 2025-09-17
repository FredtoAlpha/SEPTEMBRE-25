// =========================================================
//         FICHIER : Phase5.V12.gs - PARTIE 1/5
//         Phase 5 (Finalisation) - Style "Maquette"
//         Version ULTRA-OPTIMISÉE avec colonnes utiles uniquement
// =========================================================

// --- Fonctions utilitaires LOCALES (au cas où Utils.gs ne serait pas modifiable ou complet) ---
function _phase5_ensureGlobalUtils() {
    const localLog = Logger.log; // Utiliser Logger.log de base pour cette fonction de vérification

    if (typeof logAction !== 'function') {
        localLog("WARN: Fonction globale 'logAction' non trouvée. Utilisation d'un fallback local pour Phase5.");
        // @ts-ignore // Supprime l'avertissement si 'logAction' est déjà défini globalement
        logAction = function(message) { Logger.log("(Phase5 Local Log) " + message); };
    }
    if (typeof normalizeHeader !== 'function') {
        localLog("WARN: Fonction globale 'normalizeHeader' non trouvée. Utilisation d'un fallback local pour Phase5.");
        // @ts-ignore
        normalizeHeader = function(s) { return String(s||"").replace(/[\s\u00A0\u200B-\u200D]/g, "").toUpperCase(); };
    }
    if (typeof getHeaders !== 'function') {
        localLog("WARN: Fonction globale 'getHeaders' non trouvée. Utilisation d'un fallback local pour Phase5.");
        // @ts-ignore
        getHeaders = function(sheet, doNormalize = true) {
            if (!sheet || sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) return [];
            try {
                const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
                // @ts-ignore
                return doNormalize ? headers.map(h => normalizeHeader(String(h || ""))) : headers.map(h => String(h || ""));
            } catch (e) { return []; }
        };
    }
    if (typeof getTestSheets !== 'function') {
        localLog("WARN: Fonction globale 'getTestSheets' non trouvée. Utilisation d'un fallback local pour Phase5.");
        // @ts-ignore
        getTestSheets = function(CFG_param) { // CFG_param est la config chargée
            const currentConfig = CFG_param || getConfig(); // S'assurer d'avoir une config
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sheets = ss.getSheets();
            const suffixTest = (currentConfig || {}).TEST_SUFFIX || "TEST";
            const reEscapeFn = typeof reEscape === 'function' ? reEscape : function(s) { return String(s).replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'); };
            const suffixTestRegex = new RegExp(reEscapeFn(suffixTest) + '$', 'i');
            return sheets.filter(sheet => suffixTestRegex.test(sheet.getName()));
        };
    }
    if (typeof ecartType !== 'function') {
        localLog("WARN: Fonction globale 'ecartType' non trouvée. Utilisation d'un fallback local pour Phase5.");
        // @ts-ignore
        ecartType = function(values) {
            if (!values || values.length <= 1) return 0;
            const n = values.length; const mean = values.reduce((a, b) => a + b, 0) / n;
            if (n === 0) return 0;
            const variance = values.map(x => Math.pow(x - mean, 2)).reduce((a, b) => a + b, 0) / n;
            return Math.sqrt(variance);
        };
    }
    if (typeof getOrCreateSheet !== 'function') {
        localLog("WARN: Fonction globale 'getOrCreateSheet' non trouvée. Utilisation d'un fallback local pour Phase5.");
        // @ts-ignore
        getOrCreateSheet = function(name) {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            return ss.getSheetByName(name) || ss.insertSheet(name);
        };
    }
    if (typeof cleanUnusedRowsAndColumns !== 'function') {
        localLog("WARN: Fonction globale 'cleanUnusedRowsAndColumns' non trouvée. Utilisation d'un fallback local pour Phase5.");
        // @ts-ignore
        cleanUnusedRowsAndColumns = function(sheet, lastUsedRow, lastUsedCol, bufferRows = 3, bufferCols = 1) {
            try { 
                // Version simplifiée pour le fallback
                if (!sheet) return;
                const maxRows = sheet.getMaxRows();
                const maxCols = sheet.getMaxColumns();
                
                if (maxRows > lastUsedRow + bufferRows) {
                    sheet.deleteRows(lastUsedRow + bufferRows + 1, maxRows - (lastUsedRow + bufferRows));
                }
                
                if (maxCols > lastUsedCol + bufferCols) {
                    sheet.deleteColumns(lastUsedCol + bufferCols + 1, maxCols - (lastUsedCol + bufferCols));
                }
            } catch (e) { Logger.log(`(Local Clean) WARN: ${e.message}`); }
        };
    }
    if (typeof idx !== 'function') {
        localLog("WARN: Fonction globale 'idx' non trouvée. Utilisation d'un fallback local pour _phase5_getColIndex.");
        // La fonction _phase5_getColIndex utilisera son propre fallback interne si idx global n'est pas trouvé.
    }
    if (typeof reEscape !== 'function') {
        localLog("WARN: Fonction globale 'reEscape' non trouvée. Utilisation d'un fallback local pour Phase5.");
        // @ts-ignore
        reEscape = function(s) { return String(s).replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'); };
    }
}
_phase5_ensureGlobalUtils(); // Appel pour définir les fallbacks si nécessaire au chargement du script

/** Point d'entrée pour les actions de la sidebar de finalisation */
function Phase5_ProcessAction(request) {
  Logger.log(">>> Phase5_ProcessAction: DÉBUT DE LA FONCTION <<<");
  Logger.log("Phase5_ProcessAction: Request reçu: " + JSON.stringify(request));

  let CFG;
  try {
    CFG = getConfig(); // Doit être la getConfig() corrigée de Config.gs
    Logger.log("Phase5_ProcessAction: getConfig() terminé. CFG.NIVEAU (test): " + (CFG ? CFG.NIVEAU : "CFG non défini"));
  } catch (e) {
    Logger.log("!!! ERREUR CRITIQUE dans Phase5_ProcessAction lors de l'appel à getConfig(): " + e.message + " Stack: " + e.stack);
    SpreadsheetApp.getUi().alert("Erreur Fatale", "Impossible de charger la configuration : " + e.message);
    return { success: false, message: "Erreur chargement configuration: " + e.message };
  }

  if (!CFG) { Logger.log("!!! ERREUR Phase5_ProcessAction: CFG est NULL après getConfig()"); return { success: false, message: "Erreur fatale: CFG indéfini." };}
  if (!CFG.ERROR_CODES) { Logger.log("!!! ERREUR Phase5_ProcessAction: CFG.ERROR_CODES est UNDEFINED"); return { success: false, message: "Erreur fatale: CFG.ERROR_CODES indéfini." }; }

  const action = request.action;
  const classes = request.classes || [];
  let result;
  
  logAction(`Phase 5 Action reçue: ${action} pour classes: ${classes.join(', ')}`);
  Logger.log(`Phase5_ProcessAction: Action = ${action}, Classes = ${classes.join(', ')}`);

  switch (action) {
    case "finaliser":
      Logger.log("Phase5_ProcessAction: Appel de Phase5_FinaliserClasses...");
      try {
        result = Phase5_FinaliserClasses(classes, request.hideTmp || false, CFG);
        Logger.log("Phase5_ProcessAction: Retour de Phase5_FinaliserClasses: " + JSON.stringify(result));
      } catch (e) {
        Logger.log("!!! ERREUR LORS DE L'APPEL à Phase5_FinaliserClasses: " + e.message + " Stack: " + e.stack);
        result = { success: false, message: "Erreur appel Phase5_FinaliserClasses: " + e.message, errorCode: CFG.ERROR_CODES.UNCAUGHT_EXCEPTION || "ERR_UNCAUGHT_FALLBACK" };
      }
      return result;
    case "genererPDF":
      Logger.log("Phase5_ProcessAction: Appel de Phase5_GenererPDF...");
      return Phase5_GenererPDF(classes, request.avecColoration || false, CFG);
    case "exporterExcel":
      Logger.log("Phase5_ProcessAction: Appel de Phase5_ExporterExcel...");
      return Phase5_ExporterExcel(classes, CFG);
    case "comparerClasses":
       logAction("Action 'comparerClasses' non implémentée séparément.");
       return {success: false, message: "Comparaison intégrée au Bilan DEF."};
    case "reinitialiser":
      Logger.log("Phase5_ProcessAction: Appel de Phase5_ReinitialiserTout...");
      return Phase5_ReinitialiserTout(CFG);
    case "toggleVisibilite":
       Logger.log("Phase5_ProcessAction: Appel de _phase5_gererVisibiliteOnglets...");
       const ss = SpreadsheetApp.getActiveSpreadsheet();
       const ongletsDef = classes.map(c => (c + (CFG.DEF_SUFFIX || "DEF")));
       _phase5_gererVisibiliteOnglets(ss, ongletsDef, !request.masquer, CFG);
       return { success: true, message: request.masquer ? "Onglets non finalisés masqués." : "Tous les onglets sont visibles." };
    default:
      logAction(`Action Phase 5 non reconnue: ${action}`);
      Logger.log(`Phase5_ProcessAction: Action non reconnue - ${action}`);
      return { success: false, errorCode: CFG.ERROR_CODES.ACTION_NOT_RECOGNIZED || "ERR_ACTION_NON_RECONNUE_FALLBACK", message: `Action non reconnue: ${action}` };
  }
}
// =========================================================
//         FICHIER : Phase5.V12.gs - PARTIE 2/5
//         Phase 5 (Finalisation) - Style "Maquette"
//         Version ULTRA-OPTIMISÉE avec colonnes utiles uniquement
// =========================================================

/** Récupère la liste des classes de base. */
function Phase5_GetClassesDisponibles() {
  const CFG = getConfig(); // Utilise la getConfig() globale
  logAction("Phase5_GetClassesDisponibles appelée...");
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const classesBase = new Set();
    const classesAvecDef = new Set();
    const suffixTest = CFG.TEST_SUFFIX || "TEST";
    const suffixDef = CFG.DEF_SUFFIX || "DEF";
    const reEscapeFn = typeof reEscape === 'function' ? reEscape : function(s) { return String(s).replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'); };

    const suffixTestRegex = new RegExp(reEscapeFn(suffixTest) + '$', 'i');
    const suffixDefRegex = new RegExp(reEscapeFn(suffixDef) + '$', 'i');

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const testMatch = sheetName.match(suffixTestRegex);
      const defMatch = sheetName.match(suffixDefRegex);
      if (testMatch) {
          const baseName = sheetName.replace(suffixTestRegex, '');
          if (!baseName.toUpperCase().startsWith("BILAN") && !baseName.startsWith("_")) { // Exclure bilans etc.
            classesBase.add(baseName);
            if (ss.getSheetByName(baseName + suffixDef)) classesAvecDef.add(baseName);
          }
      } else if (defMatch) {
          const baseName = sheetName.replace(suffixDefRegex, '');
           if (!baseName.toUpperCase().startsWith("BILAN") && !baseName.startsWith("_")) { // Exclure bilans etc.
             classesAvecDef.add(baseName);
           }
      }
    });
    const classesBaseArray = Array.from(classesBase).sort((a, b) => {
      const levelA = a.match(/(\d+)°/)?.[1] || a; const levelB = b.match(/(\d+)°/)?.[1] || b;
      const comp = (Number(levelB) || 0) - (Number(levelA) || 0);
      return comp === 0 ? a.localeCompare(b) : comp;
    });
    logAction(`Classes disponibles: ${classesBaseArray.length}, dont ${classesAvecDef.size} avec DEF.`);
    return { success: true, classes: classesBaseArray, classesAvecDef: Array.from(classesAvecDef) };
  } catch (e) {
    logAction(`Erreur Phase5_GetClassesDisponibles: ${e.message}`);
    return { success: false, message: `Erreur: ${e.message}`, classes: [], classesAvecDef: [] };
  }
}

/** Fonction ULTRA-OPTIMISÉE de finalisation - SEULEMENT COLONNES UTILES */
function Phase5_FinaliserClasses(classesBaseSelectionnees, masquerOngletsTemp, CFG) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const S_DEF = CFG.DEF_SUFFIX || "DEF";
  const out = { classes: [], erreurs: [] };
  let lock;

  logAction(`ULTRA-OPTIMISÉ: Début Phase5_FinaliserClasses pour: ${classesBaseSelectionnees.join(', ')}`);

  try {
    lock = LockService.getDocumentLock();
    if (!lock.tryLock(CFG.LOCK_TIMEOUT_FINALISER || 30000)) {
      return { success: false, errorCode: (CFG.ERROR_CODES || {}).LOCK_TIMEOUT || "ERR_LOCK_TIMEOUT", message: "Verrouillage impossible." };
    }

    // 1. SEULEMENT LES COLONNES UTILES QUE VOUS AVEZ SPÉCIFIÉES
    const COLONNES_A_CONSERVER = [
      { nom: "NOM & PRENOM", srcCol: 4 }, // D
      { nom: "SEXE", srcCol: 5 },         // E
      { nom: "LV2", srcCol: 6 },          // F
      { nom: "OPT", srcCol: 7 },          // G
      { nom: "COM", srcCol: 8 },          // H
      { nom: "TRA", srcCol: 9 },          // I
      { nom: "PART", srcCol: 10 },        // J
      { nom: "ABS", srcCol: 11 },         // K
      { nom: "INDICATEUR", srcCol: 12 },  // L
      { nom: "ASSO", srcCol: 13 },        // M
      { nom: "DISSO", srcCol: 14 },       // N
      { nom: "SOURCE", srcCol: -1 }       // Colonne virtuelle pour l'onglet d'origine
    ];

    // 2. Récupérer tous les onglets TEST
    const testSheets = getTestSheets();
    if (!testSheets || !testSheets.length) {
      logAction("ULTRA-OPTIMISÉ: Aucun onglet TEST trouvé.");
      return { success: false, errorCode: (CFG.ERROR_CODES || {}).NO_TEST_SHEETS || "ERR_NO_TEST_SHEETS", message: "Aucun onglet TEST trouvé." };
    }
    
    // 3. Pour chaque classe à finaliser
    for (const baseClass of classesBaseSelectionnees) {
      const defName = `${baseClass}${S_DEF}`;
      logAction(`ULTRA-OPTIMISÉ: Traitement de ${baseClass} -> ${defName}`);
      
      try {
        // 4. Créer l'onglet DEF
        let targetSheet = ss.getSheetByName(defName);
        if (targetSheet) {
          ss.deleteSheet(targetSheet);
        }
        targetSheet = ss.insertSheet(defName);
        
        // 5. Créer les en-têtes (UNIQUEMENT colonnes utiles)
        const headers = COLONNES_A_CONSERVER.map(c => c.nom);
        targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        
        // 6. Collecter les élèves de cette classe
        let elevesTrouves = [];
        
        for (const testSheet of testSheets) {
          const testSheetName = testSheet.getName();
          const data = testSheet.getDataRange().getValues();
          if (data.length <= 1) continue;
          
          // Colonne R = 18 (1-based) = 17 (0-based)
          const R_COLUMN_INDEX = 17;
          
          // Rechercher les élèves de cette classe
          for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row.length <= R_COLUMN_INDEX) continue;
            
            const classeValeur = String(row[R_COLUMN_INDEX] || "").trim();
            
            // Vérifier si l'élève est dans cette classe
            if (classeValeur === defName || classeValeur === baseClass) {
              // Extraire uniquement les colonnes utiles
              const eleveData = COLONNES_A_CONSERVER.map(col => {
                if (col.nom === "SOURCE") return testSheetName;
                return col.srcCol < row.length ? row[col.srcCol - 1] : "";
              });
              
              elevesTrouves.push(eleveData);
            }
          }
        }
// =========================================================
//         FICHIER : Phase5.V12.gs - PARTIE 3/5
//         Phase 5 (Finalisation) - Style "Maquette"
//         Version ULTRA-OPTIMISÉE avec colonnes utiles uniquement
// =========================================================

        // 7. Écrire les données dans l'onglet DEF
        if (elevesTrouves.length > 0) {
          targetSheet.getRange(2, 1, elevesTrouves.length, headers.length).setValues(elevesTrouves);
          logAction(`ULTRA-OPTIMISÉ: ${elevesTrouves.length} élèves ajoutés à ${defName}`);
        }
        
        // 8. Mise en forme
        const styleCfg = CFG.STYLE || {};
        
        // En-têtes
        targetSheet.getRange(1, 1, 1, headers.length)
          .setFontWeight("bold")
          .setBackground(styleCfg.HEADER_BG || "#C6E0B4")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
        
        // Largeurs de colonnes
        const colNomPrenomIndex = headers.indexOf("NOM & PRENOM") + 1;
        if (colNomPrenomIndex > 0) {
          targetSheet.setColumnWidth(colNomPrenomIndex, styleCfg.WIDTH_NOM_PRENOM || 200);
        }
        
        ["SEXE", "LV2", "OPT", "INDICATEUR", "ASSO", "DISSO"].forEach(colName => {
          const colIndex = headers.indexOf(colName) + 1;
          if (colIndex > 0) {
            targetSheet.setColumnWidth(colIndex, styleCfg.WIDTH_INFO_SHORT || 70);
          }
        });
        
        ["COM", "TRA", "PART", "ABS"].forEach(colName => {
          const colIndex = headers.indexOf(colName) + 1;
          if (colIndex > 0) {
            targetSheet.setColumnWidth(colIndex, styleCfg.WIDTH_CRITERE || 60);
          }
        });
        
        // Formatage conditionnel
        if (elevesTrouves.length > 0) {
          _phase5_appliquerFormatageCondSimple(targetSheet, elevesTrouves.length, headers, styleCfg);
          
          // Lignes alternées
          for (let r = 2; r <= elevesTrouves.length + 1; r++) {
            if (r % 2 === 0) {
              targetSheet.getRange(r, 1, 1, headers.length).setBackground(styleCfg.EVEN_ROW_BG || "#f8f9fa");
            }
          }
          
          // NOM & PRENOM en gras
          const colNomIndex = headers.indexOf("NOM & PRENOM") + 1;
          if (colNomIndex > 0) {
            targetSheet.getRange(2, colNomIndex, elevesTrouves.length, 1).setFontWeight("bold");
          }
          
          // Filtre
          targetSheet.getRange(1, 1, elevesTrouves.length + 1, headers.length).createFilter();
        }
        
        // 9. Ajouter les statistiques
        _phase5_ajouterStatsUltra(targetSheet, elevesTrouves, headers, styleCfg);
        
        // 10. Personnalisation finale
        targetSheet.setTabColor(styleCfg.DEF_TAB_COLOR || "#1a73e8");
        const protection = targetSheet.protect().setDescription(`Protection ${defName}`);
        protection.setWarningOnly(true);
        
        out.classes.push({ base: baseClass, ongletDef: defName, count: elevesTrouves.length });
      } catch (e) {
        logAction(`ULTRA-OPTIMISÉ: Erreur pour ${baseClass}: ${e.message}`);
        out.erreurs.push(`${baseClass}: ${e.message}`);
      }
    }
    
    // 11. Pas de création de bilan (supprimé pour gagner du temps)
    logAction("ULTRA-OPTIMISÉ: Bilan DEF ignoré pour améliorer les performances");
    
    // 12. Gérer la visibilité
    const defsCrees = out.classes.map(c => c.ongletDef);
    _phase5_gererVisibiliteOnglets(ss, masquerOngletsTemp ? defsCrees : [], !masquerOngletsTemp, CFG);

    return {
      success: out.erreurs.length === 0,
      message: out.erreurs.length ? `Finalisation partielle (${out.erreurs.length} erreurs).` : `${out.classes.length} classe(s) finalisée(s).`,
      details: out,
      classesAvecDef: Phase5_GetClassesDisponibles().classesAvecDef
    };
  } catch (e) {
    logAction(`ULTRA-OPTIMISÉ: Erreur majeure: ${e.message}`);
    return { success: false, errorCode: (CFG.ERROR_CODES || {}).UNCAUGHT_EXCEPTION || "ERR_UNCAUGHT", message: e.message };
  } finally {
    if (lock) lock.releaseLock();
  }
}

/** Formatage conditionnel simplifié */
function _phase5_appliquerFormatageCondSimple(sheet, nRows, headers, styleCfg) {
  try {
    sheet.clearConditionalFormatRules();
    const rules = [];
    
    // Fonction pour créer une règle avec moins de code
    const addRule = (colName, textValue, bgColor, fgColor = null) => {
      const colIndex = headers.indexOf(colName) + 1;
      if (colIndex > 0) {
        const range = sheet.getRange(2, colIndex, nRows, 1);
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(textValue)
          .setBackground(bgColor);
        
        if (fgColor) rule.setFontColor(fgColor);
        rules.push(rule.setRanges([range]).build());
        
        // Mettre en gras et centrer
        range.setFontWeight("bold").setHorizontalAlignment("center");
      }
    };
    
    // Fonction pour les scores numériques
    const addScoreRule = (colName, score, bgColor, fgColor = null) => {
      const colIndex = headers.indexOf(colName) + 1;
      if (colIndex > 0) {
        const range = sheet.getRange(2, colIndex, nRows, 1);
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberEqualTo(score)
          .setBackground(bgColor);
        
        if (fgColor) rule.setFontColor(fgColor);
        rules.push(rule.setRanges([range]).build());
        
        // Mettre en gras et centrer
        range.setFontWeight("bold").setHorizontalAlignment("center");
      }
    };
    
    // SEXE
    if (styleCfg.SEXE_F_COLOR) addRule("SEXE", "F", styleCfg.SEXE_F_COLOR);
    if (styleCfg.SEXE_M_COLOR) addRule("SEXE", "M", styleCfg.SEXE_M_COLOR, "#FFFFFF");
    
    // LV2
    if (styleCfg.LV2_COLORS) {
      if (styleCfg.LV2_COLORS.ESP) addRule("LV2", "ESP", styleCfg.LV2_COLORS.ESP);
      if (styleCfg.LV2_COLORS.ITA) addRule("LV2", "ITA", styleCfg.LV2_COLORS.ITA);
    }
    
    // OPTIONS
    if (styleCfg.OPTION_FORMATS && styleCfg.OPTION_FORMATS.length > 0) {
      styleCfg.OPTION_FORMATS.forEach(opt => {
        if (opt.text && opt.bgColor) {
          addRule("OPT", opt.text, opt.bgColor, opt.fgColor);
        }
      });
    }
    
    // CRITÈRES
    if (styleCfg.SCORE_COLORS && styleCfg.SCORE_FONT_COLORS) {
      ["COM", "TRA", "PART", "ABS"].forEach(critere => {
        for (let score = 1; score <= 4; score++) {
          const key = "S" + score;
          addScoreRule(critere, score, styleCfg.SCORE_COLORS[key], styleCfg.SCORE_FONT_COLORS[key]);
        }
      });
    }
    
    // Appliquer toutes les règles
    if (rules.length > 0) {
      sheet.setConditionalFormatRules(rules);
    }
  } catch (e) {
    logAction(`ULTRA-OPTIMISÉ: Erreur formatage conditionnel: ${e.message}`);
  }
}
// =========================================================
//         FICHIER : Phase5.V12.gs - PARTIE 4/5
//         Phase 5 (Finalisation) - Style "Maquette"
//         Version ULTRA-OPTIMISÉE avec colonnes utiles uniquement
// =========================================================

/** Statistiques ultra-simplifiées */
function _phase5_ajouterStatsUltra(sheet, eleves, headers, styleCfg) {
  if (eleves.length === 0) return;
  
  try {
    const statsStartRow = eleves.length + 3;
    
    // Titre
    sheet.getRange(statsStartRow, 1, 1, headers.length).merge()
      .setValue("STATISTIQUES DE LA CLASSE")
      .setFontWeight("bold")
      .setBackground(styleCfg.STATS_TITLE_BG || "#b6d7a8")
      .setHorizontalAlignment("center");
    
    const baseStatsRow = statsStartRow + 1;
    
    // Calculer les stats en mémoire
    const stats = {
      filles: 0,
      garcons: 0,
      esp: 0,
      autres: 0,
      options: 0,
      com: {1:0, 2:0, 3:0, 4:0, moy:0},
      tra: {1:0, 2:0, 3:0, 4:0, moy:0},
      part: {1:0, 2:0, 3:0, 4:0, moy:0},
      abs: {1:0, 2:0, 3:0, 4:0, moy:0}
    };
    
    // Indices des colonnes
    const idxSexe = headers.indexOf("SEXE");
    const idxLV2 = headers.indexOf("LV2");
    const idxOpt = headers.indexOf("OPT");
    const idxCom = headers.indexOf("COM");
    const idxTra = headers.indexOf("TRA");
    const idxPart = headers.indexOf("PART");
    const idxAbs = headers.indexOf("ABS");
    
    // Calcul des stats
    eleves.forEach(row => {
      // SEXE
      if (idxSexe >= 0) {
        const sexe = String(row[idxSexe]).trim().toUpperCase();
        if (sexe === "F") stats.filles++;
        else if (sexe === "M") stats.garcons++;
      }
      
      // LV2
      if (idxLV2 >= 0) {
        const lv2 = String(row[idxLV2]).trim().toUpperCase();
        if (lv2 === "ESP") stats.esp++;
        else if (lv2) stats.autres++;
      }
      
      // OPTIONS
      if (idxOpt >= 0 && String(row[idxOpt]).trim()) {
        stats.options++;
      }
      
      // CRITÈRES
      const processScore = (critere, idx) => {
        if (idx >= 0) {
          const scoreStr = String(row[idx]);
          const score = parseFloat(scoreStr.replace(',', '.'));
          if (!isNaN(score) && score >= 1 && score <= 4) {
            stats[critere][Math.round(score)]++;
          }
        }
      };
      
      processScore("com", idxCom);
      processScore("tra", idxTra);
      processScore("part", idxPart);
      processScore("abs", idxAbs);
    });
    
    // Calcul des moyennes
    ["com", "tra", "part", "abs"].forEach(crit => {
      let total = 0, count = 0;
      for (let i = 1; i <= 4; i++) {
        total += i * stats[crit][i];
        count += stats[crit][i];
      }
      stats[crit].moy = count > 0 ? total / count : 0;
    });
    
    // Écrire les stats
    const writeStats = (col, nom, valeur, style) => {
      const idx = headers.indexOf(nom);
      if (idx >= 0) {
        const cell = sheet.getRange(col.row, idx + 1);
        cell.setValue(valeur);
        
        if (style.bold) cell.setFontWeight("bold");
        if (style.bg) cell.setBackground(style.bg);
        if (style.fg) cell.setFontColor(style.fg);
        if (style.align) cell.setHorizontalAlignment(style.align);
        if (style.format) cell.setNumberFormat(style.format);
      }
    };
    
    // SEXE
    writeStats({row: baseStatsRow}, "SEXE", stats.filles, {bold: true, bg: styleCfg.SEXE_F_COLOR || "#F28EA8", align: "center"});
    writeStats({row: baseStatsRow + 1}, "SEXE", stats.garcons, {bold: true, bg: styleCfg.SEXE_M_COLOR || "#4F81BD", fg: "#FFFFFF", align: "center"});
    
    // LV2
    writeStats({row: baseStatsRow}, "LV2", stats.esp, {bold: true, bg: styleCfg.LV2_COLORS?.ESP || "#E59838", align: "center"});
    writeStats({row: baseStatsRow + 1}, "LV2", stats.autres, {bold: true, bg: styleCfg.LV2_COLORS?.AUTRE || "#A3E4D7", align: "center"});
    
    // OPTIONS
    writeStats({row: baseStatsRow}, "OPT", stats.options, {bold: true, align: "center"});
    
    // CRITÈRES - S'assurer que l'ordre des scores est 4, 3, 2, 1 (de haut en bas)
    ["COM", "TRA", "PART", "ABS"].forEach((nom, idx) => {
      const critKey = nom.toLowerCase();
      
      // Scores (4, 3, 2, 1)
      [4, 3, 2, 1].forEach((score, i) => {
        const style = {
          bold: true,
          bg: styleCfg.SCORE_COLORS?.[`S${score}`] || "#FFFFFF",
          fg: styleCfg.SCORE_FONT_COLORS?.[`S${score}`] || "#000000",
          align: "center"
        };
        writeStats({row: baseStatsRow + i}, nom, stats[critKey][score], style);
      });
      
      // Moyenne
      writeStats({row: baseStatsRow + 4}, nom, stats[critKey].moy, {
        bold: true, bg: "#34495e", fg: "#FFFFFF", align: "center", format: "0.00"
      });
    });
    
    // Bordure pour les stats
    sheet.getRange(statsStartRow, 1, 6, headers.length)
      .setBorder(true, true, true, true, true, true, styleCfg.STATS_BORDER_COLOR || "#000000", SpreadsheetApp.BorderStyle.SOLID);
    
  } catch (e) {
    logAction(`ULTRA-OPTIMISÉ: Erreur stats: ${e.message}`);
  }
}

/** Gère la visibilité des onglets. */
function _phase5_gererVisibiliteOnglets(ss, ongletsDefAffichesSiMasquagePartiel, toutAfficherSaufSiMasquage, CFG) {
  const onglets = ss.getSheets();
  const cfgSheets = CFG.SHEETS || {};
  const ongletsToujoursVisibles = [
      ...(CFG.PROTECTED_SHEETS || []),
      cfgSheets.BILAN_DEF, cfgSheets.BILAN_COMPARE, cfgSheets.STATS_FINAL 
  ].filter(Boolean);

  const reEscapeFn = typeof reEscape === 'function' ? reEscape : function(s) { return String(s).replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'); };
  const suffixTest = CFG.TEST_SUFFIX || "TEST";
  const suffixDef = CFG.DEF_SUFFIX || "DEF";
  const suffixTestRegex = new RegExp(reEscapeFn(suffixTest) + '$', 'i');
  const suffixDefRegex = new RegExp(reEscapeFn(suffixDef) + '$', 'i');

  logAction(`_phase5_gererVisibiliteOnglets: toutAfficher=${toutAfficherSaufSiMasquage}, defsVisiblesSiMasque=${(ongletsDefAffichesSiMasquagePartiel||[]).join(',')}`);

  onglets.forEach(onglet => {
    const nomOnglet = onglet.getName();
    let doitEtreVisible = false;

    if (toutAfficherSaufSiMasquage) {
        doitEtreVisible = true;
    } else { 
        if (ongletsToujoursVisibles.includes(nomOnglet)) {
            doitEtreVisible = true;
        } else if (suffixDefRegex.test(nomOnglet) && (ongletsDefAffichesSiMasquagePartiel || []).includes(nomOnglet)) {
            doitEtreVisible = true;
        } else {
            doitEtreVisible = false;
        }
    }
    try {
        if (doitEtreVisible && onglet.isSheetHidden()) onglet.showSheet();
        else if (!doitEtreVisible && !onglet.isSheetHidden()) onglet.hideSheet();
    } catch (e) { logAction(`WARN: Visibilité ${nomOnglet}: ${e.message}`); }
  });
  logAction(`Visibilité des onglets mise à jour.`);
}
// =========================================================
//         FICHIER : Phase5.V12.gs - PARTIE 5/5
//         Phase 5 (Finalisation) - Style "Maquette"
//         Version ULTRA-OPTIMISÉE avec colonnes utiles uniquement
// =========================================================

/**
 * Génère des PDF pour les classes sélectionnées,
 * vide les anciens fichiers du dossier « PDF Classes <NIVEAU> »,
 * puis renvoie l’URL du dossier et un message de confirmation.
 */
function Phase5_GenererPDF(classesBaseSelectionnees, avecColoration, CFG) {
  logAction(`Début PDF pour : ${classesBaseSelectionnees.join(', ')}`);
  let lock;
  try {
    // 1) Verrou
    lock = LockService.getDocumentLock();
    if (!lock.tryLock(CFG.LOCK_TIMEOUT_PDF || 20000)) {
      return { success: false, errorCode: CFG.ERROR_CODES.LOCK_TIMEOUT };
    }

    // 2) Validation
    if (!classesBaseSelectionnees || classesBaseSelectionnees.length === 0) {
      return { success: false, errorCode: CFG.ERROR_CODES.NO_CLASSES_SELECTED };
    }

    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const tz      = ss.getSpreadsheetTimeZone();
    const dateStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const niveau  = CFG.NIVEAU || '';
    
    // 3) Préparation du dossier unique par niveau, suppression des anciens PDFs
    const folderName = `PDF Classes ${niveau}`;
    let folder;
    const it = DriveApp.getFoldersByName(folderName);
    if (it.hasNext()) {
      folder = it.next();
      const oldFiles = folder.getFiles();
      while (oldFiles.hasNext()) {
        oldFiles.next().setTrashed(true);
      }
      logAction(`Anciens fichiers dans "${folderName}" effacés.`);
    } else {
      folder = DriveApp.createFolder(folderName);
      logAction(`Dossier "${folderName}" créé.`);
    }

    // 4) Configuration des paramètres d’export
    const baseUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export`;
    const token   = ScriptApp.getOAuthToken();
    const opts    = {
      exportFormat: 'pdf',
      format:       'pdf',
      size:         'A4',
      portrait:     'true',
      fitw:         'true',
      sheetnames:   'false',
      printtitle:   'false',
      pagenumbers:  'true',
      fzr:          'false'
    };

    // 5) Génération des PDF
    let count = 0;
    classesBaseSelectionnees.forEach(classe => {
      const sheet = ss.getSheetByName(classe + (CFG.DEF_SUFFIX || 'DEF'));
      if (!sheet) return;
      const gid = sheet.getSheetId();
      const params = Object.assign({}, opts, {
        gridlines: (!avecColoration).toString(),
        gid:       gid.toString()
      });
      const qs = Object.entries(params)
                       .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
                       .join('&');
      const url  = `${baseUrl}?${qs}`;
      const resp = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        const blob = resp.getBlob().setName(`${classe}_${dateStr}.pdf`);
        folder.createFile(blob);
        count++;
      } else {
        logAction(`Erreur HTTP ${resp.getResponseCode()} pour ${classe}`);
      }
    });

    // 6) Retour à l’UI
    return {
      success:   true,
      folderUrl: folder.getUrl(),
      message:   `${count} PDF généré${count > 1 ? 's' : ''} dans "${folderName}".`
    };

  } catch (e) {
    logAction(`Erreur Phase5_GenererPDF: ${e.message}`);
    return {
      success:   false,
      errorCode: CFG.ERROR_CODES.PDF_EXPORT_FAILED,
      message:   `Erreur génération PDF: ${e.message}`
    };
  } finally {
    if (lock) lock.releaseLock();
  }
}


/** Génère un fichier Excel. */
function Phase5_ExporterExcel(classesBaseSelectionnees, CFG) {
  logAction(`Début Excel pour: ${classesBaseSelectionnees.join(', ')}`);
   let lock;
   try {
    lock = LockService.getDocumentLock();
    if (!lock.tryLock(CFG.LOCK_TIMEOUT_EXCEL || 45000)) return { success: false, errorCode: (CFG.ERROR_CODES||{}).LOCK_TIMEOUT, message: "Verrouillage." };
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!classesBaseSelectionnees || classesBaseSelectionnees.length === 0) return { success: false, errorCode: (CFG.ERROR_CODES||{}).NO_CLASSES_SELECTED, message: "Aucune classe." };

    const niveau = CFG.NIVEAU || "Niveau";
    const fileName = `Export Classes ${niveau} ${CFG.DEF_SUFFIX || "DEF"} - ${new Date().toISOString().slice(0,10)}`;
    const newSS = SpreadsheetApp.create(fileName);
    const fileId = newSS.getId();
    const erreurs = [];

    classesBaseSelectionnees.forEach((classeBase, i) => {
      const sheetName = classeBase + (CFG.DEF_SUFFIX || "DEF");
      const sourceSheet = ss.getSheetByName(sheetName);
      if (!sourceSheet) { erreurs.push(`${sheetName} non trouvé.`); return; }
      try {
          const targetSheetName = classeBase.replace(/°/g, '');
          let targetSheet;
          if (i === 0 && newSS.getSheets().length === 1 && newSS.getSheets()[0].getName() === "Sheet1") {
              targetSheet = newSS.getSheets()[0].setName(targetSheetName);
          } else {
              targetSheet = newSS.insertSheet(targetSheetName);
          }
          
          const sourceData = sourceSheet.getDataRange().getValues();
          if (sourceData.length > 0 && sourceData[0].length > 0) {
              targetSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
              for(let k=1; k <= sourceSheet.getLastColumn(); k++) {
                   if (!sourceSheet.isColumnHiddenByUser(k)) targetSheet.setColumnWidth(k, sourceSheet.getColumnWidth(k));
              }
          }
      } catch (e) { 
          erreurs.push(`Erreur copie ${classeBase}: ${e.message}`); 
          logAction(`Erreur copie Excel ${classeBase}: ${e.message} ${e.stack}`);
      }
    });
    
    const defaultSheet = newSS.getSheetByName("Sheet1");
    if (defaultSheet && (newSS.getSheets().length > (classesBaseSelectionnees.length - erreurs.length) || ( (classesBaseSelectionnees.length - erreurs.length === 0) && newSS.getSheets().length === 1) ) ) {
        if(defaultSheet.getLastRow() === 0 && defaultSheet.getLastColumn() === 0) {
            if (newSS.getSheets().length > 1 || (classesBaseSelectionnees.length - erreurs.length) === 0) {
                 newSS.deleteSheet(defaultSheet);
            }
        }
    }
    logAction(`Excel exporté: ${fileName}. Erreurs: ${erreurs.length}`);
    return {
      success: erreurs.length === 0,
      message: erreurs.length === 0 ? `Export Excel '${fileName}' créé.` : `Export Excel '${fileName}' créé avec ${erreurs.length} erreur(s).`,
      files: [{ name: `${fileName}.xlsx`, url: `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`, type: 'excel' }],
      erreurs, details: { fileName: fileName, docId: fileId }
    };
  } catch (e) { 
    logAction(`Erreur globale Phase5_ExporterExcel: ${e.message} ${e.stack}`);
    return { success: false, errorCode: (CFG.ERROR_CODES||{}).EXCEL_EXPORT_FAILED, message: `Erreur export Excel: ${e.message}` };
  } finally { if (lock) lock.releaseLock(); }
}

/** Réinitialise. */
function Phase5_ReinitialiserTout(CFG) {
  logAction("Début Réinitialisation Phase 5...");
  let lock;
  try {
    lock = LockService.getDocumentLock();
    if (!lock.tryLock(CFG.LOCK_TIMEOUT_RESET || 30000)) return { success: false, errorCode: (CFG.ERROR_CODES||{}).LOCK_TIMEOUT, message: "Verrouillage." };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ongletsASupprimerNoms = [];
    const reEscapeFn = typeof reEscape === 'function' ? reEscape : function(s) { return String(s).replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'); };
    const suffixDef = CFG.DEF_SUFFIX || "DEF";
    const suffixDefRegex = new RegExp(reEscapeFn(suffixDef) + '$', 'i');
    
    const cfgSheets = CFG.SHEETS || {};
    const bilansASupprimer = [cfgSheets.BILAN_DEF, cfgSheets.BILAN_COMPARE, cfgSheets.STATS_FINAL].filter(Boolean);
    const protectedSheets = CFG.PROTECTED_SHEETS || [];

    ss.getSheets().forEach(onglet => {
      const nomOnglet = onglet.getName();
      if (suffixDefRegex.test(nomOnglet) || bilansASupprimer.includes(nomOnglet) ) {
          if(!protectedSheets.includes(nomOnglet)) ongletsASupprimerNoms.push(nomOnglet);
      }
    });
    logAction(`Onglets à supprimer: ${ongletsASupprimerNoms.join(', ')}`);
    let deletedCount = 0;
    ongletsASupprimerNoms.forEach(nom => {
      try { const sheet = ss.getSheetByName(nom); if(sheet) { ss.deleteSheet(sheet); deletedCount++; }}
      catch (e) { logAction(`WARN: Suppr ${nom}: ${e.message}`); }
    });
    _phase5_gererVisibiliteOnglets(ss, [], true, CFG);
    logAction(`Réinitialisation terminée: ${deletedCount} onglets supprimés.`);
    return { success: true, message: `Réinitialisation: ${deletedCount} onglet(s) supprimé(s).`, details: {ongletsSupprimes: ongletsASupprimerNoms}};
  } catch (e) { 
      logAction(`Erreur Phase5_ReinitialiserTout: ${e.message} ${e.stack}`);
      return { success: false, errorCode: (CFG.ERROR_CODES||{}).RESET_FAILED, message: `Erreur Reset: ${e.message}` };
  } finally { if (lock) lock.releaseLock(); }
}

/**
 * Point d'entrée pour les actions de la sidebar de finalisation
 */
function Phase5_ProcessAction(request) {
  Logger.log(">>> Phase5_ProcessAction: DÉBUT DE LA FONCTION <<<");
  Logger.log("Phase5_ProcessAction: Request reçu: " + JSON.stringify(request));

  let CFG;
  try {
    CFG = getConfig();
    Logger.log("Phase5_ProcessAction: getConfig() terminé. CFG.NIVEAU (test): " + (CFG ? CFG.NIVEAU : "CFG non défini"));
  } catch (e) {
    Logger.log("!!! ERREUR CRITIQUE dans Phase5_ProcessAction lors de l'appel à getConfig(): " + e.message);
    SpreadsheetApp.getUi().alert("Erreur Fatale", "Impossible de charger la configuration : " + e.message);
    return { success: false, message: "Erreur chargement configuration: " + e.message };
  }

  if (!CFG) return { success: false, message: "Erreur fatale: CFG indéfini." };
  if (!CFG.ERROR_CODES) return { success: false, message: "Erreur fatale: CFG.ERROR_CODES indéfini." };

  const action  = request.action;
  const classes = request.classes || [];
  let result;

  logAction(`Phase 5 Action reçue: ${action} pour classes: ${classes.join(', ')}`);
  Logger.log(`Phase5_ProcessAction: Action = ${action}, Classes = ${classes.join(', ')}`);

  switch (action) {
    case "finaliser":
      Logger.log("Phase5_ProcessAction: Appel de Phase5_FinaliserClasses...");
      try {
        result = Phase5_FinaliserClasses(classes, request.hideTmp || false, CFG);
      } catch (e) {
        Logger.log("!!! ERREUR LORS DE L'APPEL à Phase5_FinaliserClasses: " + e.message);
        result = { success: false, message: "Erreur appel Phase5_FinaliserClasses: " + e.message, errorCode: CFG.ERROR_CODES.UNCAUGHT_EXCEPTION };
      }
      return result;

    // alias pour exporterPDF → genererPDF
    case "exporterPDF":
      logAction("Phase5_ProcessAction: Alias exporterPDF → genererPDF");
    case "genererPDF":
      Logger.log("Phase5_ProcessAction: Appel de Phase5_GenererPDF...");
      return Phase5_GenererPDF(classes, request.avecColoration || false, CFG);

    case "exporterExcel":
      Logger.log("Phase5_ProcessAction: Appel de Phase5_ExporterExcel...");
      return Phase5_ExporterExcel(classes, CFG);

    case "comparerClasses":
      logAction("Action 'comparerClasses' non implémentée séparément.");
      return { success: false, message: "Comparaison intégrée au Bilan DEF." };

    case "reinitialiser":
      Logger.log("Phase5_ProcessAction: Appel de Phase5_ReinitialiserTout...");
      return Phase5_ReinitialiserTout(CFG);

    case "toggleVisibilite":
      Logger.log("Phase5_ProcessAction: Appel de _phase5_gererVisibiliteOnglets...");
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const ongletsDef = classes.map(c => c + (CFG.DEF_SUFFIX || "DEF"));
      _phase5_gererVisibiliteOnglets(ss, ongletsDef, !request.masquer, CFG);
      return { success: true, message: request.masquer ? "Onglets non finalisés masqués." : "Tous les onglets sont visibles." };

    default:
      logAction(`Action Phase 5 non reconnue: ${action}`);
      Logger.log(`Phase5_ProcessAction: Action non reconnue - ${action}`);
      return { success: false, errorCode: CFG.ERROR_CODES.ACTION_NOT_RECOGNIZED, message: `Action non reconnue: ${action}` };
  }
}

// UI Wrappers
function Phase5_GetClassesDisponiblesPourUI() { 
  return Phase5_GetClassesDisponibles(); 
}
function Phase5_ProcessActionPourUI(request) { 
  return Phase5_ProcessAction(request); 
}

// Fonctions de Menu
function showOptimisationSidebar() {
  try {
    const html = HtmlService
      .createHtmlOutputFromFile('Phase4UI.html')
      .setTitle("Optimisation")
      .setWidth(450);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch(e) {
    try { logAction(`Erreur Sidebar Phase 4: ${e.message}`); }
    catch(_) { Logger.log(`Erreur Sidebar Phase 4: ${e.message}`); }
    SpreadsheetApp.getUi().alert(`Erreur: Sidebar Optimisation (${e.message})`);
  }
}
function showFinalisationSidebar() {
  try {
    const html = HtmlService
      .createHtmlOutputFromFile('FinalisationUI.html')
      .setTitle("Finalisation Phase 5")
      .setWidth(550);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch(e) {
    try { logAction(`Erreur Sidebar Phase 5: ${e.message}`); }
    catch(_) { Logger.log(`Erreur Sidebar Phase 5: ${e.message}`); }
    SpreadsheetApp.getUi().alert(`Erreur: Sidebar Finalisation (${e.message})`);
  }
}


/**
 * Utilitaire pour afficher les valeurs uniques d'une colonne dans une feuille.
 * Pratique pour le dépannage.
 */
function _phase5_analyseColonneValeurs(sheet, colIndex_1based, colNameForLog) {
  const localLog = typeof logAction === 'function' ? logAction : Logger.log;
  if (!sheet || colIndex_1based < 1) {
    localLog(`_phase5_analyseColonneValeurs: Feuille ou index de colonne invalide.`);
    return [];
  }
  try {
    if (colIndex_1based > sheet.getMaxColumns()) {
        localLog(`_phase5_analyseColonneValeurs: Index ${colIndex_1based} > max cols ${sheet.getMaxColumns()} pour ${sheet.getName()}`);
        return [];
    }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
        localLog(`_phase5_analyseColonneValeurs: Pas de données dans ${sheet.getName()}`);
        return [];
    }
    const values = data.slice(1).map(row => String(row[colIndex_1based-1] || "").trim()).filter(val => val !== "");
    const uniqueValues = [...new Set(values)];
    localLog(`Analyse colonne "${colNameForLog}" (${colIndex_1based}) dans ${sheet.getName()}: ${uniqueValues.length} uniques.`);
    if (uniqueValues.length > 0 && uniqueValues.length < 20) { // Limiter le log si trop de valeurs
      localLog(`Valeurs uniques: [${uniqueValues.join(", ")}]`);
    } else if (uniqueValues.length >= 20) {
      localLog(`Valeurs uniques (échantillon): [${uniqueValues.slice(0,20).join(", ")}...]`);
    }
    return uniqueValues;
  } catch (e) {
    localLog(`ERREUR _phase5_analyseColonneValeurs col "${colNameForLog}" sheet ${sheet.getName()}: ${e.message}`);
    return [];
  }
}