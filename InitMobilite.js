/**
 * InitMobilite.gs — génère ou met à jour la colonne T (MOBILITE)
 * PREND EN COMPTE LV2 COMME OPTION SI LA COLONNE OPT EST VIDE.
 */

// CONFIG LOCALES POUR CE SCRIPT
const MOB_TARGET_COLUMN_LETTER_INIT = 'T';
const MOB_HEADER_TEXT_INIT          = 'MOBILITE';
const OPTION_SEPARATOR_RE_INIT      = /[,/]/; 

/** 
 * Fonction principale pour initialiser la mobilité dans la colonne cible (T).
 */
function initMobilite() {
  const t0 = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const cfg = (typeof getConfig === 'function') ? getConfig() : { TEST_SUFFIX: 'TEST', SHEETS: { STRUCTURE: '_STRUCTURE' } };
  if (!cfg.SHEETS) cfg.SHEETS = { STRUCTURE: '_STRUCTURE' };

  Logger.log("InitMobilite: Démarrage du processus...");
  Logger.log(`InitMobilite: Config - TEST_SUFFIX: '${cfg.TEST_SUFFIX}', STRUCTURE: '${cfg.SHEETS.STRUCTURE}'`);

  let optionPools;
  try {
    optionPools = buildOptionPoolsForInitMobilite(cfg); 
    Logger.log("InitMobilite: Pools d'options détectés (après normalisation) : " + JSON.stringify(optionPools));
  } catch (e) {
    Logger.log("ERREUR CRITIQUE InitMobilite: buildOptionPoolsForInitMobilite a échoué: " + e.message);
    Logger.log(e.stack);
    ui.alert("Erreur InitMobilité", "Erreur construction pools d'options: " + e.message, ui.ButtonSet.OK);
    return { success: false, message: "Erreur pools d'options: " + e.message };
  }

  const testSheets = getTestSheetsFromInitMobilite(cfg.TEST_SUFFIX);
  if (testSheets.length === 0) {
    Logger.log("InitMobilite: Aucune feuille TEST trouvée.");
    ui.alert("InitMobilité", "Aucune feuille ...TEST trouvée.", ui.ButtonSet.OK);
    return { success: false, message: "Aucune feuille TEST" };
  }
  Logger.log(`InitMobilite: ${testSheets.length} feuille(s) TEST à traiter: [${testSheets.map(s=>s.getName()).join(', ')}]`);

  const targetColNumber = MOB_TARGET_COLUMN_LETTER_INIT.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
  
  let totalUpdatedActual = 0;
  let totalProcessedWithId = 0;
  let missingColumnsLogged = {}; 

  for (let i = 0; i < testSheets.length; i++) {
    const sheet = testSheets[i];
    const sheetName = sheet.getName();
    Logger.log(`InitMobilite: Traitement de la feuille '${sheetName}'...`);
    
    const range = sheet.getDataRange();
    const data = range.getValues();

    if (data.length <= 1) {
      Logger.log(`InitMobilite: Feuille '${sheetName}' ignorée (vide ou en-tête).`);
      continue; 
    }

    if (sheet.getMaxColumns() < targetColNumber) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), targetColNumber - sheet.getMaxColumns());
    }
    
    const headerCell = sheet.getRange(1, targetColNumber);
    if (String(headerCell.getValue()).trim().toUpperCase() !== MOB_HEADER_TEXT_INIT.toUpperCase()) {
        headerCell.setValue(MOB_HEADER_TEXT_INIT);
    }
    
    const headers = data[0].map(h => String(h).trim().toUpperCase());
    const idEleveColIndex = headers.indexOf('ID_ELEVE'); 
    const optionColIndex = headers.indexOf('OPT');
    const dissoColIndex = headers.indexOf('DISSO');
    const assoColIndex = headers.indexOf('ASSO');
    const lv2ColIndex = headers.indexOf('LV2'); // <<< RÉCUPÉRER L'INDEX DE LA COLONNE LV2

    if (idEleveColIndex === -1) {
      Logger.log(`ERREUR InitMobilite: Colonne 'ID_ELEVE' non trouvée dans '${sheetName}'.`);
      continue; 
    }

    if (!missingColumnsLogged[sheetName]) {
        if (optionColIndex === -1) Logger.log(`WARN InitMobilite: Colonne 'OPT' non trouvée dans '${sheetName}'.`);
        if (lv2ColIndex === -1) Logger.log(`WARN InitMobilite: Colonne 'LV2' non trouvée dans '${sheetName}'. Le fallback sur LV2 ne sera pas possible.`);
        if (dissoColIndex === -1) Logger.log(`WARN InitMobilite: Colonne 'DISSO' non trouvée dans '${sheetName}'.`);
        if (assoColIndex === -1) Logger.log(`WARN InitMobilite: Colonne 'ASSO' non trouvée dans '${sheetName}'.`);
        missingColumnsLogged[sheetName] = true;
    }
    
    const valuesToSetInMobilityColumn = []; 

    for (let r = 1; r < data.length; r++) { 
      const rowData = data[r];
      const idEleve = (idEleveColIndex >= 0 && idEleveColIndex < rowData.length) ? rowData[idEleveColIndex] : null; 

      if (idEleve === null || String(idEleve).trim() === '') {
        valuesToSetInMobilityColumn.push([""]); 
        continue;
      }
      totalProcessedWithId++;

      // --- APPLICATION DE VOTRE PATCH ICI ---
      const studentData = {
        id:    String(idEleve).trim(),
        opt:   (optionColIndex >= 0 && optionColIndex < rowData.length)
                 ? String(rowData[optionColIndex] || '').trim()
                 : '',
        lv2:   (lv2ColIndex >= 0 && lv2ColIndex < rowData.length)   // Lire LV2
                 ? String(rowData[lv2ColIndex] || '').trim()
                 : '',
        disso: (dissoColIndex >= 0 && dissoColIndex < rowData.length)
                 ? String(rowData[dissoColIndex] || '').trim().toUpperCase()
                 : '',
        asso:  (assoColIndex >= 0 && assoColIndex < rowData.length)
                 ? String(rowData[assoColIndex] || '').trim()
                 : ''
      };

      // Si OPT est vide mais LV2 est présente, utiliser LV2 comme option pour la mobilité
      if (studentData.opt === '' && studentData.lv2 !== '') {
        studentData.opt = studentData.lv2; 
        Logger.log(`InitMobilite: Pour élève ${studentData.id} dans ${sheetName}, OPT était vide, LV2 ('${studentData.lv2}') utilisée comme option pour la mobilité.`);
      }
      // --- FIN DU PATCH ---
      
      const mobilityValue = decideMobilityForInitMobilite(studentData, optionPools, sheetName); 
      valuesToSetInMobilityColumn.push([mobilityValue]); 

      const existingMobility = (targetColNumber -1 < rowData.length) ? String(rowData[targetColNumber-1] || "").trim().toUpperCase() : "";
      if (existingMobility !== mobilityValue) {
        totalUpdatedActual++;
      }
    }

    if (valuesToSetInMobilityColumn.length > 0) {
      try {
        const targetColumnRangeForWrite = sheet.getRange(2, targetColNumber, valuesToSetInMobilityColumn.length, 1);
        const targetColumnForHide = sheet.getRange(1, targetColNumber); 

        if (sheet.isColumnHiddenByUser(targetColNumber)) {
            sheet.unhideColumn(targetColumnForHide);
        }
        
        targetColumnRangeForWrite.setValues(valuesToSetInMobilityColumn);
        Logger.log(`InitMobilite: '${sheetName}' - ${valuesToSetInMobilityColumn.length} valeurs écrites en col ${MOB_TARGET_COLUMN_LETTER_INIT}.`);
        
        sheet.hideColumns(targetColNumber);
      } catch (e) {
        Logger.log(`ERREUR InitMobilite: Écriture/Masquage col ${MOB_TARGET_COLUMN_LETTER_INIT} de '${sheetName}': ${e.message}\n${e.stack}`);
      }
    }
    SpreadsheetApp.flush(); 
  } 

  const duration = (new Date().getTime() - t0.getTime()) / 1000;
  const message = `InitMobilité: ${totalUpdatedActual} mobilités modifiées pour ${totalProcessedWithId} élèves avec ID. Durée: ${duration.toFixed(1)}s.`;
  Logger.log(message);
  ui.alert("Rapport InitMobilité", message, ui.ButtonSet.OK);
  
  return { success: true, countUpdated: totalUpdatedActual, countProcessed: totalProcessedWithId, message: message };
}

/** 
 * Construit la table { OPTION_NORMALISEE: [CLASSE_TEST_MAJUSCULE, ...] } 
 * Spécifique pour InitMobilite.
 */
function buildOptionPoolsForInitMobilite(cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const structureSheetName = cfg.SHEETS && cfg.SHEETS.STRUCTURE ? cfg.SHEETS.STRUCTURE : '_STRUCTURE';
  const structSheet = ss.getSheetByName(structureSheetName);
  Logger.log(`buildOptionPoolsForInitMobilite: Lecture de la feuille de structure '${structureSheetName}'...`);

  if (!structSheet) {
    throw new Error(`Feuille de structure '${structureSheetName}' introuvable.`);
  }
  
  const data = structSheet.getDataRange().getValues();
  if (data.length <= 1) {
      Logger.log(`buildOptionPoolsForInitMobilite: Feuille de structure '${structureSheetName}' est vide ou ne contient que l'en-tête.`);
      return {};
  }

  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const classeCol = headers.indexOf('CLASSE_DEST');
  const optionsCol = headers.indexOf('OPTIONS');

  if (classeCol === -1 || optionsCol === -1) {
    throw new Error(`Colonnes 'CLASSE_DEST' ou 'OPTIONS' manquantes dans '${structureSheetName}'.`);
  }
  Logger.log(`buildOptionPoolsForInitMobilite: Indices - CLASSE_DEST=${classeCol}, OPTIONS=${optionsCol}`);

  const pools = {};
  const suffixConfig = String(cfg.TEST_SUFFIX || 'TEST'); 

  for (let i = 1; i < data.length; i++) { 
    const rowData = data[i];
    const classeRaw = (classeCol < rowData.length) ? String(rowData[classeCol] || "").trim() : "";
    if (!classeRaw) continue; 
    
    let nomClassePourPool = classeRaw;
    if (!classeRaw.toUpperCase().endsWith(suffixConfig.toUpperCase())) { 
        nomClassePourPool = classeRaw + suffixConfig; 
    }
    const nomClasseNormalisePourPool = nomClassePourPool.toUpperCase(); 
    Logger.log(`DEBUG buildOptionPools: Classe Ligne ${i+1}: '${classeRaw}' -> '${nomClasseNormalisePourPool}'`);
    
    const optsRaw = (optionsCol < rowData.length) ? String(rowData[optionsCol] || "").trim() : "";
    if (!optsRaw) continue; 

    optsRaw.split(OPTION_SEPARATOR_RE_INIT).forEach(opt => {
      let optionKey = opt.trim().toUpperCase();
      if (!optionKey) return; 
      
      if (optionKey.includes('=')) { 
          optionKey = optionKey.split('=')[0].trim();
      }
      if (!optionKey) return; 

      Logger.log(`DEBUG buildOptionPools: Option brute '${opt}', Clé normalisée: '${optionKey}', ClassePourPool='${nomClasseNormalisePourPool}'`);

      if (!pools[optionKey]) {
        pools[optionKey] = [];
      }
      if (!pools[optionKey].includes(nomClasseNormalisePourPool)) { 
          pools[optionKey].push(nomClasseNormalisePourPool);
          Logger.log(`DEBUG buildOptionPools: Option '${optionKey}' ajoutée pour classe '${nomClasseNormalisePourPool}'`);
      }
    });
  }
  Logger.log(`buildOptionPoolsForInitMobilite: Pools finaux: ${JSON.stringify(pools)}`);
  return pools;
}

/** 
 * Retourne la mobilité (FIXE, PERMUT, CONDI, SPEC, LIBRE) pour un élève.
 * Spécifique pour InitMobilite. Normalise les clés d'option de l'élève.
 */
function decideMobilityForInitMobilite(student, optionPools, sheetNameForLog) { 
  const studentIdForLog = student.id || 'ID_INCONNU';
  // LOG 1: Afficher les données brutes de l'élève et le contexte
  Logger.log(`DEBUG decideMobility: Élève ${studentIdForLog} (Feuille: ${sheetNameForLog||'N/A'}), OPT BRUTE='${student.opt}', DISSO='${student.disso}', ASSO='${student.asso}'`); 

  // Condition 1: Vérifier le code ASSO
  if (student.asso && String(student.asso).trim() !== '') {
    Logger.log(`DEBUG decideMobility: Élève ${studentIdForLog} => SPEC (cause: ASSO='${student.asso}')`); 
    return 'SPEC';
  }

  // Condition 2: Gérer les options (si student.opt n'est pas vide)
  let mainOptionKey = '';
  if (student.opt && String(student.opt).trim() !== '') {
    const studentOptValue = String(student.opt).trim().toUpperCase(); // Option de l'élève, normalisée pour la logique interne
    // Extrait la première partie de l'option (ex: "CHAV" de "CHAV=12" ou "CHAV/EUR")
    const firstOptPart = studentOptValue.split(OPTION_SEPARATOR_RE_INIT)[0].trim(); 
    // Normalise en enlevant la partie "=X" si elle existe
    mainOptionKey = firstOptPart.includes('=') ? firstOptPart.split('=')[0].trim() : firstOptPart;
    
    // LOG 2: Afficher l'option traitée et la clé d'option extraite
    Logger.log(`DEBUG decideMobility: Élève ${studentIdForLog}, student.opt traitée='${studentOptValue}', mainOptionKey extraite='${mainOptionKey}'`); 

    if (mainOptionKey) { // Si une clé d'option principale a été valablement extraite
        const pool = optionPools[mainOptionKey] || []; // Récupère le pool pour cette option (les clés dans optionPools sont déjà normalisées et en MAJUSCULES)
        // LOG 3: Afficher le pool trouvé pour cette option et sa taille
        Logger.log(`DEBUG decideMobility: Élève ${studentIdForLog}, Option cherchée dans pools='${mainOptionKey}', Pool trouvé=${JSON.stringify(pool)}, TaillePool=${pool.length}`); 
        
        if (pool.length === 1) { // L'option n'est proposée que dans UNE seule classe
            Logger.log(`DEBUG decideMobility: Élève ${studentIdForLog} => FIXE (option '${mainOptionKey}' dans pool de taille 1)`); 
            return 'FIXE'; 
        }
        if (pool.length >= 2) { // L'option est partagée entre AU MOINS DEUX classes
            const mobility = (student.disso && String(student.disso).trim() !== '') ? 'CONDI' : 'PERMUT';
            Logger.log(`DEBUG decideMobility: Élève ${studentIdForLog} => ${mobility} (option '${mainOptionKey}' dans pool de taille ${pool.length}, Disso='${student.disso}')`); 
            return mobility;
        }
        // Si pool.length est 0: l'option de l'élève (mainOptionKey) n'a pas été trouvée comme clé dans optionPools.
        if (pool.length === 0) { 
             Logger.log(`WARN MOBILITE: Élève ${studentIdForLog} (Feuille: ${sheetNameForLog||'N/A'}), mainOptionKey '${mainOptionKey}' NON TROUVÉE comme clé dans optionPools. L'élève sera traité comme n'ayant pas cette option pour la mobilité.`); 
        }
    } else { // Si, après traitement, mainOptionKey est vide (ex: student.opt était "=")
         Logger.log(`DEBUG MOBILITE: Élève ${studentIdForLog} (Feuille: ${sheetNameForLog||'N/A'}), Aucune mainOptionKey valide extraite de student.opt='${student.opt}'.`); 
    }
  } // Fin du bloc if (student.opt && ...)

  // Condition 3: Si pas d'option contraignante (ou option non trouvée dans les pools), vérifier DISSO
  if (student.disso && String(student.disso).trim() !== '') {
    Logger.log(`DEBUG decideMobility: Élève ${studentIdForLog} => CONDI (pas d'option contraignante mais DISSO: '${student.disso}')`); 
    return 'CONDI';
  }

  // Condition 4: Par défaut, si aucune des conditions ci-dessus n'est remplie
  Logger.log(`DEBUG decideMobility: Élève ${studentIdForLog} => LIBRE (par défaut)`); 
  return 'LIBRE';
}

/** 
 * Renvoie le tableau des feuilles dont le nom se termine par le suffixe donné.
 * Spécifique pour InitMobilite.
 */
function getTestSheetsFromInitMobilite(suffix) {
  // ... (Identique à la version précédente)
  const sheetSuffix = suffix || 'TEST'; 
  const regex = new RegExp(sheetSuffix + '$', 'i'); 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const filteredSheets = [];
  for (let i = 0; i < allSheets.length; i++) {
    if (regex.test(allSheets[i].getName())) {
      filteredSheets.push(allSheets[i]);
    }
  }
  return filteredSheets;
}

// --- forceInitMobilite (Mise à jour pour utiliser la même logique de fallback LV2) ---
function forceInitMobilite() {
  const t0 = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi(); 

  const cfg = (typeof getConfig === 'function') ? getConfig() : { TEST_SUFFIX: 'TEST', SHEETS: { STRUCTURE: '_STRUCTURE' } }; 
  if (!cfg.SHEETS) cfg.SHEETS = { STRUCTURE: '_STRUCTURE' };

  Logger.log("forceInitMobilite: Démarrage du processus...");
  Logger.log(`forceInitMobilite: Config - TEST_SUFFIX: '${cfg.TEST_SUFFIX}', STRUCTURE: '${cfg.SHEETS.STRUCTURE}'`);

  let optionPools;
  try {
     optionPools = buildOptionPoolsForInitMobilite(cfg); 
     Logger.log('forceInitMobilite: Pools d\'option détectés (après normalisation) : ' + JSON.stringify(optionPools));
  } catch (e) {
     Logger.log("ERREUR CRITIQUE forceInitMobilite: buildOptionPoolsForInitMobilite a échoué: " + e.message);
     Logger.log(e.stack);
     ui.alert("Erreur InitMobilité (Force)", "Erreur construction pools d'options: " + e.message, ui.ButtonSet.OK);
     return { success: false, message: "Erreur pools d'options (Force): " + e.message };
  }

  const testSheets = getTestSheetsFromInitMobilite(cfg.TEST_SUFFIX); 
  if (testSheets.length === 0) {
    Logger.log('forceInitMobilite : aucune feuille …' + cfg.TEST_SUFFIX + ' trouvée.');
    ui.alert('InitMobilité (Force)', 'aucune feuille …' + cfg.TEST_SUFFIX + ' trouvée');
    return { success: false, message: "Aucune feuille TEST trouvée (Force)" };
  }
   Logger.log(`forceInitMobilite: ${testSheets.length} feuille(s) TEST à traiter: [${testSheets.map(s=>s.getName()).join(', ')}]`);

  const targetColNumber = MOB_TARGET_COLUMN_LETTER_INIT.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
  
  let totalRowsProcessedWithIdInForce = 0; 
  
  for (let i = 0; i < testSheets.length; i++) { 
    const sheet = testSheets[i];
    const sheetName = sheet.getName();
    Logger.log(`forceInitMobilite: Traitement de la feuille '${sheetName}'...`);
    
    const range = sheet.getDataRange();
    const data = range.getValues();
    if (data.length <= 1) {
       Logger.log(`forceInitMobilite: Feuille '${sheetName}' ignorée (vide ou en-tête).`);
       continue; 
    }

    if (sheet.getMaxColumns() < targetColNumber) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), targetColNumber - sheet.getMaxColumns());
    }
    
    const headerCell = sheet.getRange(1, targetColNumber);
    if (String(headerCell.getValue()).trim().toUpperCase() !== MOB_HEADER_TEXT_INIT.toUpperCase()) {
         headerCell.setValue(MOB_HEADER_TEXT_INIT);
    }

    const headers = data[0].map(h => String(h).trim().toUpperCase());
    const idColIndex = headers.indexOf('ID_ELEVE');
    const optionCol = headers.indexOf('OPT');
    const dissoCol = headers.indexOf('DISSO');
    const assoCol = headers.indexOf('ASSO');
    const lv2ColIndex = headers.indexOf('LV2'); // <<< AJOUT pour forceInitMobilite aussi

    if (idColIndex === -1) {
      Logger.log(`ERREUR forceInitMobilite: Colonne 'ID_ELEVE' non trouvée dans '${sheetName}'.`);
      continue; 
    }

    const valuesToSetInMobilityColumn = []; 

    for (let r = 1; r < data.length; r++) { 
      const rowData = data[r]; 
      const idEleve = (idColIndex >= 0 && idColIndex < rowData.length) ? rowData[idColIndex] : null;
      if (idEleve === null || String(idEleve).trim() === '') {
        valuesToSetInMobilityColumn.push([""]); 
        continue; 
      }
      totalRowsProcessedWithIdInForce++;
      
      // --- APPLICATION DE VOTRE PATCH ICI AUSSI ---
      const studentData = {
        id:    String(idEleve).trim(),
        opt:   (optionCol >= 0 && optionCol < rowData.length) ? String(rowData[optionCol] || '').trim() : '',
        lv2:   (lv2ColIndex >= 0 && lv2ColIndex < rowData.length) ? String(rowData[lv2ColIndex] || '').trim() : '', // Lire LV2
        disso: (dissoCol >= 0 && dissoCol < rowData.length) ? String(rowData[dissoCol] || '').trim().toUpperCase() : '',
        asso:  (assoCol >= 0 && assoCol < rowData.length) ? String(rowData[assoCol] || '').trim() : ''
      };

      // Si OPT est vide mais LV2 est présente, utiliser LV2 comme option pour la mobilité
      if (studentData.opt === '' && studentData.lv2 !== '') {
        studentData.opt = studentData.lv2; 
        Logger.log(`forceInitMobilite: Pour élève ${studentData.id} dans ${sheetName}, OPT était vide, LV2 ('${studentData.lv2}') utilisée comme option.`);
      }
      // --- FIN DU PATCH ---
      
      const mobilityValue = decideMobilityForInitMobilite(studentData, optionPools, sheetName); 
      valuesToSetInMobilityColumn.push([mobilityValue]); 
    }

     if (valuesToSetInMobilityColumn.length > 0) {
        try {
            const targetColumnRangeForWrite = sheet.getRange(2, targetColNumber, valuesToSetInMobilityColumn.length, 1);
            const targetColumnForHide = sheet.getRange(1, targetColNumber); 

            if (sheet.isColumnHiddenByUser(targetColNumber)) {
                sheet.unhideColumn(targetColumnForHide);
            }

            targetColumnRangeForWrite.setValues(valuesToSetInMobilityColumn);
            Logger.log(`forceInitMobilite: ${valuesToSetInMobilityColumn.length} valeurs écrites dans col ${MOB_TARGET_COLUMN_LETTER_INIT} de '${sheetName}'.`);

            sheet.hideColumns(targetColNumber);
        } catch (e) {
             Logger.log(`ERREUR forceInitMobilite: Écriture/Masquage col ${MOB_TARGET_COLUMN_LETTER_INIT} de '${sheetName}': ${e.message}\n${e.stack}`);
        }
    }
     SpreadsheetApp.flush(); 
  } 

  const duration = (new Date().getTime() - t0.getTime()) / 1000;
  const message = `InitMobilité (Force) terminé. ${totalRowsProcessedWithIdInForce} élèves avec ID ont eu leur mobilité recalculée/écrite. Durée: ${duration.toFixed(1)}s.`;
  Logger.log(message);
  ui.alert('Rapport InitMobilité (Force)', message, ui.ButtonSet.OK);
  
  return { success: true, countProcessed: totalRowsProcessedWithIdInForce, message: message };
}