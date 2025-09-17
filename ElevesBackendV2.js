// Ajouter cette fonction dans votre fichier ElevesBackendV2.gs

/******************** updateStructureRules ***********************/
/**
 * Met à jour les règles dans l'onglet _STRUCTURE
 * @param {Object} newRules - Objet avec les nouvelles règles
 * Format: {
 *   "5°1": { capacity: 28, quotas: { ESP: 12, ITA: 8, CHAV: 6 } },
 *   ...
 * }
 */
function updateStructureRules(newRules) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ELEVES_MODULE_CONFIG.STRUCTURE_SHEET);
    if (!sh) {
      // Créer l'onglet s'il n'existe pas
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      sh = ss.insertSheet(ELEVES_MODULE_CONFIG.STRUCTURE_SHEET);
      
      // Ajouter les en-têtes
      sh.getRange(1, 1, 1, 4).setValues([["", "CLASSE_DEST", "EFFECTIF", "OPTIONS"]]);
      sh.getRange(1, 1, 1, 4)
        .setBackground('#5b21b6')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }
    
    const data = sh.getDataRange().getValues();
    const header = data[0].map(h => _eleves_up(h));
    
    // Identifier les colonnes
    let colClasse = -1;
    let colEffectif = -1;
    let colOptions = -1;
    
    for (let i = 0; i < header.length; i++) {
      if (header[i].includes('CLASSE') && header[i].includes('DEST')) colClasse = i;
      if (header[i].includes('EFFECTIF')) colEffectif = i;
      if (header[i].includes('OPTIONS')) colOptions = i;
    }
    
    if (colClasse === -1 || colEffectif === -1 || colOptions === -1) {
      return {success: false, error: "Colonnes requises non trouvées dans _STRUCTURE"};
    }
    
    // Mettre à jour les données existantes
    const classMap = {};
    for (let i = 1; i < data.length; i++) {
      const classe = _eleves_s(data[i][colClasse]);
      if (classe) classMap[classe] = i;
    }
    
    // Appliquer les nouvelles règles
    Object.keys(newRules).forEach(classe => {
      const rule = newRules[classe];
      const quotasStr = Object.entries(rule.quotas || {})
        .map(([k, v]) => `${k}=${v}`)
        .join(', ');
      
      if (classMap[classe]) {
        // Mettre à jour la ligne existante
        const row = classMap[classe];
        sh.getRange(row + 1, colEffectif + 1).setValue(rule.capacity);
        sh.getRange(row + 1, colOptions + 1).setValue(quotasStr);
      } else {
        // Ajouter une nouvelle ligne
        const newRow = sh.getLastRow() + 1;
        sh.getRange(newRow, colClasse + 1).setValue(classe);
        sh.getRange(newRow, colEffectif + 1).setValue(rule.capacity);
        sh.getRange(newRow, colOptions + 1).setValue(quotasStr);
      }
    });
    
    // Log de l'opération
    try {
      const timestamp = new Date();
      const user = Session.getActiveUser().getEmail();
      console.log(`Règles _STRUCTURE mises à jour par ${user} à ${timestamp}`);
    } catch (e) {
      // Ignorer si pas d'accès à Session
    }
    
    return {success: true, message: "Règles mises à jour avec succès"};
    
  } catch (e) {
    console.error('Erreur dans updateStructureRules:', e);
    return {success: false, error: e.toString()};
  }
}
/**
 * Sauvegarde automatique des données en mode CACHE
 * Appelée par le frontend toutes les 30 secondes
 * @param {Object} cacheData - {date, disposition, mode}
 * @return {Object} {success: true/false, message/error}
 */
function saveCacheData(cacheData) {
  try {
    console.log('📁 Sauvegarde automatique CACHE démarrée...');
    
    // Vérifier que nous avons bien les données nécessaires
    if (!cacheData || !cacheData.disposition) {
      return {
        success: false,
        error: "Données de cache invalides"
      };
    }
    
    // Utiliser la fonction existante saveElevesCache
    const result = saveElevesCache(cacheData.disposition);
    
    // Ajouter des métadonnées de sauvegarde
    if (result.success) {
      // Stocker la date et le mode dans les propriétés du document
      const props = PropertiesService.getDocumentProperties();
      props.setProperty('lastCacheDate', cacheData.date || new Date().toISOString());
      props.setProperty('lastCacheMode', cacheData.mode || 'CACHE');
      
      console.log('✅ Cache sauvegardé avec succès');
      return {
        success: true,
        message: "Cache sauvegardé automatiquement"
      };
    }
    
    return result;
    
  } catch (error) {
    console.error('❌ Erreur lors de la sauvegarde du cache:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Fonction pour récupérer les informations du dernier cache sauvegardé
 * Amélioration de getCacheInfo existante
 */
function getLastCacheInfo() {
  try {
    const props = PropertiesService.getDocumentProperties();
    const lastDate = props.getProperty('lastCacheDate');
    const lastMode = props.getProperty('lastCacheMode');
    
    // Utiliser aussi getCacheInfo existante pour double vérification
    const existingInfo = getCacheInfo();
    
    return {
      exists: existingInfo.exists || !!lastDate,
      date: lastDate || (existingInfo.exists ? existingInfo.date : null),
      mode: lastMode || 'CACHE'
    };
  } catch (error) {
    console.error('Erreur getLastCacheInfo:', error);
    return { exists: false };
  }
}