/**
 * Affiche les statistiques dans une interface utilisateur propre
 */
function afficherStatistiquesDef() {
  try {
    // Créer et afficher l'interface HTML
    const html = HtmlService.createHtmlOutputFromFile('StatistiquesDashboard')
      .setWidth(900)
      .setHeight(600)
      .setTitle('Statistiques Classes Phase 5');
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Statistiques Classes Phase 5');
    
    return { success: true };
  } catch (e) {
    Logger.log("Erreur lors de l'affichage des statistiques DEF: " + e.message);
    Browser.msgBox("Erreur", "Impossible d'afficher les statistiques : " + e.message, Browser.Buttons.OK);
    return { success: false, error: e.message };
  }
}

/**
 * Fournit les données au tableau de bord HTML
 * 1) Si TEMP_STATS_DATA existe     → on la renvoie, puis on la vide
 * 2) Sinon                         → on calcule à la volée (DEF par défaut)
 */
function getStatsData() {

  /* ---------- cas 1 : données déjà prêtes ------------------------- */
  const props = PropertiesService.getScriptProperties();
  const cache = props.getProperty('TEMP_STATS_DATA');
  if (cache) {
    props.deleteProperty('TEMP_STATS_DATA');   // on nettoie
    return JSON.parse(cache);                  // et on sert
  }

  /* ---------- cas 2 : calcul “à la demande” ----------------------- */
  const view = props.getProperty('currentView') || 'def';

  // sélecteurs d’onglets par vue
  const selector = {
    def    : getDefSheets,
    test   : getTestSheets,
    sources: getSourceSheets
  };

  if (view === 'compare') {
    return {
      dataDef : calculerStatistiquesFeuilles(getDefSheets()),
      dataTest: calculerStatistiquesFeuilles(getTestSheets()),
      params  : { viewType: 'compare',
                  title   : 'Comparaison TEST / DEF' }
    };
  }

  const sheets = (selector[view] || getDefSheets)();
  if (!sheets.length)
    return { error: `Aucun onglet ${view.toUpperCase()} trouvé.` };

  return {
    data  : calculerStatistiquesFeuilles(sheets),
    params: { viewType: view,
              title   : `Statistiques ${view.toUpperCase()}` }
  };
}
// Ajoutez cette fonction pour récupérer les feuilles TEST
function getTestSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  return sheets.filter(sheet => {
    const name = sheet.getName();
    return name.includes('TEST') || name.includes('TST');
  });
}


/**
 * Calcule les statistiques pour un ensemble de feuilles
 * @param {Sheet[]} sheets - Liste des feuilles
 * @return {Array} Données statistiques calculées
 */
function calculerStatistiquesFeuilles(sheets) {
  const config = getConfig();
  const statsData = [];
  
  for (const sheet of sheets) {
    try {
      const sheetName = sheet.getName();
      const range = sheet.getDataRange();
      const values = range.getValues();
      
      if (values.length <= 1) continue; // Ignorer les feuilles vides
      
      const headers = values[0].map(h => String(h || "").trim().toUpperCase());
      
      // Rechercher les indices des colonnes importantes en utilisant des alias
      let colIndexNomPrenom = -1;
      let colIndexSexe = -1;
      let colIndexCOM = -1;
      let colIndexTRA = -1;
      let colIndexPART = -1;
      let colIndexABS = -1;
      
      // Essayer d'utiliser findColumnIndex de Utils.gs si disponible
      try {
        if (typeof findColumnIndex === 'function') {
          colIndexNomPrenom = findColumnIndex(headers, config.COLUMN_ALIASES.NOM_PRENOM);
          colIndexSexe = findColumnIndex(headers, config.COLUMN_ALIASES.SEXE);
          colIndexCOM = findColumnIndex(headers, config.COLUMN_ALIASES.COM);
          colIndexTRA = findColumnIndex(headers, config.COLUMN_ALIASES.TRA);
          colIndexPART = findColumnIndex(headers, config.COLUMN_ALIASES.PART);
          colIndexABS = findColumnIndex(headers, config.COLUMN_ALIASES.ABS);
        }
      } catch (e) {
        // Fallback manuel si findColumnIndex n'est pas disponible ou échoue
        for (let i = 0; i < headers.length; i++) {
          const header = headers[i];
          
          if ((config.COLUMN_ALIASES.NOM_PRENOM || []).some(alias => header.includes(alias))) {
            colIndexNomPrenom = i;
          } else if ((config.COLUMN_ALIASES.SEXE || []).some(alias => header.includes(alias))) {
            colIndexSexe = i;
          } else if ((config.COLUMN_ALIASES.COM || []).some(alias => header.includes(alias))) {
            colIndexCOM = i;
          } else if ((config.COLUMN_ALIASES.TRA || []).some(alias => header.includes(alias))) {
            colIndexTRA = i;
          } else if ((config.COLUMN_ALIASES.PART || []).some(alias => header.includes(alias))) {
            colIndexPART = i;
          } else if ((config.COLUMN_ALIASES.ABS || []).some(alias => header.includes(alias))) {
            colIndexABS = i;
          }
        }
      }
      
      // S'il manque des colonnes essentielles, essayer un fallback plus simple
      if (colIndexNomPrenom === -1 || colIndexSexe === -1) {
        for (let i = 0; i < headers.length; i++) {
          const header = headers[i];
          
          if (header.includes("NOM")) {
            colIndexNomPrenom = i;
          } else if (header === "SEXE" || header === "S") {
            colIndexSexe = i;
          } else if (header === "COM" || header === "H") {
            colIndexCOM = i;
          } else if (header === "TRA" || header === "I") {
            colIndexTRA = i;
          } else if (header === "PART" || header === "J") {
            colIndexPART = i;
          } else if (header === "ABS" || header === "K") {
            colIndexABS = i;
          }
        }
      }
      
      // Vérifier que les colonnes nécessaires existent
      if (colIndexNomPrenom === -1 || colIndexSexe === -1) {
        Logger.log(`WARN: Colonnes requises manquantes dans ${sheetName}`);
        continue;
      }
      
      // Compter les filles et garçons
      let femaleCount = 0;
      let maleCount = 0;
      
      // Initialiser les compteurs de scores
      const scores = {
        com: [0, 0, 0, 0],  // Index 0=score 1, index 1=score 2, etc.
        tra: [0, 0, 0, 0],
        part: [0, 0, 0, 0],
        abs: [0, 0, 0, 0]
      };
      
      // Analyser les données des élèves
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const nomPrenom = row[colIndexNomPrenom];
        
        // Ignorer les lignes vides
        if (!nomPrenom) continue;
        
        // Compter par sexe
        const sexe = String(row[colIndexSexe] || "").trim().toUpperCase();
        if (sexe === "F") {
          femaleCount++;
        } else if (sexe === "M" || sexe === "G") {
          maleCount++;
        }
        
        // Compter les scores pour chaque critère
        if (colIndexCOM !== -1) {
          const comValue = parseInt(row[colIndexCOM]);
          if (comValue >= 1 && comValue <= 4) {
            scores.com[comValue - 1]++;
          }
        }
        
        if (colIndexTRA !== -1) {
          const traValue = parseInt(row[colIndexTRA]);
          if (traValue >= 1 && traValue <= 4) {
            scores.tra[traValue - 1]++;
          }
        }
        
        if (colIndexPART !== -1) {
          const partValue = parseInt(row[colIndexPART]);
          if (partValue >= 1 && partValue <= 4) {
            scores.part[partValue - 1]++;
          }
        }
        
        if (colIndexABS !== -1) {
          const absValue = parseInt(row[colIndexABS]);
          if (absValue >= 1 && absValue <= 4) {
            scores.abs[absValue - 1]++;
          }
        }
      }
      
      // Ajouter les données calculées
      statsData.push({
        class: sheetName,
        female: femaleCount,
        male: maleCount,
        total: femaleCount + maleCount,
        scores: scores
      });
      
    } catch (e) {
      Logger.log(`Erreur dans le calcul pour la feuille ${sheet.getName()}: ${e.message}`);
    }
  }
  
  return statsData;
}

function comparerStatistiques() {
  try {
    PropertiesService.getScriptProperties().setProperty('currentView', 'compare');
    return { success: true };
  } catch (e) {
    Logger.log("Erreur lors de la comparaison des statistiques: " + e.message);
    return { success: false, error: e.message };
  }
}
/************************************************************
 *  Statistiques.gs   —  LIGHT VERSION
 *  Pas d’ouverture de fenêtre, pas de calculs.
 *  Sert UNIQUEMENT à indiquer à la modale quel jeu de données
 *  afficher quand on clique sur « Classes Sources » DANS la
 *  fenêtre déjà ouverte.
 ************************************************************/
function afficherStatistiquesSources() {
  try {
    PropertiesService.getScriptProperties()
      .setProperty('currentView', 'sources');
    return { success: true };
  } catch (e) {
    Logger.log('Erreur afficherStatistiquesSources (light) : ' + e.message);
    return { success: false, error: e.message };
  }
}