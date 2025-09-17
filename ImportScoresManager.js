/**
 * ImportScoresManager.js
 * Version corrigée pour onglet GLOBAL → onglets INT
 * 
 * PASSE 1 : Mapping NOM+PRENOM et attribution des ID dans colonne P de GLOBAL
 * PASSE 2 : Lecture scores M/N de GLOBAL et écriture U/V dans INT
 */

const ImportScoresManager = (function() {
  'use strict';

  // ========== CONFIGURATION CORRIGÉE ==========
  const CONFIG = {
    // Colonnes dans l'onglet GLOBAL
    GLOBAL_COLUMNS: {
      NOM: 'B',           // Colonne B
      PRENOM: 'C',        // Colonne C
      SCORE_F: 'M',       // Colonne M = NIVEAU FRANCAIS
      SCORE_M: 'N',       // Colonne N = NIVEAU MATHS
      ID_DEST: 'P'        // Colonne P pour écrire les ID trouvés
    },
    
    // Colonnes dans les onglets INT
    INT_COLUMNS: {
      ID_ELEVE: 'A',      // Colonne A = ID_ELEVE
      NOM_PRENOM: 'D',    // Colonne D = NOM & PRENOM
      SCORE_F: 'U',       // Colonne U pour Score F
      SCORE_M: 'V'        // Colonne V pour Score M
    },
    
    // Nom de l'onglet source
    GLOBAL_SHEET: 'GLOBAL',
    
    // Paramètres de correspondance
    MATCHING: {
      TOLERANCE_NOM: 0.8,
      TOLERANCE_PRENOM: 0.7,
      MIN_SCORE: 0,
      MAX_SCORE: 20
    }
  };

  // ========== ÉTAT GLOBAL ==========
  let state = {
    globalData: null,
    intData: null,
    mapping: {},
    errors: [],
    warnings: [],
    stats: {
      totalGlobal: 0,
      totalInt: 0,
      mapped: 0,
      errors: 0,
      warnings: 0
    }
  };

  // ========== UTILITAIRES ==========
  function normalizeText(text) {
    return String(text || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9\s]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function calculateSimilarity(str1, str2) {
    const s1 = normalizeText(str1);
    const s2 = normalizeText(str2);
    
    if (s1 === s2) return 1.0;
    if (s1.length === 0 || s2.length === 0) return 0.0;
    
    const longer = s1.length > s2.length ? s1 : s2;
    const shorter = s1.length > s2.length ? s2 : s1;
    
    if (longer.length === 0) return 1.0;
    
    const distance = levenshteinDistance(longer, shorter);
    return (longer.length - distance) / longer.length;
  }

  function levenshteinDistance(str1, str2) {
    const matrix = [];
    
    for (let i = 0; i <= str2.length; i++) {
      matrix[i] = [i];
    }
    
    for (let j = 0; j <= str1.length; j++) {
      matrix[0][j] = j;
    }
    
    for (let i = 1; i <= str2.length; i++) {
      for (let j = 1; j <= str1.length; j++) {
        if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1,
            matrix[i][j - 1] + 1,
            matrix[i - 1][j] + 1
          );
        }
      }
    }
    
    return matrix[str2.length][str1.length];
  }

  function validateScore(score) {
    const numScore = parseFloat(score);
    if (isNaN(numScore)) return null;
    if (numScore < CONFIG.MATCHING.MIN_SCORE || numScore > CONFIG.MATCHING.MAX_SCORE) return null;
    return numScore;
  }

  // ========== PASSE 1 : MAPPING ET ATTRIBUTION D'ID ==========
  function loadGlobalData() {
    try {
      console.log('📁 Chargement de l\'onglet GLOBAL...');
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const globalSheet = ss.getSheetByName(CONFIG.GLOBAL_SHEET);
      
      if (!globalSheet) {
        throw new Error(`Onglet '${CONFIG.GLOBAL_SHEET}' non trouvé`);
      }
      
      const data = globalSheet.getDataRange().getValues();
      if (data.length < 2) {
        throw new Error('Onglet GLOBAL vide ou sans données');
      }
      
      // Extraire les données (ligne 1 = headers, ligne 2+ = données)
      const students = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const nom = String(row[1] || '').trim();     // Colonne B
        const prenom = String(row[2] || '').trim();  // Colonne C
        
        if (nom) { // Au minimum un nom requis
          students.push({
            rowIndex: i + 1, // +1 pour Google Sheets (base 1)
            nom: nom,
            prenom: prenom,
            scoreF: row[12] || '', // Colonne M
            scoreM: row[13] || ''  // Colonne N
          });
        }
      }
      
      state.globalData = {
        sheet: globalSheet,
        students: students
      };
      
      state.stats.totalGlobal = students.length;
      console.log(`✅ Onglet GLOBAL chargé : ${students.length} élèves`);
      
      return true;
    } catch (error) {
      console.error('❌ Erreur chargement GLOBAL:', error);
      state.errors.push(`Erreur chargement GLOBAL: ${error.message}`);
      return false;
    }
  }

  function loadIntData() {
    try {
      console.log('📊 Chargement des onglets INT...');
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = ss.getSheets();
      const intSheets = sheets.filter(sheet => sheet.getName().endsWith('INT'));
      
      if (intSheets.length === 0) {
        throw new Error('Aucun onglet INT trouvé');
      }
      
      const intData = {};
      let totalStudents = 0;
      
      intSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        const data = sheet.getDataRange().getValues();
        
        if (data.length < 2) return;
        
        const students = [];
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const id = String(row[0] || '').trim();      // Colonne A = ID_ELEVE
          const nomPrenom = String(row[3] || '').trim(); // Colonne D = NOM & PRENOM
          
          if (id && nomPrenom) {
            students.push({
              rowIndex: i + 1,
              id: id,
              nomPrenom: nomPrenom
            });
          }
        }
        
        intData[sheetName] = {
          sheet: sheet,
          students: students
        };
        
        totalStudents += students.length;
      });
      
      state.intData = intData;
      state.stats.totalInt = totalStudents;
      console.log(`✅ Onglets INT chargés : ${intSheets.length} onglets, ${totalStudents} élèves`);
      
      return true;
    } catch (error) {
      console.error('❌ Erreur chargement INT:', error);
      state.errors.push(`Erreur chargement INT: ${error.message}`);
      return false;
    }
  }

  function createMappingAndWriteIds() {
    try {
      console.log('🔗 Création du mapping et écriture des ID...');
      
      const mapping = {};
      const unmatched = [];
      let written = 0;
      
      // Pour chaque élève de GLOBAL
      state.globalData.students.forEach(globalStudent => {
        const globalNom = globalStudent.nom;
        const globalPrenom = globalStudent.prenom;
        const globalFullName = `${globalNom} ${globalPrenom}`.trim();
        
        // Chercher dans tous les onglets INT
        let bestMatch = null;
        let bestSimilarity = 0;
        
        Object.keys(state.intData).forEach(sheetName => {
          state.intData[sheetName].students.forEach(intStudent => {
            // Comparer avec le nom complet de l'INT
            const similarity = calculateSimilarity(globalFullName, intStudent.nomPrenom);
            
            if (similarity > bestSimilarity && similarity >= CONFIG.MATCHING.TOLERANCE_NOM) {
              bestSimilarity = similarity;
              bestMatch = {
                id: intStudent.id,
                sheet: sheetName,
                nomPrenom: intStudent.nomPrenom,
                similarity: similarity
              };
            }
          });
        });
        
        if (bestMatch) {
          // ÉCRIRE L'ID dans la colonne P de GLOBAL
          try {
            state.globalData.sheet.getRange(globalStudent.rowIndex, 16).setValue(bestMatch.id); // Colonne P = 16
            written++;
            
            mapping[globalStudent.rowIndex] = bestMatch;
            state.stats.mapped++;
            
            if (bestMatch.similarity < 1.0) {
              state.warnings.push({
                type: 'correspondance_partielle',
                global: { nom: globalNom, prenom: globalPrenom },
                int: bestMatch,
                similarity: bestMatch.similarity
              });
            }
          } catch (writeError) {
            console.error(`Erreur écriture ID pour ${globalFullName}:`, writeError);
            unmatched.push({
              ligne: globalStudent.rowIndex,
              raison: 'Erreur écriture ID',
              data: { nom: globalNom, prenom: globalPrenom }
            });
          }
        } else {
          unmatched.push({
            ligne: globalStudent.rowIndex,
            raison: 'Aucune correspondance trouvée',
            data: { nom: globalNom, prenom: globalPrenom }
          });
        }
      });
      
      state.mapping = mapping;
      state.stats.errors = unmatched.length;
      
      console.log(`✅ Mapping créé : ${state.stats.mapped} correspondances, ${written} ID écrits, ${unmatched.length} non trouvés`);
      
      // Afficher les résultats
      displayMappingResults(unmatched);
      
      return unmatched.length === 0;
    } catch (error) {
      console.error('❌ Erreur création mapping:', error);
      state.errors.push(`Erreur création mapping: ${error.message}`);
      return false;
    }
  }

  // ========== PASSE 2 : ÉCRITURE DES SCORES ==========
  function writeScoresToInt() {
    try {
      console.log('✍️ Écriture des scores dans les onglets INT...');
      
      let totalWritten = 0;
      let totalErrors = 0;
      
      // Relire l'onglet GLOBAL pour récupérer les ID écrits
      const data = state.globalData.sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const id = String(row[15] || '').trim(); // Colonne P = index 15
        const scoreF = row[12]; // Colonne M
        const scoreM = row[13]; // Colonne N
        
        if (!id) continue; // Pas d'ID mappé
        
        // Valider les scores
        const validScoreF = validateScore(scoreF);
        const validScoreM = validateScore(scoreM);
        
        if (validScoreF === null && validScoreM === null) continue;
        
        // Trouver l'onglet INT contenant cet ID
        let targetSheet = null;
        let targetRow = -1;
        
        Object.keys(state.intData).forEach(sheetName => {
          state.intData[sheetName].students.forEach(student => {
            if (student.id === id) {
              targetSheet = state.intData[sheetName].sheet;
              targetRow = student.rowIndex;
            }
          });
        });
        
        if (targetSheet && targetRow > 0) {
          try {
            // Écrire les scores
            if (validScoreF !== null) {
              targetSheet.getRange(targetRow, 21).setValue(validScoreF); // Colonne U = 21
            }
            if (validScoreM !== null) {
              targetSheet.getRange(targetRow, 22).setValue(validScoreM); // Colonne V = 22
            }
            
            totalWritten++;
          } catch (error) {
            console.error(`Erreur écriture scores pour ID ${id}:`, error);
            totalErrors++;
          }
        } else {
          totalErrors++;
        }
      }
      
      console.log(`✅ Écriture terminée : ${totalWritten} scores écrits, ${totalErrors} erreurs`);
      
      // Afficher le rapport final
      displayFinalReport(totalWritten, totalErrors);
      
      return totalErrors === 0;
    } catch (error) {
      console.error('❌ Erreur écriture scores:', error);
      state.errors.push(`Erreur écriture scores: ${error.message}`);
      return false;
    }
  }

  // ========== AFFICHAGE DES RÉSULTATS ==========
  function displayMappingResults(unmatched) {
    const modal = document.createElement('div');
    modal.className = 'fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50';
    
    let content = `
      <div class="bg-white rounded-lg shadow-xl p-6 max-w-4xl w-full max-h-[90vh] overflow-y-auto">
        <div class="flex justify-between items-center mb-4">
          <h3 class="text-lg font-bold">
            <i class="fas fa-link text-blue-600 mr-2"></i>
            Résultats du mapping (Passe 1) - GLOBAL → INT
          </h3>
          <button onclick="this.closest('.fixed').remove()" class="text-gray-400 hover:text-gray-600">
            <i class="fas fa-times text-xl"></i>
          </button>
        </div>
        
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
          <div class="bg-green-50 p-4 rounded-lg">
            <div class="text-2xl font-bold text-green-600">${state.stats.mapped}</div>
            <div class="text-sm text-green-700">ID écrits dans GLOBAL</div>
          </div>
          <div class="bg-red-50 p-4 rounded-lg">
            <div class="text-2xl font-bold text-red-600">${unmatched.length}</div>
            <div class="text-sm text-red-700">Non trouvés</div>
          </div>
          <div class="bg-blue-50 p-4 rounded-lg">
            <div class="text-2xl font-bold text-blue-600">${state.warnings.length}</div>
            <div class="text-sm text-blue-700">Avertissements</div>
          </div>
        </div>
        
        <div class="bg-gray-50 p-4 rounded-lg mb-6">
          <h4 class="font-bold text-gray-800 mb-2">Processus :</h4>
          <div class="text-sm text-gray-600 space-y-1">
            <div>1. 📁 Lecture onglet GLOBAL (${state.stats.totalGlobal} élèves)</div>
            <div>2. 📊 Lecture onglets INT (${state.stats.totalInt} élèves)</div>
            <div>3. 🔗 Mapping NOM+PRENOM GLOBAL → INT</div>
            <div>4. ✍️ Écriture ID dans colonne P de GLOBAL</div>
          </div>
        </div>
    `;
    
    if (unmatched.length > 0) {
      content += `
        <div class="mb-6">
          <h4 class="font-bold text-red-600 mb-3">
            <i class="fas fa-exclamation-triangle mr-2"></i>
            Élèves non trouvés (${unmatched.length})
          </h4>
          <div class="bg-red-50 p-4 rounded-lg max-h-60 overflow-y-auto">
            <div class="text-sm space-y-2">
      `;
      
      unmatched.slice(0, 20).forEach(item => {
        content += `
          <div class="flex justify-between items-center">
            <span>Ligne ${item.ligne}: ${item.data.nom} ${item.data.prenom}</span>
            <span class="text-red-500 text-xs">${item.raison}</span>
          </div>
        `;
      });
      
      if (unmatched.length > 20) {
        content += `<div class="text-gray-500 text-xs">... et ${unmatched.length - 20} autres</div>`;
      }
      
      content += `
            </div>
          </div>
        </div>
      `;
    }
    
    if (state.warnings.length > 0) {
      content += `
        <div class="mb-6">
          <h4 class="font-bold text-orange-600 mb-3">
            <i class="fas fa-exclamation-circle mr-2"></i>
            Correspondances partielles (${state.warnings.length})
          </h4>
          <div class="bg-orange-50 p-4 rounded-lg max-h-60 overflow-y-auto">
            <div class="text-sm space-y-2">
      `;
      
      state.warnings.slice(0, 10).forEach(warning => {
        content += `
          <div class="flex justify-between items-center">
            <span>${warning.global.nom} ${warning.global.prenom} → ${warning.int.nomPrenom}</span>
            <span class="text-orange-500 text-xs">${Math.round(warning.similarity * 100)}%</span>
          </div>
        `;
      });
      
      if (state.warnings.length > 10) {
        content += `<div class="text-gray-500 text-xs">... et ${state.warnings.length - 10} autres</div>`;
      }
      
      content += `
            </div>
          </div>
        </div>
      `;
    }
    
    content += `
        <div class="flex justify-end gap-3">
          <button onclick="this.closest('.fixed').remove()" class="btn btn-secondary">
            Annuler
          </button>
          <button onclick="ImportScoresManager.proceedToPhase2()" class="btn btn-primary" ${unmatched.length > 0 ? 'disabled' : ''}>
            <i class="fas fa-arrow-right mr-2"></i>
            Continuer vers Passe 2 (Écriture scores)
          </button>
        </div>
      </div>
    `;
    
    modal.innerHTML = content;
    document.body.appendChild(modal);
  }

  function displayFinalReport(written, errors) {
    const modal = document.createElement('div');
    modal.className = 'fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50';
    
    const success = errors === 0;
    
    modal.innerHTML = `
      <div class="bg-white rounded-lg shadow-xl p-6 max-w-2xl w-full">
        <div class="text-center">
          <div class="text-6xl mb-4">
            ${success ? '✅' : '⚠️'}
          </div>
          <h3 class="text-xl font-bold mb-4">
            ${success ? 'Import terminé avec succès' : 'Import terminé avec des erreurs'}
          </h3>
          
          <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
            <div class="bg-green-50 p-4 rounded-lg">
              <div class="text-2xl font-bold text-green-600">${written}</div>
              <div class="text-sm text-green-700">Scores écrits</div>
            </div>
            <div class="bg-red-50 p-4 rounded-lg">
              <div class="text-2xl font-bold text-red-600">${errors}</div>
              <div class="text-sm text-red-700">Erreurs</div>
            </div>
            <div class="bg-blue-50 p-4 rounded-lg">
              <div class="text-2xl font-bold text-blue-600">${state.warnings.length}</div>
              <div class="text-sm text-blue-700">Avertissements</div>
            </div>
          </div>
          
          <div class="bg-gray-50 p-4 rounded-lg mb-6">
            <h4 class="font-bold text-gray-800 mb-2">Processus Passe 2 :</h4>
            <div class="text-sm text-gray-600 space-y-1">
              <div>1. 📖 Lecture ID dans colonne P de GLOBAL</div>
              <div>2. 📊 Lecture scores M (Français) et N (Maths) de GLOBAL</div>
              <div>3. ✍️ Écriture dans onglets INT : U (Français) et V (Maths)</div>
            </div>
          </div>
          
          <button onclick="this.closest('.fixed').remove()" class="btn btn-primary">
            <i class="fas fa-check mr-2"></i>
            Fermer
          </button>
        </div>
      </div>
    `;
    
    document.body.appendChild(modal);
  }

  // ========== FONCTIONS PUBLIQUES ==========
  function startImport() {
    console.log('🚀 Début de l\'import GLOBAL → INT...');
    
    // Reset de l'état
    state = {
      globalData: null,
      intData: null,
      mapping: {},
      errors: [],
      warnings: [],
      stats: { totalGlobal: 0, totalInt: 0, mapped: 0, errors: 0, warnings: 0 }
    };
    
    // PASSE 1 : Mapping et écriture des ID
    if (!loadGlobalData()) return false;
    if (!loadIntData()) return false;
    if (!createMappingAndWriteIds()) return false;
    
    return true;
  }

  function proceedToPhase2() {
    console.log('🔄 Passage à la phase 2...');
    
    // Fermer le modal de mapping
    const modal = document.querySelector('.fixed');
    if (modal) modal.remove();
    
    // PASSE 2 : Écriture des scores
    writeScoresToInt();
  }

  function generateReport() {
    return {
      success: state.errors.length === 0,
      stats: state.stats,
      errors: state.errors,
      warnings: state.warnings,
      mapping: Object.keys(state.mapping).length
    };
  }

  // ========== EXPORT PUBLIC ==========
  return {
    startImport: startImport,
    proceedToPhase2: proceedToPhase2,
    generateReport: generateReport,
    CONFIG: CONFIG
  };

})();

// ========== INTÉGRATION AVEC L'INTERFACE ==========

/**
 * Fonction appelée par le bouton "Import Scores"
 * Version Google Apps Script qui ne nécessite pas de fichier
 */
function handleImportScores() {
  // Lancer l'import depuis l'onglet GLOBAL
  const success = ImportScoresManager.startImport();
  
  if (!success) {
    console.error('❌ Erreur lors de l\'import');
    if (typeof toast === 'function') {
      toast('Erreur lors de l\'import. Vérifiez la console.', 'error');
    }
  }
}

// ========== EXPORT POUR GOOGLE APPS SCRIPT ==========
if (typeof global !== 'undefined') {
  global.ImportScoresManager = ImportScoresManager;
  global.handleImportScores = handleImportScores;
} 