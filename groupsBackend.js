// ========== CODE SERVEUR GOOGLE APPS SCRIPT POUR LES GROUPES ==========

// Fonction pour obtenir l'interface HTML des groupes
function getGroupsInterface() {
  try {
    // Charger le template HTML principal
    const template = HtmlService.createTemplateFromFile('groupsInterface');
    
    // Évaluer et retourner le HTML
    return template.evaluate().getContent();
  } catch (error) {
    console.error('Erreur getGroupsInterface:', error);
    throw new Error('Impossible de charger l\'interface des groupes');
  }
}

// Fonction include pour inclure d'autres fichiers HTML
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    console.error('Erreur include:', error);
    return '';
  }
}

// Obtenir le nombre de groupes
function getGroupsCount() {
  try {
    const cache = CacheService.getScriptCache();
    const groupsData = cache.get('GROUPS_DATA');
    
    if (groupsData) {
      const groups = JSON.parse(groupsData);
      return Object.keys(groups).length;
    }
    
    // Si pas de cache, vérifier dans la feuille
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const groupsSheet = ss.getSheetByName('GROUPES');
    
    if (!groupsSheet) {
      return 0;
    }
    
    // Compter les groupes non vides
    const data = groupsSheet.getDataRange().getValues();
    const headers = data[0];
    let count = 0;
    
    for (let i = 1; i < headers.length; i++) {
      if (headers[i] && headers[i].toString().trim() !== '') {
        count++;
      }
    }
    
    return count;
  } catch (error) {
    console.error('Erreur getGroupsCount:', error);
    return 0;
  }
}

// Sauvegarder les groupes
function saveGroups(groupsData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let groupsSheet = ss.getSheetByName('GROUPES');
    
    // Créer la feuille si elle n'existe pas
    if (!groupsSheet) {
      groupsSheet = ss.insertSheet('GROUPES');
      groupsSheet.setTabColor('#9333ea'); // Couleur violette
    }
    
    // Effacer le contenu existant
    groupsSheet.clear();
    
    // Préparer les données pour l'écriture
    const allStudents = [];
    const headers = ['ID', 'Nom', 'Prénom', 'Sexe', 'Classe'];
    
    // Ajouter les colonnes de groupes
    groupsData.groups.forEach(group => {
      headers.push(group.name);
    });
    
    // Collecter tous les élèves uniques
    const studentMap = new Map();
    
    groupsData.groups.forEach((group, groupIndex) => {
      group.students.forEach(studentId => {
        if (!studentMap.has(studentId)) {
          studentMap.set(studentId, {
            id: studentId,
            groups: new Array(groupsData.groups.length).fill('')
          });
        }
        studentMap.get(studentId).groups[groupIndex] = 'X';
      });
    });
    
    // Récupérer les infos des élèves depuis ELEVES
    const elevesSheet = ss.getSheetByName('ELEVES');
    if (elevesSheet) {
      const elevesData = elevesSheet.getDataRange().getValues();
      const elevesHeaders = elevesData[0];
      
      const idCol = elevesHeaders.indexOf('ID');
      const nomCol = elevesHeaders.indexOf('NOM');
      const prenomCol = elevesHeaders.indexOf('PRENOM');
      const sexeCol = elevesHeaders.indexOf('SEXE');
      const classeCol = elevesHeaders.indexOf('CLASSE');
      
      for (let i = 1; i < elevesData.length; i++) {
        const id = elevesData[i][idCol];
        if (studentMap.has(id)) {
          const student = studentMap.get(id);
          student.nom = elevesData[i][nomCol] || '';
          student.prenom = elevesData[i][prenomCol] || '';
          student.sexe = elevesData[i][sexeCol] || '';
          student.classe = elevesData[i][classeCol] || '';
        }
      }
    }
    
    // Construire le tableau de données
    const dataRows = [headers];
    
    studentMap.forEach((student, id) => {
      const row = [
        id,
        student.nom || '',
        student.prenom || '',
        student.sexe || '',
        student.classe || ''
      ];
      row.push(...student.groups);
      dataRows.push(row);
    });
    
    // Écrire les données
    if (dataRows.length > 1) {
      groupsSheet.getRange(1, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
      
      // Formater l'en-tête
      const headerRange = groupsSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#f3f4f6');
      headerRange.setFontWeight('bold');
      
      // Ajuster les colonnes
      groupsSheet.autoResizeColumns(1, headers.length);
      
      // Figer la première ligne et les 5 premières colonnes
      groupsSheet.setFrozenRows(1);
      groupsSheet.setFrozenColumns(5);
    }
    
    // Sauvegarder aussi dans le cache
    const cache = CacheService.getScriptCache();
    cache.put('GROUPS_DATA', JSON.stringify(groupsData), 3600); // 1 heure
    
    // Ajouter les métadonnées
    const metaSheet = ss.getSheetByName('GROUPES_META') || ss.insertSheet('GROUPES_META');
    metaSheet.clear();
    metaSheet.getRange(1, 1, 1, 4).setValues([['Type', 'Date', 'Config', 'Timestamp']]);
    metaSheet.getRange(2, 1, 1, 4).setValues([[
      groupsData.type,
      new Date().toLocaleDateString('fr-FR'),
      JSON.stringify(groupsData.config),
      groupsData.timestamp
    ]]);
    metaSheet.hideSheet();
    
    return true;
  } catch (error) {
    console.error('Erreur saveGroups:', error);
    throw new Error('Erreur lors de la sauvegarde des groupes: ' + error.message);
  }
}

// Obtenir tous les groupes
function getAllGroups() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const groupsSheet = ss.getSheetByName('GROUPES');
    
    if (!groupsSheet) {
      return {};
    }
    
    const data = groupsSheet.getDataRange().getValues();
    if (data.length < 2) {
      return {};
    }
    
    const headers = data[0];
    const groups = {};
    
    // Parser les groupes (colonnes après les 5 premières)
    for (let col = 5; col < headers.length; col++) {
      const groupName = headers[col];
      if (!groupName) continue;
      
      groups[groupName] = {
        name: groupName,
        students: [],
        type: 'unknown'
      };
      
      // Collecter les élèves du groupe
      for (let row = 1; row < data.length; row++) {
        if (data[row][col] === 'X') {
          groups[groupName].students.push({
            id: data[row][0],
            nom: data[row][1],
            prenom: data[row][2],
            sexe: data[row][3],
            classe: data[row][4]
          });
        }
      }
    }
    
    // Récupérer le type depuis les métadonnées
    const metaSheet = ss.getSheetByName('GROUPES_META');
    if (metaSheet) {
      const metaData = metaSheet.getRange(2, 1, 1, 1).getValue();
      Object.values(groups).forEach(group => {
        group.type = metaData || 'unknown';
      });
    }
    
    return groups;
  } catch (error) {
    console.error('Erreur getAllGroups:', error);
    return {};
  }
}

// Supprimer tous les groupes
function deleteAllGroups() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Supprimer la feuille GROUPES
    const groupsSheet = ss.getSheetByName('GROUPES');
    if (groupsSheet) {
      ss.deleteSheet(groupsSheet);
    }
    
    // Supprimer la feuille GROUPES_META
    const metaSheet = ss.getSheetByName('GROUPES_META');
    if (metaSheet) {
      ss.deleteSheet(metaSheet);
    }
    
    // Vider le cache
    const cache = CacheService.getScriptCache();
    cache.remove('GROUPS_DATA');
    
    return true;
  } catch (error) {
    console.error('Erreur deleteAllGroups:', error);
    throw new Error('Erreur lors de la suppression des groupes: ' + error.message);
  }
}


// Obtenir les classes disponibles
function getAvailableClasses() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classes = new Set();
    
    // Chercher dans les onglets INT
    const sheets = ss.getSheets();
    const intSheets = sheets.filter(sheet => sheet.getName().endsWith('INT'));
    
    if (intSheets.length > 0) {
      const latestSheet = intSheets[intSheets.length - 1];
      const data = latestSheet.getDataRange().getValues();
      
      if (data.length > 1) {
        const headers = data[0];
        const classeCol = headers.indexOf('CLASSE');
        
        for (let i = 1; i < data.length; i++) {
          const classe = data[i][classeCol];
          if (classe) {
            classes.add(classe.toString());
          }
        }
      }
    }
    
    // Si pas de classes trouvées, utiliser des valeurs par défaut
    if (classes.size === 0) {
      ['501', '502', '503', '504', '505', '506'].forEach(c => classes.add(c));
    }
    
    return Array.from(classes).sort();
  } catch (error) {
    console.error('Erreur getAvailableClasses:', error);
    return ['501', '502', '503', '504', '505', '506'];
  }
}

// Fonction pour créer automatiquement la structure de démonstration
function createGroupsDemoData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Créer quelques groupes de démonstration
    const demoGroups = {
      type: 'language',
      groups: [
        {
          name: 'ESP - Groupe 1',
          students: ['eleve1', 'eleve5', 'eleve9']
        },
        {
          name: 'ESP - Groupe 2',
          students: ['eleve2', 'eleve6', 'eleve10']
        },
        {
          name: 'ITA - Groupe 1',
          students: ['eleve3', 'eleve7', 'eleve11']
        },
        {
          name: 'ITA - Groupe 2',
          students: ['eleve4', 'eleve8', 'eleve12']
        }
      ],
      config: {
        language: 'ESP',
        classes: ['501', '502', '503']
      },
      timestamp: new Date().toISOString()
    };
    
    // Sauvegarder les groupes de démonstration
    saveGroups(demoGroups);
    
    return true;
  } catch (error) {
    console.error('Erreur createGroupsDemoData :', error);
    return false;
  }
}   // <-- FIN de createGroupsDemoData()


// ========== PATCHES SERVEUR POUR LA GESTION DES SCORES ==========


// 3. PATCH: Sauvegarder les scores dans les colonnes U et V
function saveScoresToSheet(className, scores) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = className + 'INT';
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Onglet ${sheetName} non trouvé`);
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Vérifier/ajouter les en-têtes pour les scores
    if (!headers[20] || headers[20] !== 'SCORE F') {
      sheet.getRange(1, 21).setValue('SCORE F');
    }
    if (!headers[21] || headers[21] !== 'SCORE M') {
      sheet.getRange(1, 22).setValue('SCORE M');
    }
    
    // Mettre à jour les scores pour chaque élève
    scores.forEach(score => {
      // Trouver l'élève par nom et prénom
      for (let i = 1; i < data.length; i++) {
        if (data[i][headers.indexOf('NOM')] === score.nom && 
            data[i][headers.indexOf('PRENOM')] === score.prenom) {
          // Mettre à jour les scores
          sheet.getRange(i + 1, 21).setValue(score.scoreF || '');
          sheet.getRange(i + 1, 22).setValue(score.scoreM || '');
          break;
        }
      }
    });
    
    // Formater les colonnes de scores
    const scoreRange = sheet.getRange(2, 21, sheet.getLastRow() - 1, 2);
    scoreRange.setHorizontalAlignment('center');
    scoreRange.setBackground('#f0f9ff'); // Bleu clair
    
    return true;
  } catch (error) {
    console.error('Erreur saveScoresToSheet:', error);
    throw error;
  }
}

// 4. PATCH: Importer les scores depuis un fichier CSV
function importScoresFromCSV(csvData) {
  try {
    const lines = csvData.split('\n');
    const headers = lines[0].split(',').map(h => h.trim());
    
    // Vérifier les colonnes requises
    const requiredColumns = ['Classe', 'Nom', 'Prénom', 'Score F', 'Score M'];
    const missingColumns = requiredColumns.filter(col => !headers.includes(col));
    
    if (missingColumns.length > 0) {
      throw new Error(`Colonnes manquantes : ${missingColumns.join(', ')}`);
    }
    
    // Parser les scores
    const scoresByClass = {};
    
    for (let i = 1; i < lines.length; i++) {
      const values = lines[i].split(',').map(v => v.trim());
      if (values.length < headers.length) continue;
      
      const row = {};
      headers.forEach((header, index) => {
        row[header] = values[index];
      });
      
      const className = row['Classe'];
      if (!scoresByClass[className]) {
        scoresByClass[className] = [];
      }
      
      scoresByClass[className].push({
        nom: row['Nom'],
        prenom: row['Prénom'],
        scoreF: parseInt(row['Score F']) || 0,
        scoreM: parseInt(row['Score M']) || 0
      });
    }
    
    // Sauvegarder les scores dans chaque classe
    let totalUpdated = 0;
    Object.entries(scoresByClass).forEach(([className, scores]) => {
      try {
        saveScoresToSheet(className, scores);
        totalUpdated += scores.length;
      } catch (error) {
        console.error(`Erreur pour la classe ${className}:`, error);
      }
    });
    
    return {
      success: true,
      message: `${totalUpdated} scores importés avec succès`,
      details: scoresByClass
    };
  } catch (error) {
    console.error('Erreur importScoresFromCSV:', error);
    throw error;
  }
}

// 5. PATCH: Générer un template CSV pour l'import des scores
function generateScoreTemplate() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const students = getStudentsForGroups();
    
    // Créer le CSV
    let csv = 'Classe,Nom,Prénom,Score F,Score M\n';
    
    students.forEach(student => {
      csv += `"${student.classe}","${student.nom}","${student.prenom}","",""\n`;
    });
    
    return csv;
  } catch (error) {
    console.error('Erreur generateScoreTemplate:', error);
    throw error;
  }
}

// 6. PATCH: Statistiques des scores par classe
function getScoreStatistics() {
  try {
    const students = getStudentsForGroups();
    const stats = {};
    
    students.forEach(student => {
      const className = student.classe;
      if (!stats[className]) {
        stats[className] = {
          total: 0,
          scoreF: {1: 0, 2: 0, 3: 0, 4: 0, 0: 0},
          scoreM: {1: 0, 2: 0, 3: 0, 4: 0, 0: 0},
          avgF: 0,
          avgM: 0
        };
      }
      
      stats[className].total++;
      stats[className].scoreF[student.scores.F || 0]++;
      stats[className].scoreM[student.scores.M || 0]++;
    });
    
    // Calculer les moyennes
    Object.keys(stats).forEach(className => {
      let sumF = 0, sumM = 0, countF = 0, countM = 0;
      
      [1, 2, 3, 4].forEach(score => {
        sumF += score * stats[className].scoreF[score];
        countF += stats[className].scoreF[score];
        sumM += score * stats[className].scoreM[score];
        countM += stats[className].scoreM[score];
      });
      
      stats[className].avgF = countF > 0 ? (sumF / countF).toFixed(2) : 0;
      stats[className].avgM = countM > 0 ? (sumM / countM).toFixed(2) : 0;
    });
    
    return stats;
  } catch (error) {
    console.error('Erreur getScoreStatistics:', error);
    return {};
  }
}

/* ====================================================================
   CLASSES INT DISPONIBLES POUR L’INTERFACE
   --------------------------------------------------------------------
   - getINTClasses()            → renvoie ["6°1", "6°2", …] (sans “INT”)
   - getINTClassesForInterface  → alias conservé pour le front‑end
   ==================================================================== */

/**
 * Parcourt le classeur et renvoie la liste des classes pour
 * lesquelles il existe un onglet nommé « …INT ».
 *
 * Exemple : un onglet “6°1INT” ⇒ la chaîne “6°1” sera renvoyée.
 */
function getINTClasses() {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const classes = ss.getSheets()
                      .map(s => s.getName())
                      .filter(n => n.endsWith('INT'))   // ne garder que les onglets INT
                      .map(n => n.replace(/INT$/, ''))  // retirer “INT” pour l’affichage
                      .sort();

    // Valeurs de secours si aucune classe trouvée
    return classes.length ? classes
                          : ['501', '502', '503', '504', '505', '506'];
  } catch (err) {
    console.error('Erreur getINTClasses :', err);
    // Valeurs par défaut en cas de problème
    return ['501', '502', '503', '504', '505', '506'];
  }
}

/**
 * Alias conservé pour les appels existants côté front‑end
 * (`google.script.run.getINTClassesForInterface()`).
 */
function getINTClassesForInterface() {
  return getINTClasses();
 }

 // ========== FONCTION DE DEBUG POUR LES CLASSES INT ==========
// À ajouter dans le code serveur Google Apps Script

function debugINTClasses() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    console.log('=== DEBUG CLASSES INT ===');
    console.log('Nombre total d\'onglets:', sheets.length);
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      console.log('Onglet:', name, '- Se termine par INT?', name.endsWith('INT'));
    });
    
    const intSheets = sheets.filter(sheet => sheet.getName().endsWith('INT'));
    console.log('Onglets INT trouvés:', intSheets.length);
    
    const classes = intSheets.map(sheet => sheet.getName().replace('INT', ''));
    console.log('Classes extraites:', classes);
    
    return {
      totalSheets: sheets.length,
      sheetNames: sheets.map(s => s.getName()),
      intSheets: intSheets.map(s => s.getName()),
      classes: classes
    };
  } catch (error) {
    console.error('Erreur debugINTClasses:', error);
    return { error: error.toString() };
  }
}

// Corriger getStudentsForGroups pour retourner les classes INT correctement
function getStudentsForGroups() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const students = [];
    
    // Chercher tous les onglets se terminant par INT
    const sheets = ss.getSheets();
    const intSheets = sheets.filter(sheet => sheet.getName().endsWith('INT'));
    
    console.log('Onglets INT trouvés:', intSheets.map(s => s.getName()));
    
    if (intSheets.length === 0) {
      console.log('Aucun onglet INT trouvé, recherche alternative...');
      
      // Alternative : chercher des onglets avec pattern de classe
      const classPattern = /^[0-9]°[0-9]/; // Pattern pour 6°1, 5°2, etc.
      const classSheets = sheets.filter(sheet => classPattern.test(sheet.getName()));
      
      if (classSheets.length > 0) {
        console.log('Onglets de classe trouvés:', classSheets.map(s => s.getName()));
        
        classSheets.forEach(sheet => {
          const className = sheet.getName();
          const data = sheet.getDataRange().getValues();
          
          if (data.length > 1) {
            const headers = data[0];
            
                      // Indices des colonnes
          const indices = {
            nom: headers.indexOf('NOM'),
            prenom: headers.indexOf('PRENOM'),
            sexe: headers.indexOf('SEXE'),
            lv2: headers.indexOf('LV2'),  // ← CORRECTION : lire LV2 au lieu de LV1
            com: headers.indexOf('COM'),
            tra: headers.indexOf('TRA'),
            part: headers.indexOf('PART'),
            scoreF: 20, // Colonne U
            scoreM: 21  // Colonne V
          };
            
            // Lire les élèves
            for (let i = 1; i < data.length; i++) {
              const row = data[i];
              if (row[indices.nom]) {
                              students.push({
                id: `${className}_${i}`,
                nom: row[indices.nom] || '',
                prenom: row[indices.prenom] || '',
                sexe: row[indices.sexe] || '',
                classe: nettoyerNomClasse(className),  // ← CORRECTION : nettoyer le nom de classe
                lv2: row[indices.lv2] || '',  // ← CORRECTION : assigner à lv2 au lieu de lv1
                com: parseFloat(row[indices.com]) || 0,
                tra: parseFloat(row[indices.tra]) || 0,
                part: parseFloat(row[indices.part]) || 0,
                scores: {
                  F: parseInt(row[indices.scoreF]) || 0,
                  M: parseInt(row[indices.scoreM]) || 0
                }
              });
              }
            }
          }
        });
      }
    } else {
      // Parcourir chaque onglet INT
      intSheets.forEach(sheet => {
        const className = sheet.getName();
        const data = sheet.getDataRange().getValues();
        
        if (data.length > 1) {
          const headers = data[0];
          
          // Indices des colonnes
          const indices = {
            nom: headers.indexOf('NOM'),
            prenom: headers.indexOf('PRENOM'),
            sexe: headers.indexOf('SEXE'),
            lv2: headers.indexOf('LV2'),  // ← CORRECTION : lire LV2 au lieu de LV1
            com: headers.indexOf('COM'),
            tra: headers.indexOf('TRA'),
            part: headers.indexOf('PART'),
            scoreF: 20, // Colonne U
            scoreM: 21  // Colonne V
          };
          
          // Lire les élèves
          for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row[indices.nom]) {
              students.push({
                id: `${className}_${i}`,
                nom: row[indices.nom] || '',
                prenom: row[indices.prenom] || '',
                sexe: row[indices.sexe] || '',
                classe: nettoyerNomClasse(className.replace('INT', '')),  // ← CORRECTION : nettoyer le nom de classe
                lv2: row[indices.lv2] || '',  // ← CORRECTION : assigner à lv2 au lieu de lv1
                com: parseFloat(row[indices.com]) || 0,
                tra: parseFloat(row[indices.tra]) || 0,
                part: parseFloat(row[indices.part]) || 0,
                scores: {
                  F: parseInt(row[indices.scoreF]) || 0,
                  M: parseInt(row[indices.scoreM]) || 0
                }
              });
            }
          }
        }
      });
    }
    
    console.log('Total élèves trouvés:', students.length);
    return students;
  } catch (error) {
    console.error('Erreur getStudentsForGroups:', error);
    throw error;
  }
}

// ========== FONCTION DE NETTOYAGE DES NOMS DE CLASSES ==========
function nettoyerNomClasse(nomClasse) {
  if (!nomClasse) return '';
  
  let nomNettoye = nomClasse.toString()
    .replace(/Â°/g, '°')  // Corriger "6Â°4" → "6°4"
    .replace(/°/g, '°')   // Normaliser le caractère degré
    .replace(/\s+/g, ' ') // Normaliser les espaces
    .trim();
  
  // Log pour debug
  if (nomClasse !== nomNettoye) {
    console.log(`Classe corrigée: "${nomClasse}" → "${nomNettoye}"`);
  }
  
  return nomNettoye;
}

// ========== AJOUT : Fonction generateGroupsOnServer attendue par le frontend ========== //
function generateGroupsOnServer(params) {
  try {
    console.log('Génération des groupes avec params:', params);
    const students = getStudentsForGroups();
    
    // Filtrer les élèves selon les classes sélectionnées
    const filteredStudents = students.filter(student => 
      params.selectedClasses.includes(student.classe)
    );
    
    console.log(`📊 Élèves filtrés par classe (${params.selectedClasses.join(', ')}): ${filteredStudents.length}`);
    
    if (params.groupType === 'language') {
      // CORRECTION : Utiliser la même logique de filtrage que le frontend
      const langStudents = filteredStudents.filter(student => {
        // Normaliser la valeur de langue
        const langue = (student.lv2 || '').trim().toUpperCase();  // ← CORRECTION : utiliser lv2
        
        // Logique corrigée pour ESP/ITA
        if (params.selectedLanguage === 'ITA') {
          return langue === 'ITA';
        } else {
          // Pour ESP : inclure ESP explicite ET les valeurs vides (par défaut)
          return langue === 'ESP' || langue === '' || !langue;
        }
      });
      
      console.log(`📊 Élèves filtrés par langue (${params.selectedLanguage}): ${langStudents.length}`);
      console.log('📋 Élèves ESP/ITA trouvés:', langStudents.map(s => `${s.nom} ${s.prenom} (${s.lv2})`));  // ← CORRECTION : afficher lv2
      
      // Trier par score PART
      langStudents.sort((a, b) => (b.part || 0) - (a.part || 0));
      
      // Créer les groupes
      const groups = [];
      for (let i = 0; i < params.numGroups; i++) {
        groups.push({
          name: `${params.selectedLanguage} - Groupe ${i + 1}`,
          students: []
        });
      }
      
      // Répartir en serpentin
      langStudents.forEach((student, index) => {
        const groupIndex = index % params.numGroups;
        groups[groupIndex].students.push(student.id);
      });
      
      console.log('✅ Groupes créés:', groups.map(g => `${g.name}: ${g.students.length} élèves`));
      return { groups: groups };
    } else if (params.groupType === 'needs') {
      // Calculer les scores composites
      filteredStudents.forEach(student => {
        if (params.selectedSubject === 'Both') {
          student.compositeScore = ((student.scores?.M || 0) + (student.scores?.F || 0)) / 2;
        } else if (params.selectedSubject === 'Maths') {
          student.compositeScore = student.scores?.M || 0;
        } else {
          student.compositeScore = student.scores?.F || 0;
        }
      });
      // Trier par score
      filteredStudents.sort((a, b) => b.compositeScore - a.compositeScore);
      const groups = [];
      for (let i = 0; i < params.numLevelGroups; i++) {
        groups.push({
          name: `Niveau ${i + 1}`,
          students: []
        });
      }
      if (params.selectedDistributionType === 'homogeneous') {
        // Groupes homogènes
        const groupSize = Math.ceil(filteredStudents.length / params.numLevelGroups);
        filteredStudents.forEach((student, index) => {
          const groupIndex = Math.floor(index / groupSize);
          if (groupIndex < params.numLevelGroups) {
            groups[groupIndex].students.push(student.id);
          }
        });
      } else {
        // Groupes hétérogènes
        // Créer 4 niveaux de scores
        const scoreGroups = {1: [], 2: [], 3: [], 4: []};
        filteredStudents.forEach(student => {
          let scoreLevel;
          if (student.compositeScore === 0) {
            scoreLevel = 1;
          } else if (student.compositeScore <= 1) {
            scoreLevel = 1;
          } else if (student.compositeScore <= 2) {
            scoreLevel = 2;
          } else if (student.compositeScore <= 3) {
            scoreLevel = 3;
          } else {
            scoreLevel = 4;
          }
          scoreGroups[scoreLevel].push(student);
        });
        // Distribuer équitablement
        [1, 2, 3, 4].forEach(level => {
          scoreGroups[level].forEach((student, index) => {
            const groupIndex = index % params.numLevelGroups;
            groups[groupIndex].students.push(student.id);
          });
        });
      }
      return { groups: groups };
    }
  } catch (error) {
    console.error('Erreur generateGroupsOnServer:', error);
    throw error;
  }
}

// ========== PATCH DE SÉPARATION DES SAUVEGARDES GROUPS/INTERFACEV2 ===========

// Nouvelle fonction isolée pour la sauvegarde des groupes
function saveGroups_ISOLATED(groupsData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Utiliser uniquement la feuille GROUPES
    let groupsSheet = ss.getSheetByName('GROUPES');
    if (!groupsSheet) {
      groupsSheet = ss.insertSheet('GROUPES');
      groupsSheet.setTabColor('#9333ea');
    }
    groupsSheet.clear();
    // En-têtes fixes
    const headers = ['Groupe', 'Type', 'EleveID', 'Nom', 'Prenom', 'Sexe', 'Classe', 'Langue'];
    const dataRows = [headers];
    // Remplir les données à partir de groupsData
    groupsData.groups.forEach(group => {
      group.students.forEach(studentId => {
        const studentData = group.studentData?.find(s => s.id === studentId) || {};
        const row = [
          group.name || '',
          groupsData.type || '',
          studentId || '',
          studentData.nom || '',
          studentData.prenom || '',
          studentData.sexe || '',
          studentData.classe || '',
          studentData.lv2 || ''
        ];
        dataRows.push(row);
      });
    });
    if (dataRows.length > 1) {
      groupsSheet.getRange(1, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
      const headerRange = groupsSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#f3f4f6');
      headerRange.setFontWeight('bold');
      groupsSheet.autoResizeColumns(1, headers.length);
      groupsSheet.setFrozenRows(1);
    }
    // Métadonnées dans une feuille séparée
    let metaSheet = ss.getSheetByName('GROUPES_META');
    if (!metaSheet) {
      metaSheet = ss.insertSheet('GROUPES_META');
      metaSheet.hideSheet();
    }
    metaSheet.clear();
    metaSheet.getRange(1, 1, 1, 4).setValues([["Type", "Date", "Config", "Timestamp"]]);
    metaSheet.getRange(2, 1, 1, 4).setValues([[groupsData.type, new Date().toLocaleDateString('fr-FR'), JSON.stringify(groupsData.config), groupsData.timestamp]]);
    // Cache séparé
    const cache = CacheService.getScriptCache();
    cache.put('GROUPS_DATA_ONLY', JSON.stringify(groupsData), 3600);
    return true;
  } catch (error) {
    console.error('Erreur saveGroups_ISOLATED:', error);
    throw new Error('Erreur lors de la sauvegarde des groupes: ' + error.message);
  }
}

// Fonction serveur dédiée pour InterfaceV2
function saveCacheDataInterfaceV2(cacheData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let cacheSheet = ss.getSheetByName('CACHE_INTERFACEV2');
    if (!cacheSheet) {
      cacheSheet = ss.insertSheet('CACHE_INTERFACEV2');
      cacheSheet.hideSheet();
    }
    cacheSheet.clear();
    cacheSheet.getRange(1, 1, 1, 4).setValues([["Date", "Mode", "Disposition", "Source"]]);
    cacheSheet.getRange(2, 1, 1, 4).setValues([[cacheData.date, cacheData.mode, JSON.stringify(cacheData.disposition), 'INTERFACEV2']]);
    return { success: true };
  } catch (error) {
    console.error('Erreur saveCacheDataInterfaceV2:', error);
    return { success: false, error: error.message };
  }
}

function loadCacheDataInterfaceV2() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cacheSheet = ss.getSheetByName('CACHE_INTERFACEV2');
    if (!cacheSheet) {
      return { success: false, error: 'Aucune sauvegarde InterfaceV2 trouvée' };
    }
    const data = cacheSheet.getRange(2, 1, 1, 4).getValues()[0];
    if (data && data[0]) {
      return {
        success: true,
        data: {
          date: data[0],
          mode: data[1],
          disposition: JSON.parse(data[2]),
          source: data[3]
        }
      };
    }
    return { success: false, error: 'Aucune donnée trouvée' };
  } catch (error) {
    console.error('Erreur loadCacheDataInterfaceV2:', error);
    return { success: false, error: error.message };
  }
}