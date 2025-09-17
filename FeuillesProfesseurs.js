// ===== Module 1 : Configuration Utilities =====

/**
 * Récupère la configuration depuis l'onglet _CONFIG (utilise la version centralisée)
 * @return {Object}
 */
function getConfigProfesseurs() {
  return getConfig(); // Appel la version centralisée
}

/**
 * Sauvegarde la liste des fichiers professeurs dans _CONFIG
 * @param {Array<Object>} files
 */
function saveCreatedFiles(files) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let row = data.findIndex(r => r[0] === 'FICHIERS_PROFESSEURS');
  if (row < 0) row = sheet.getLastRow();
  sheet.getRange(row+1,1,1,2)
       .setValues([['FICHIERS_PROFESSEURS', JSON.stringify(files)]]);
}

/**
 * Récupère la liste des fichiers professeurs créés (stockée en JSON dans _CONFIG)
 * @return {Array<{name:string,id:string,url:string}>}
 */
function getFichiersProfesseurs() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'FICHIERS_PROFESSEURS') {
      try {
        return JSON.parse(data[i][1] || '[]');
      } catch (e) {
        Logger.log('JSON invalide dans FICHIERS_PROFESSEURS : ' + e);
        return [];
      }
    }
  }
  return [];
}

// ===== Module 2 : Pondération =====

/**
 * Récupère la table de pondération depuis l'onglet PONDERATION
 * @return {Object} coefficients par matière et critère
 */
function getPonderationTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PONDERATION);
  if (!sheet) return getDefaultPonderationTable();
  const data = sheet.getDataRange().getValues();
  const keys = data[0].slice(2); // entêtes critères
  const table = {};
  for (let r = 1; r < data.length; r++) {
    const mat = data[r][0];
    table[mat] = {};
    for (let c = 2; c < data[r].length; c++) {
      table[mat][keys[c-2]] = Number(data[r][c]) || 0;
    }
  }
  return table;
}

/**
 * Valeurs par défaut si pas d'onglet PONDERATION
 */
function getDefaultPonderationTable() {
  return {
    FRA: {COM:4, TRA:5, PART:3, ABS:1},
    ANG: {COM:3, TRA:3, PART:8, ABS:1},
    MAT: {COM:2, TRA:6, PART:2, ABS:1},
    PHY: {COM:2, TRA:6, PART:2, ABS:1},
    SVT: {COM:1, TRA:5, PART:2, ABS:1},
    ESP: {COM:2, TRA:5, PART:2, ABS:1},
    // etc. pour toutes les matières
  };
}

/**
 * Calcule la moyenne pondérée, résultat borné entre 1 et 4
 * @param {Object} notes  {COM:number,TRA:number,PART:number,ABS:number}
 * @param {Object} coefs  {COM:number,TRA:number,PART:number,ABS:number}
 * @return {number}
 */
function calculerMoyennePonderee(notes, coefs) {
  let num = 0, den = 0;
  ['COM','TRA','PART','ABS'].forEach(k => {
    const v = Number(notes[k]);
    const w = Number(coefs[k]);
    if (!isNaN(v) && w) { num += v*w; den += w; }
  });
  const m = den ? num/den : 0;
  return Math.max(1, Math.min(4, Math.round(m*100)/100));
}

/**
 * Crée le dossier de destination avec gestion des sous-dossiers
 * @param {string} folderPath - Chemin du dossier (ex: "Sous " dans Parent)
 * @param {GoogleAppsScript.Ui.Ui} ui - Interface utilisateur
 * @return {GoogleAppsScript.Drive.Folder|null}
 */
function creerDossierDestination(folderPath, ui) {
  try {
    const parts = folderPath.split(' dans ');
    let parentFolder;
    if (parts.length > 1) {
      const parentName = parts[1];
      const parentIter = DriveApp.getFoldersByName(parentName);
      parentFolder = parentIter.hasNext() ? parentIter.next() : DriveApp.createFolder(parentName);
      const subName = parts[0];
      const subIter = parentFolder.getFoldersByName(subName);
      return subIter.hasNext() ? subIter.next() : parentFolder.createFolder(subName);
    } else {
      const iter = DriveApp.getFoldersByName(folderPath);
      return iter.hasNext() ? iter.next() : DriveApp.createFolder(folderPath);
    }
  } catch (e) {
    ui.alert('Erreur création dossier', e.toString(), ui.ButtonSet.OK);
    return null;
  }
}

// ===== Module 3 & 4 : Création des feuilles professeurs (HTML UI) & Collecte/Écriture des données =====

// ----- HTML Interface -----
/**
 * Affiche une boîte de dialogue HTML pour configurer la création
 */
function creerFeuillesProfesseurs() {
  const tpl = HtmlService.createTemplateFromFile('CreationDialog');
  tpl.matiereList = getListeMatieres(); // [{nom, code}, ...]
  tpl.config = getConfigProfesseurs();
  const html = tpl.evaluate()
     .setWidth(600)
     .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Créer feuilles professeurs');
}

/**
 * Appelé depuis le client HTML après validation du formulaire
 * @param {Object} options {folderPath, niveau, matieres:[], classes:[]}
 */
function creerFeuillesProfesseursAvecOptions(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // --- 1) Normalisation & validation des inputs ---
  const folderPath = options.folderPath;
  const niveau     = options.niveau;

  // s'assurer que matieres est toujours un tableau
  let matieres = options.matieres;
  if (typeof matieres === 'string') {
    matieres = matieres.split(',').map(m => m.trim()).filter(m => m);
  }
  if (!Array.isArray(matieres)) matieres = [];

  // idem pour classes
  let classes = options.classes;
  if (typeof classes === 'string') {
    if (classes.toLowerCase() === 'toutes') {
      // Obtenir toutes les classes disponibles
      classes = getAllSourceSheets().map(sheet => sheet.getName());
    } else {
      // Convertir la chaîne en tableau (séparé par des virgules)
      classes = classes.split(',').map(c => c.trim()).filter(c => c);
    }
  }
  if (!Array.isArray(classes)) classes = [];

  if (!folderPath || !niveau || matieres.length === 0 || classes.length === 0) {
    ui.alert(
      'Données manquantes',
      'Veuillez renseigner le dossier, le niveau, au moins une matière et au moins une classe.',
      ui.ButtonSet.OK
    );
    return;
  }

  // --- 2) Sécurité & dossier cible ---
  if (!verifierMotDePasse('Création des feuilles professeurs')) return;
  const folder = creerDossierDestination(folderPath, ui);
  if (!folder) return;

  // --- 3) Création des classeurs par matière ---
  const filesCreated = [];
  matieres.forEach(codeMat => {
    const matiere = codeMat;
    const newSS   = SpreadsheetApp.create(`${niveau} - ${matiere} - Critères`);
    const file    = DriveApp.getFileById(newSS.getId());
    DriveApp.getRootFolder().removeFile(file);
    folder.addFile(file);

    const profSS = SpreadsheetApp.openById(newSS.getId());
    
    // Supprimer la "Feuille 1" par défaut après avoir créé au moins une vraie feuille
    const defaultSheet = profSS.getSheetByName("Feuille 1");
    
    // Configurer les feuilles de classes
    classes.forEach(classeName => {
      configurerFeuilleProfesseur(profSS, classeName);
    });
    
    // Ajouter l'onglet d'instructions
    ajouterOngletInstructions(profSS, matiere);
    
    // Maintenant, supprimer la Feuille 1 si elle existe encore
    if (defaultSheet) {
      profSS.deleteSheet(defaultSheet);
    }
    
    protegerCellulesFeuilleProfesseur(profSS, classes);

    filesCreated.push({
      name: newSS.getName(),
      url:  profSS.getUrl(),
      id:   newSS.getId()
    });
  });

  // --- 4) Sauvegarde & récapitulatif ---
  saveCreatedFiles(filesCreated);
  ui.alert(
    'Feuilles créées',
    filesCreated.map(f => f.url).join('\n'),
    ui.ButtonSet.OK
  );
  Utils.logAction(`Création de ${filesCreated.length} classeurs`);
}

/**
 * Configure une feuille de saisie pour un professeur
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Le classeur
 * @param {string} classeSource - Nom de la classe source
 */
function configurerFeuilleProfesseur(ss, classeSource) {
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(classeSource);
  if (!sourceSheet) return;
  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length <= 1) return; // Ignorer les onglets vides

  // Créer un onglet pour cette classe
  let sheet = ss.getSheetByName(classeSource);
  if (!sheet) {
    sheet = ss.insertSheet(classeSource);
  } else {
    sheet.clear();
  }

  // Préparer les données pour la feuille professeur (seulement certaines colonnes)
  const headers = ["NOM_PRENOM", "COM", "TRA", "PART", "ABS"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d5dbdb");

  // Chercher l'index de la colonne NOM_PRENOM dans la source
  const sourceHeader = sourceData[0];
  let nomPrenomIndex = sourceHeader.indexOf("NOM_PRENOM");
  if (nomPrenomIndex === -1) {
    const possibleHeaders = ["NOM ET PRENOM", "ELEVE", "NOM COMPLET"];
    for (const header of possibleHeaders) {
      const index = sourceHeader.indexOf(header);
      if (index !== -1) {
        nomPrenomIndex = index;
        break;
      }
    }
    if (nomPrenomIndex === -1) nomPrenomIndex = 3; // fallback
  }

  // Extraire et copier les données nécessaires
  const rowsToCopy = [];
  for (let i = 1; i < sourceData.length; i++) {
    rowsToCopy.push([sourceData[i][nomPrenomIndex], "", "", "", ""]);
  }
  
  if (rowsToCopy.length > 0) {
    sheet.getRange(2, 1, rowsToCopy.length, headers.length).setValues(rowsToCopy);
    
    // Ajouter des listes déroulantes pour COM, TRA, PART, ABS
    const regleCriteres = SpreadsheetApp.newDataValidation()
      .requireValueInList(['1', '2', '3', '4'], true)
      .setAllowInvalid(false)
      .build();
      
    for (let col = 2; col <= 5; col++) {
      sheet.getRange(2, col, rowsToCopy.length, 1).setDataValidation(regleCriteres);
    }
    
    // Mise en forme
    sheet.autoResizeColumn(1); // NOM_PRENOM
    for (let col = 2; col <= 5; col++) {
      sheet.setColumnWidth(col, 60);
      sheet.getRange(2, col, rowsToCopy.length, 1).setHorizontalAlignment("center");
    }
    
    // Formatage conditionnel
    const rules = [];
    const COLORS = {
      CRIT_1: "#F4CCCC", // Rouge clair pour les notes 1-1.99
      CRIT_2: "#FCE5CD", // Orange clair pour les notes 2-2.99
      CRIT_3: "#D9EAD3", // Vert clair pour les notes 3-3.99
      CRIT_4: "#93C47D"  // Vert pour les notes égales à 4
    };
    
    for (let col = 2; col <= 5; col++) {
      const range = sheet.getRange(2, col, rowsToCopy.length, 1);
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("1")
          .setBackground(COLORS.CRIT_1)
          .setRanges([range])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("2")
          .setBackground(COLORS.CRIT_2)
          .setRanges([range])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("3")
          .setBackground(COLORS.CRIT_3)
          .setRanges([range])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("4")
          .setBackground(COLORS.CRIT_4)
          .setRanges([range])
          .build()
      );
    }
    
    sheet.setConditionalFormatRules(rules);
    
    // Alterner les couleurs des lignes
    for (let r = 2; r <= rowsToCopy.length + 1; r += 2) {
      sheet.getRange(r, 1, 1, headers.length).setBackground("#f8f9fa");
    }
  }
}

/**
 * Ajoute un onglet d'instructions à un classeur professeur
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Le classeur
 * @param {string} matiere - Nom de la matière
 */
function ajouterOngletInstructions(ss, matiere) {
  let instructionSheet = ss.getSheetByName("INSTRUCTIONS");
  if (!instructionSheet) {
    instructionSheet = ss.insertSheet("INSTRUCTIONS", 0);
  }
  instructionSheet.clear();
  
  const instructions = [
    ["INSTRUCTIONS POUR LES PROFESSEURS"],
    [""],
    [`Bienvenue dans le fichier de saisie des critères pour la matière ${matiere}.`],
    [""],
    ["Pour chaque élève, veuillez évaluer les critères suivants sur une échelle de 1 à 4 :"],
    [""],
    ["COM = Comportement (1=Difficile, 2=Irrégulier, 3=Satisfaisant, 4=Excellent)"],
    ["TRA = Travail (1=Insuffisant, 2=Fragile, 3=Satisfaisant, 4=Excellent)"],
    ["PART = Participation (1=Jamais, 2=Peu, 3=Régulière, 4=Très active)"],
    ["ABS = Absentéisme (1=Fréquent, 2=Occasionnel, 3=Rare, 4=Aucun)"],
    [""],
    ["Merci de compléter tous les élèves pour les classes qui vous concernent."],
    ["Les données seront collectées pour la répartition des classes."]
  ];
  
  instructionSheet.getRange(1, 1, instructions.length, 1).setValues(instructions);
  instructionSheet.getRange(1, 1).setFontWeight("bold").setFontSize(14);
  instructionSheet.getRange(7, 1, 4, 1).setFontWeight("bold");
  instructionSheet.autoResizeColumn(1);
}

/**
 * Protège les cellules sensibles dans un classeur professeur
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Le classeur
 * @param {Array<string>} selectedClasses - Liste des classes
 */
function protegerCellulesFeuilleProfesseur(ss, selectedClasses) {
  // Protection des onglets d'instructions
  const instructionsSheet = ss.getSheetByName('INSTRUCTIONS');
  if (instructionsSheet) {
    const protection = instructionsSheet.protect();
    protection.setDescription('Protection onglet instructions');
    // Retirer tous les éditeurs actuels
    const me = Session.getEffectiveUser();
    protection.removeEditors(protection.getEditors());
    protection.addEditor(me);
    protection.setWarningOnly(true); // Avertissement mais édition possible
  }
  
  // Protection des en-têtes dans chaque feuille de classe
  selectedClasses.forEach(className => {
    const sheet = ss.getSheetByName(className);
    if (sheet) {
      const headerRange = sheet.getRange(1, 1, 1, 5); // A1:E1
      const protection = headerRange.protect();
      protection.setDescription(`Protection en-tête ${className}`);
      const me = Session.getEffectiveUser();
      protection.removeEditors(protection.getEditors());
      protection.addEditor(me);
      protection.setWarningOnly(false); // Interdiction stricte d'éditer
    }
  });
}

// ----- Collecte et écriture -----
/**
 * Collecte les données dans les classeurs professeurs et écrit les onglets Crit
 */
function collecterDonneesProfesseurs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const fichiers = getFichiersProfesseurs();
  if (!fichiers.length) { ui.alert('Aucun fichier à collecter'); return; }
  if (ui.alert('Collecte', `Traiter ${fichiers.length} fichiers?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  // 1) Charger coefficients
  const ponderationTable = getPonderationTable();
  Utils.logAction('Pondération chargée');

  // 2) Lire et agréger
  const donneesParClasse = {};
  const matieres = [];
  fichiers.forEach(f => {
    try { processFichierProfesseur(f, donneesParClasse, matieres); }
    catch (e) { Utils.logAction('Erreur: ' + e.toString()); }
  });

  // 3) Calcul pondéré
  Object.keys(donneesParClasse).forEach(cl => {
    const eleves = donneesParClasse[cl];
    Object.keys(eleves).forEach(nom => {
      const el = eleves[nom];
      Object.keys(el.criteres).forEach(codeMat => {
        const notes = el.criteres[codeMat];
        const coefs = ponderationTable[codeMat] || getDefaultPonderationTable()[codeMat] || {};
        el.criteres[codeMat].moyenne = calculerMoyennePonderee(notes, coefs);
      });
    });
  });

  // 4) Ecrire par classe
  const sheetsCreated = [];
  Object.keys(donneesParClasse).forEach(cl => {
    const critSheet = prepareFeuilleOngletCriteres(ss, cl, matieres);
    ecrireLignesEtMoyennes(critSheet, donneesParClasse[cl], matieres);
    formatCritSheet(critSheet, matieres);
    sheetsCreated.push(critSheet.getName());
  });

  // 5) Récap
  ui.alert('Collecte terminée', `Onglets créés: ${sheetsCreated.join(', ')}`, ui.ButtonSet.OK);
  Utils.logAction(`Collecte: ${sheetsCreated.join(', ')}`);
}

/**
 * Traite un fichier professeur et récupère les données
 * @param {Object} fichier - Infos du fichier (name, id, url)
 * @param {Object} donneesParClasse - Objet à remplir avec les critères
 * @param {Array<string>} matieres - Liste des matières à compléter
 */
function processFichierProfesseur(fichier, donneesParClasse, matieres) {
  try {
    const ss = SpreadsheetApp.openById(fichier.id);
    const matiere = extractMatiereFromFileName(fichier.name);
    
    // Ajouter la matière à la liste si pas déjà présente
    if (!matieres.includes(matiere)) {
      matieres.push(matiere);
    }
    
    // Pour chaque feuille (= classe)
    const sheets = ss.getSheets().filter(s => !['INSTRUCTIONS'].includes(s.getName()));
    
    sheets.forEach(sheet => {
      const className = sheet.getName();
      
      // Vérifier que la feuille n'est pas vide
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 2 || lastCol < 2) {
        Logger.log(`Feuille vide ignorée: ${className} dans ${fichier.name}`);
        return; // Ignorer cette feuille et passer à la suivante
      }
      
      // Initialiser la structure pour cette classe si n'existe pas encore
      if (!donneesParClasse[className]) {
        donneesParClasse[className] = {};
      }
      
      try {
        // Récupérer les données de la feuille
        const data = sheet.getRange(1, 1, lastRow, Math.min(lastCol, 5)).getValues();
        
        // Indices des colonnes critères (par défaut 1, 2, 3, 4 si non trouvés dans l'en-tête)
        const headers = data[0];
        const comIdx = Math.max(headers.indexOf('COM'), 0);
        const traIdx = Math.max(headers.indexOf('TRA'), 0);
        const partIdx = Math.max(headers.indexOf('PART'), 0);
        const absIdx = Math.max(headers.indexOf('ABS'), 0);
        
        // Pour chaque élève
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const nomPrenom = row[0]; // NOM_PRENOM est dans la première colonne
          
          if (!nomPrenom || nomPrenom.toString().trim() === '') {
            continue; // Ignorer les lignes vides
          }
          
          // Créer entrée élève si n'existe pas
          if (!donneesParClasse[className][nomPrenom]) {
            donneesParClasse[className][nomPrenom] = {
              id: getIdByNomPrenom(nomPrenom, className),
              criteres: {}
            };
          }
          
          // Récupérer et convertir les critères en nombres
          const COM = parseFloat(row[comIdx]) || '';
          const TRA = parseFloat(row[traIdx]) || '';
          const PART = parseFloat(row[partIdx]) || '';
          const ABS = parseFloat(row[absIdx]) || '';
          
          // Ne stocker que si au moins une valeur est présente
          if (COM !== '' || TRA !== '' || PART !== '' || ABS !== '') {
            donneesParClasse[className][nomPrenom].criteres[matiere] = {
              COM: COM,
              TRA: TRA,
              PART: PART,
              ABS: ABS
            };
          }
        }
      } catch (sheetError) {
        Logger.log(`Erreur lors du traitement de la feuille ${className} dans ${fichier.name}: ${sheetError}`);
        // Continuer avec la feuille suivante
      }
    });
  } catch (e) {
    Logger.log(`Erreur globale dans processFichierProfesseur pour ${fichier.name}: ${e}`);
    throw new Error(`Impossible de traiter le fichier ${fichier.name}: ${e.message}`);
  }
}

/**
 * Prépare une feuille d'onglet Critères pour recevoir les données
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Classeur principal
 * @param {string} className - Nom de la classe
 * @param {Array<string>} matieres - Liste des matières
 * @return {GoogleAppsScript.Spreadsheet.Sheet} La feuille créée
 */
function prepareFeuilleOngletCriteres(ss, className, matieres) {
  const critName = `${className}Crit`;
  let critSheet = ss.getSheetByName(critName);
  if (!critSheet) {
    critSheet = ss.insertSheet(critName);
  }
  critSheet.clear();
  
  // En-têtes fixes
  const fixedHeaders = ["ID", "NOM_PRENOM", "COM", "TRA", "PART", "ABS"];
  
  // En-têtes détaillées par matière
  const detailHeaders = [];
  matieres.forEach(mat => {
    detailHeaders.push(`${mat}-COM`, `${mat}-TRA`, `${mat}-PART`, `${mat}-ABS`);
  });
  
  // Écrire l'en-tête complète
  critSheet.getRange(1, 1, 1, fixedHeaders.length + detailHeaders.length)
    .setValues([fixedHeaders.concat(detailHeaders)]);
  
  return critSheet;
}

/**
 * Améliorations pour garantir que les moyennes sont bien numériques
 * et que le format conditionnel s'applique correctement.
 * Calcule les moyennes en nombre (pas en string) et écrit les données.
 * @param {Sheet} critSheet - Onglet cible de critères (déjà vidé et en-têtes créés)
 * @param {Object} donnees - Données par élève et par matière (voir collecterDonneesProfesseurs)
 * @param {Array<string>} matieres - Liste des matières
 */
function ecrireLignesEtMoyennes(critSheet, donnees, matieres) {
  const rows = [];
  for (const nomPrenom in donnees) {
    const eleve = donnees[nomPrenom];
    let somCOM = 0, somTRA = 0, somPART = 0, somABS = 0;
    let countCOM = 0, countTRA = 0, countPART = 0, countABS = 0;
    const details = [];

    matieres.forEach(matiere => {
      const c = eleve.criteres[matiere] || {};
      ['COM','TRA','PART','ABS'].forEach(key => {
        const v = Number(c[key] || NaN);
        details.push(v || '');
        if (!isNaN(v)) {
          if (key==='COM') { somCOM+=v; countCOM++; }
          if (key==='TRA') { somTRA+=v; countTRA++; }
          if (key==='PART') { somPART+=v; countPART++; }
          if (key==='ABS') { somABS+=v; countABS++; }
        }
      });
    });

    // Moyennes comme nombres, pas toFixed() string
    const moyCOM = countCOM ? somCOM/countCOM : '';
    const moyTRA = countTRA ? somTRA/countTRA : '';
    const moyPART = countPART ? somPART/countPART : '';
    const moyABS = countABS ? somABS/countABS : '';
    
    rows.push([ eleve.id||'', nomPrenom, moyCOM, moyTRA, moyPART, moyABS, ...details ]);
  }

  // Écriture et format numérique
  const start = 2;
  const nbLignes = rows.length;
  const nbCols = 6 + matieres.length*4;
  if (nbLignes===0) return;
  
  const range = critSheet.getRange(start, 1, nbLignes, nbCols);
  range.setValues(rows);
  
  // Appliquer format numérique aux moyennes C2:F
  critSheet.getRange(start, 3, nbLignes, 4).setNumberFormat('0.00');
}

/**
 * Met en place les règles de format conditionnel sur C2:F
 * @param {Sheet} sheet
 */
function appliquerFormatConditionnelMoyennes(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow<2) return;
  const rng = sheet.getRange(2,3, lastRow-1,4);
  sheet.clearConditionalFormatRules();
  
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1,1.99)
      .setBackground('#F4CCCC')
      .setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(2,2.99)
      .setBackground('#FCE5CD')
      .setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(3,3.99)
      .setBackground('#D9EAD3')
      .setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(4)
      .setBackground('#93C47D')
      .setRanges([rng]).build()
  ];
  
  sheet.setConditionalFormatRules(rules);
}

/**
 * Formate un onglet de critères pour une meilleure lisibilité et organise l'affichage des colonnes.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - L'onglet de critères
 * @param {Array<string>} matieres - Liste des matières (pour colorer les sections)
 */
function formatCritSheet(sheet, matieres) {
  const ss           = sheet.getParent();
  const spreadsheetId= ss.getId();
  const lastRow      = sheet.getLastRow();
  const lastCol      = sheet.getLastColumn();

  // Si on n'a que l'en-tête (pas de lignes de données), on fige juste l'en-tête et on quitte
  if (lastRow < 2) {
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(6);
    return;
  }

  // Définition des couleurs directement
  const COLORS = {
    CRIT_1: "#F4CCCC", // Rouge clair pour les notes 1-1.99
    CRIT_2: "#FCE5CD", // Orange clair pour les notes 2-2.99
    CRIT_3: "#D9EAD3", // Vert clair pour les notes 3-3.99
    CRIT_4: "#93C47D"  // Vert pour les notes égales à 4
  };
  
  // 1) Figer en-tête + 6 premières colonnes (A–F)
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(6);
  
  // 2) Ajustements colonnes 1–6
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
  for (let col = 3; col <= 6; col++) {
    sheet.setColumnWidth(col, 80);
    sheet.getRange(2, col, lastRow - 1, 1)
      .setHorizontalAlignment('center');
  }
  
  // Entête C–F
  sheet.getRange(1, 3, 1, 4)
    .setBackground('#4CAF50')
    .setFontColor('white')
    .setFontWeight('bold');
  
  // 3) Format conditionnel des moyennes (C2:F + dernière ligne)
  const moyRange = sheet.getRange(2, 3, lastRow - 1, 4);
  
  // S'assurer que les valeurs sont des nombres
  moyRange.setNumberFormat("0.00");
  
  // Supprimer toutes les règles existantes
  sheet.clearConditionalFormatRules();
  
  // Recréer les règles avec les couleurs explicites
  const cfRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1, 1.99)
      .setBackground(COLORS.CRIT_1)
      .setRanges([moyRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(2, 2.99)
      .setBackground(COLORS.CRIT_2)
      .setRanges([moyRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(3, 3.99)
      .setBackground(COLORS.CRIT_3)
      .setRanges([moyRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(4)
      .setBackground(COLORS.CRIT_4)
      .setRanges([moyRange])
      .build()
  ];
  
  // 4) Format conditionnel des détails (G+)
  if (lastCol > 6) {
    const detailRange = sheet.getRange(2, 7, lastRow - 1, lastCol - 6);
    [1, 2, 3, 4].forEach(val => {
      cfRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(String(val))
          .setBackground(COLORS[`CRIT_${val}`])
          .setRanges([detailRange])
          .build()
      );
    });
  }
  
  // Appliquer les règles
  sheet.setConditionalFormatRules(cfRules);
  
  // 5) Grouper/masquer colonnes détail via Sheets API
  if (lastCol > 6) {
    try {
      const requests = [
        {
          addDimensionGroup: {
            range: {
              sheetId: sheet.getSheetId(),
              dimension: 'COLUMNS',
              startIndex: 6,
              endIndex: lastCol
            }
          }
        },
        {
          updateDimensionProperties: {
            range: {
              sheetId: sheet.getSheetId(),
              dimension: 'COLUMNS',
              startIndex: 6,
              endIndex: lastCol
            },
            properties: { hiddenByUser: true },
            fields: 'hiddenByUser'
          }
        }
      ];
      
      try {
        Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
      } catch (apiError) {
        Logger.log("API Sheets non activée ou erreur API: " + apiError);
        // Continuer sans grouper si l'API n'est pas disponible
      }
      
      for (let col = 7; col <= lastCol; col++) {
        sheet.setColumnWidth(col, 60);
      }
    } catch (e) {
      Logger.log('Erreur lors du groupement des colonnes: ' + e.toString());
      // Continuer sans cette fonctionnalité
    }
  }
  
  // 6) Coloration des blocs matières - MODIFIÉ POUR MEILLEURE LISIBILITÉ
  let ptr = 7;
  let isAlternate = false;
  
  matieres.forEach(mat => {
    if (ptr + 3 <= lastCol) {
      const hdr = sheet.getRange(1, ptr, 1, 4);
      const data = sheet.getRange(2, ptr, lastRow - 1, 4)
        .setHorizontalAlignment('center');
      
      // Alternance gris clair/blanc pour les en-têtes de matières
      if (isAlternate) {
        hdr.setBackground('#E8E8E8').setFontColor('black');
      } else {
        hdr.setBackground('#F5F5F5').setFontColor('black');
      }
      
      // Mettre en gras les en-têtes
      hdr.setFontWeight('bold');
      
      // Alternance des couleurs de fond pour les données
      for (let r = 2; r <= lastRow; r += 2) {
        if (isAlternate) {
          sheet.getRange(r, ptr, 1, 4).setBackground('#F8F8F8');
        } else {
          sheet.getRange(r, ptr, 1, 4).setBackground('#FFFFFF');
        }
      }
      
      ptr += 4;
      isAlternate = !isAlternate; // Inverser pour la prochaine matière
    }
  });
  
  // 7) Bordure séparation (col F)
  sheet.getRange(1, 6, lastRow, 1)
    .setBorder(null, true, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // 8) Mise en forme ID/NOM (A–B)
  sheet.getRange(1, 1, 1, 2).setBackground('#d5dbdb');
  for (let r = 2; r <= lastRow; r += 2) {
    sheet.getRange(r, 1, 1, 2).setBackground('#f8f9fa');
  }
  
  // 9) Filtre global
  sheet.getRange(1, 1, 1, lastCol).createFilter();
}

/**
 * Formate l'onglet PONDERATION pour une meilleure lisibilité
 * @param {Sheet} sheet - L'onglet PONDERATION
 * @param {Array} matieres - Liste des matières
 */
function formatPonderationSheet(sheet, matieres) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  // Définir les couleurs explicitement
  const COLORS = {
    CRIT_1: "#F4CCCC", // Rouge clair pour les notes 1-1.99
    CRIT_2: "#FCE5CD", // Orange clair pour les notes 2-2.99
    CRIT_3: "#D9EAD3", // Vert clair pour les notes 3-3.99
    CRIT_4: "#93C47D"  // Vert pour les notes égales à 4
  };
  
  // Figer les deux premières colonnes et la première ligne
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);
  
  // Autoajuster les deux premières colonnes
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
  
  // Largeurs de colonnes uniformes pour les critères
  for (let col = 3; col <= lastCol; col++) {
    sheet.setColumnWidth(col, 60);
  }
  
  // Créer un filtre pour faciliter la navigation
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();
}

// ===== Module 5 : Formatage & Maintenance =====

/**
 * Applique le format conditionnel aux moyennes (colonnes C-F) sur une feuille Crit
 * @param {Sheet} sheet
 */
function appliquerFormatConditionnelMoyennes(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, 3, lastRow - 1, 4);
  sheet.clearConditionalFormatRules();
  const COLORS = { CRIT_1: '#F4CCCC', CRIT_2: '#FCE5CD', CRIT_3: '#D9EAD3', CRIT_4: '#93C47D' };
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(1, 1.99).setBackground(COLORS.CRIT_1).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(2, 2.99).setBackground(COLORS.CRIT_2).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(3, 3.99).setBackground(COLORS.CRIT_3).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(4).setBackground(COLORS.CRIT_4).setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
}

/**
 * Convertit les moyennes en nombres si elles sont stockées en texte
 * @param {Sheet} sheet
 */
function convertirMoyennesEnNombres(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, 3, lastRow - 1, 4);
  const vals = range.getValues();
  const converted = vals.map(row => row.map(cell => {
    if (typeof cell === 'string') {
      const n = Number(cell.replace(',', '.'));
      return isNaN(n) ? '' : n;
    }
    return cell;
  }));
  range.setValues(converted).setNumberFormat('0.00');
}

/**
 * Fonction pour tester et corriger le formatage conditionnel sur la feuille active
 */
function depannerFormatageConditionnel() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Corriger formatage', 'Appliquer correctifs sur feuille active?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  convertirMoyennesEnNombres(sheet);
  appliquerFormatConditionnelMoyennes(sheet);
  ui.alert('Formatage corrigé');
}

/**
 * Diagnostique la feuille active et logue les 5 premières valeurs et règles
 */
function diagnosticFeuille() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, 3, Math.min(5, lastRow - 1), 4);
  const values = range.getValues();
  const formats = range.getNumberFormats();
  Logger.log('--- Diagnostic feuille: ' + sheet.getName() + ' ---');
  values.forEach((row, i) => {
    row.forEach((val, j) => Logger.log(`R${i+2}C${j+3} = ${val} (${formats[i][j]})`));
  });
  const rules = sheet.getConditionalFormatRules();
  Logger.log(`Règles conditionnel: ${rules.length}`);
}

/**
 * Répare tous les onglets se terminant par 'Crit'
 */
function reparerTousOngletsCriteres() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().endsWith('Crit'));
  sheets.forEach(s => { convertirMoyennesEnNombres(s); appliquerFormatConditionnelMoyennes(s); });
  SpreadsheetApp.getUi().alert('Réparation terminée: ' + sheets.length + ' onglets.');
}

// ===== Module 6 : Menus & Aide =====

/**
 * Affiche un guide succinct d'utilisation
 */
function afficherAide() {
  const ui = SpreadsheetApp.getUi();
  const msg =
    '1. Créer feuilles → formulaire\n' +
    '2. Profs remplissent COM/TRA/PART/ABS → 1-4\n' +
    '3. Collecter données → onglets Crit + moyennes pondérées\n' +
    '4. En cas de souci, utilisez Format & Débogage.';
  ui.alert('Aide - Système Critères', msg, ui.ButtonSet.OK);
}

/**
 * Affiche les informations de version
 */
function aPropos() {
  SpreadsheetApp.getUi().alert(
    'Système Critères v1.0\nDéveloppé par [Votre Nom]'
  );
}

/**
 * Pour résoudre les problèmes de formatage conditionnel, exécutez cette fonction sur votre feuille active
 */
function fixerFormatConditionnel() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Effacer toutes les règles existantes
  sheet.clearConditionalFormatRules();
  
  // Plage de moyennes (C2:F + dernière ligne)
  const lastRow = sheet.getLastRow();
  const moyRange = sheet.getRange(2, 3, lastRow - 1, 4);
  
  // Définir des couleurs explicites
  const rouge = "#F4CCCC";
  const orange = "#FCE5CD";
  const vertClair = "#D9EAD3";
  const vert = "#93C47D";
  
  // Créer des règles simples
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1, 1.99)
      .setBackground(rouge)
      .setRanges([moyRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(2, 2.99)
      .setBackground(orange)
      .setRanges([moyRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(3, 3.99)
      .setBackground(vertClair)
      .setRanges([moyRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(4)
      .setBackground(vert)
      .setRanges([moyRange])
      .build()
  ];
  
  // Appliquer les règles
  sheet.setConditionalFormatRules(rules);
  
  // Forcer le format nombre
  moyRange.setNumberFormat("0.00");
}
// ===== FONCTIONS UTILITAIRES MANQUANTES =====

/**
 * Récupère la liste des matières disponibles
 * @return {Array<Object>} Liste d'objets {nom, code}
 */
function getListeMatieres() {
  // Liste des matières standard
  const matieres = [
    {nom: "Français", code: "FRA"},
    {nom: "Anglais", code: "ANG"},
    {nom: "Mathématiques", code: "MAT"},
    {nom: "Physique-Chimie", code: "PHY"},
    {nom: "SVT", code: "SVT"},
    {nom: "Espagnol", code: "ESP"},
    {nom: "Italien", code: "ITA"},
    {nom: "Allemand", code: "ALL"},
    {nom: "Histoire-Géographie", code: "HG"},
    {nom: "EPS", code: "EPS"},
    {nom: "Arts Plastiques", code: "ART"},
    {nom: "Musique", code: "MUS"},
    {nom: "Technologie", code: "TEC"},
    {nom: "Latin", code: "LAT"},
    {nom: "Grec", code: "GRE"}
  ];
  
  return matieres;
}

/**
 * Récupère tous les onglets sources (classes)
 * @return {Array<Sheet>} Liste des onglets de classes sources
 */
function getAllSourceSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  const testSuffix = config.TEST_SUFFIX || "TEST";
  
  // Récupérer tous les onglets qui ne sont pas des onglets système
  const protectedSheets = config.PROTECTED_SHEETS || [];
  const systemSheets = ["PONDERATION", "CreationDialog", "INSTRUCTIONS"];
  
  return ss.getSheets().filter(sheet => {
    const name = sheet.getName();
    // Exclure les onglets système et protégés
    if (protectedSheets.includes(name) || systemSheets.includes(name)) return false;
    // Exclure les onglets qui commencent par _ (onglets système)
    if (name.startsWith("_")) return false;
    // Exclure les onglets TEST
    if (name.endsWith(testSuffix)) return false;
    // Exclure les onglets Crit
    if (name.endsWith("Crit")) return false;
    
    return true;
  });
}

/**
 * Extrait la matière du nom du fichier
 * @param {string} fileName - Nom du fichier (ex: "5e - FRA - Critères")
 * @return {string} Code de la matière
 */
function extractMatiereFromFileName(fileName) {
  // Pattern: "Niveau - MATIERE - Critères"
  const parts = fileName.split(" - ");
  if (parts.length >= 2) {
    return parts[1].trim().toUpperCase();
  }
  
  // Fallback: chercher un code de matière connu dans le nom
  const matieres = getListeMatieres();
  for (const mat of matieres) {
    if (fileName.toUpperCase().includes(mat.code)) {
      return mat.code;
    }
  }
  
  // Dernier recours
  return "MAT";
}

/**
 * Récupère l'ID d'un élève par son nom complet et sa classe
 * @param {string} nomPrenom - Nom et prénom de l'élève
 * @param {string} className - Nom de la classe
 * @return {string} ID de l'élève ou chaîne vide
 */
function getIdByNomPrenom(nomPrenom, className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(className);
  
  if (!sheet) return "";
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Chercher les colonnes ID et NOM_PRENOM
  let idIndex = -1;
  let nomPrenomIndex = -1;
  
  const config = getConfig();
  const idAliases = config.COLUMN_ALIASES?.ID_ELEVE || ["ID_ELEVE", "ID", "IDENTIFIANT"];
  const nomAliases = config.COLUMN_ALIASES?.NOM_PRENOM || ["NOM_PRENOM", "NOM & PRENOM", "ELEVE"];
  
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i].toString().toUpperCase();
    if (idAliases.some(alias => header.includes(alias.toUpperCase()))) {
      idIndex = i;
    }
    if (nomAliases.some(alias => header.includes(alias.toUpperCase()))) {
      nomPrenomIndex = i;
    }
  }
  
  if (idIndex === -1 || nomPrenomIndex === -1) return "";
  
  // Rechercher l'élève
  for (let i = 1; i < data.length; i++) {
    if (data[i][nomPrenomIndex] === nomPrenom) {
      return String(data[i][idIndex] || "");
    }
  }
  
  return "";
}

/**
 * Vérifie le mot de passe administrateur
 * @param {string} context - Contexte de la vérification
 * @return {boolean} true si autorisé
 */
function verifierMotDePasse(context) {
  const ui = SpreadsheetApp.getUi();
  const config = getConfig();
  const adminPassword = config.ADMIN_PASSWORD || config.ADMIN_PASSWORD_DEFAULT || "admin123";
  
  const response = ui.prompt(
    'Authentification requise',
    `${context}\n\nVeuillez entrer le mot de passe administrateur:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return false;
  }
  
  if (response.getResponseText() !== adminPassword) {
    ui.alert('Accès refusé', 'Mot de passe incorrect.', ui.ButtonSet.OK);
    return false;
  }
  
  return true;
}