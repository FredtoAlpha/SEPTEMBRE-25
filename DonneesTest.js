/**
 * Script complet pour la génération de données de test dans un contexte scolaire
 * Ce script permet de générer des profils d'élèves avec des caractéristiques réalistes
 * et des codes ASSO (élèves à maintenir ensemble) et DISSO (élèves à séparer)
 * Version: 2.1 (Mai 2025)
 */

//---------- CONSTANTES ET CONFIGURATION ----------//

// Mode développement (activer les tests unitaires)
const DEV_MODE = false;

// Paramètres généraux
const MAX_ABSENTEISTES_PAR_CLASSE = 5;
const ELEVES_PAR_GROUPE_ASSO = 3;
const MOT_DE_PASSE_ADMIN = 'admin123'; // À modifier en production

// Codes DISSO (élèves à séparer)
const QUOTA_DISSO = { 
  D1: 3,  // comportement problématique
  D2: 4,  // travail insuffisant
  D3: 5,  // participation insuffisante
  D4: 6   // autres cas (aléatoire)
};

// Configuration LV2 et options (sera rempli dynamiquement depuis l'onglet _CONFIG)
let CONFIG_LV2 = [];
let CONFIG_OPT = [];

// Listes de noms et prénoms fictifs
const NOMS_GARCONS = [
    "Martin", "Bernard", "Thomas", "Petit", "Robert", "Richard", "Durand", "Dubois", "Moreau", "Laurent",
    "Simon", "Michel", "Lefebvre", "Leroy", "Roux", "David", "Bertrand", "Morel", "Fournier", "Girard",
    "Bonnet", "Dupont", "Lambert", "Fontaine", "Rousseau", "Vincent", "Muller", "Lefevre", "Faure", "Andre",
    "Mercier", "Blanc", "Guerin", "Boyer", "Garnier", "Chevalier", "Francois", "Legrand", "Gauthier", "Garcia",
    "Perrin", "Robin", "Clement", "Morin", "Nicolas", "Henry", "Roussel", "Mathieu", "Gautier", "Masson",
    "Ben Ali", "Boumediene", "Bensaïd", "El-Khaldi", "Khader", "Chraïbi", "Fernandez", "Martinez", "Rodriguez",
    "Lopez", "Hernandez", "Rossi", "Bianchi", "Ricci", "Ferrari", "Esposito", "Kim", "Lee", "Park", "Choi", "Jung",
    "Wojciechowski", "Kowalski", "Nowak", "Wiśniewski", "Dąbrowski"
];

const NOMS_FILLES = [
    "Martin", "Bernard", "Dubois", "Thomas", "Robert", "Richard", "Petit", "Durand", "Leroy", "Moreau",
    "Simon", "Laurent", "Lefebvre", "Michel", "Garcia", "David", "Bertrand", "Roux", "Vincent", "Fournier",
    "Morel", "Girard", "Andre", "Lefevre", "Mercier", "Dupont", "Lambert", "Bonnet", "Francois", "Martinez",
    "Legrand", "Garnier", "Faure", "Rousseau", "Blanc", "Guerin", "Muller", "Henry", "Roussel", "Nicolas",
    "Perrin", "Morin", "Mathieu", "Clement", "Gauthier", "Dumont", "Lopez", "Fontaine", "Chevalier", "Robin",
    "Amina", "Nawel", "Salma", "Isabella", "Lucia", "Sofia", "Maria", "Carmen", "Elena", "Leïla", "Yasmina",
    "Meriem", "Rossi", "Bianchi", "Ricci", "Ferrari", "Esposito", "Kim", "Lee", "Park", "Choi", "Jung",
    "Wojciechowski", "Kowalski", "Nowak", "Wiśniewski", "Dąbrowska"
];

const PRENOMS_GARCONS = [
    "Lucas", "Hugo", "Thomas", "Léo", "Raphaël", "Louis", "Mathis", "Maxime", "Nathan", "Théo",
    "Adam", "Yanis", "Paul", "Baptiste", "Clément", "Alexis", "Alexandre", "Axel", "Ethan", "Jules",
    "Evan", "Arthur", "Noah", "Antoine", "Enzo", "Maël", "Mattéo", "Valentin", "Sacha", "Gabriel",
    "Liam", "Gabin", "Nolan", "Timéo", "Kylian", "Mohamed", "Mathéo", "Tom", "Noé", "Rayan", "Victor",
    "Esteban", "Noa", "Quentin", "Théodore", "Adrien", "Eliott", "Samuel", "Martin", "Gaspard", "Youssef",
    "Mehdi", "Omar", "Rayan", "Anis", "Ilyes", "Alejandro", "Carlos", "Javier", "Miguel", "Diego", "Rafael",
    "Marco", "Leonardo", "Giovanni", "Francesco", "Alessandro", "Wei", "Yuto", "Haruto", "Ren", "Minato",
    "Mateusz", "Jakub", "Filip", "Szymon", "Jan"
];

const PRENOMS_FILLES = [
    "Emma", "Jade", "Louise", "Alice", "Chloé", "Lina", "Léa", "Manon", "Rose", "Anna",
    "Inès", "Camille", "Lola", "Ambre", "Léna", "Zoé", "Juliette", "Julia", "Lou", "Sarah",
    "Lucie", "Mila", "Agathe", "Romane", "Eva", "Inaya", "Lisa", "Louna", "Nina", "Léonie",
    "Charlotte", "Sofia", "Olivia", "Margaux", "Victoire", "Jeanne", "Océane", "Amélie", "Lana", "Maëlys",
    "Clara", "Éléna", "Victoria", "Diane", "Anaïs", "Adèle", "Yasmine", "Soline", "Noémie", "Margot",
    "Isabella", "Lucia", "Sofia", "Maria", "Carmen", "Elena", "Leïla", "Yasmina", "Meriem", "Giulia",
    "Sofia", "Aria", "Mei", "Hana", "Yuki", "Sakura", "Zuzanna", "Amelia", "Wiktoria", "Oliwia"
];

// Dispositifs particuliers avec leurs probabilités
const DISPOSITIFS = [
  { nom: "PAP", proba: 0.06 },    // 6% Plan d'Accompagnement Personnalisé
  { nom: "PPRE", proba: 0.04 },   // 4% Programme Personnalisé de Réussite Éducative
  { nom: "PAI", proba: 0.03 },    // 3% Projet d'Accueil Individualisé
  { nom: "ULIS", proba: 0.01 },   // 1% Unité Localisée pour l'Inclusion Scolaire
  { nom: "GEVASCO", proba: 0.01 } // 1% Guide d'Évaluation Scolaire
];

// Profils types d'élèves (avec pondération)
const PROFILS_TYPES = [
  { 
    nom: 'studieux', 
    poids: 0.5,  // 50%
    generer: () => ({
      ABS: 4, 
      TRA: 4, 
      PART: Math.random() < 0.7 ? 4 : 3, 
      COM: Math.random() < 0.7 ? 4 : 3
    }) 
  },
  { 
    nom: 'absenteiste', 
    poids: 0.08, // 8% mais limité à MAX_ABSENTEISTES_PAR_CLASSE par classe
    generer: () => ({
      ABS: 1, 
      TRA: Math.random() < 0.7 ? 1 : 2, 
      PART: Math.random() < 0.7 ? 1 : 2, 
      COM: Math.random() < 0.7 ? 1 : 2
    }) 
  },
  { 
    nom: 'moyen', 
    poids: 0.25, // 25%
    generer: () => ({
      ABS: 3, 
      TRA: 3, 
      PART: 3, 
      COM: 3
    }) 
  },
  { 
    nom: 'discret', 
    poids: 0.1,  // 10%
    generer: () => ({
      ABS: 4, 
      TRA: Math.random() < 0.5 ? 2 : 3, 
      PART: Math.random() < 0.7 ? 1 : 2, 
      COM: 3
    }) 
  },
  { 
    nom: 'turbulent', 
    poids: 0.07, // 7%
    generer: () => ({
      ABS: Math.random() < 0.5 ? 2 : 3, 
      TRA: 2, 
      PART: Math.random() < 0.5 ? 2 : 3, 
      COM: Math.random() < 0.7 ? 1 : 2
    }) 
  }
];

// Colonnes requises dans les feuilles
const COLONNES_REQUISES = [
  'ID_ELEVE', 'NOM', 'PRENOM', 'NOM_PRENOM', 'SEXE', 
  'LV2', 'OPT', 'COM', 'TRA', 'PART', 'ABS', 
  'DISPO', 'ASSO', 'DISSO'
];

/**
 * Charge la configuration depuis l'onglet _CONFIG
 * Récupère les LV2 et OPT disponibles
 */
function chargerConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('_CONFIG');
  const ui = SpreadsheetApp.getUi();
  
  if (!configSheet) {
    ui.alert(
      'Configuration manquante',
      "L'onglet '_CONFIG' est manquant. Veuillez créer cet onglet avec les paramètres nécessaires.",
      ui.ButtonSet.OK
    );
    throw new Error("Onglet '_CONFIG' introuvable.");
  }
  
  const configData = configSheet.getDataRange().getValues();
  if (configData.length <= 1) {
    ui.alert(
      'Configuration incomplète',
      "L'onglet '_CONFIG' ne contient pas assez de données.",
      ui.ButtonSet.OK
    );
    throw new Error("Onglet '_CONFIG' incomplet.");
  }
  
  let lv2List = '';
  let optList = '';
  let adminPassword = MOT_DE_PASSE_ADMIN; // Valeur par défaut
  
  // Parcourir les lignes de configuration
  for (let i = 1; i < configData.length; i++) {
    const param = configData[i][0];
    const value = configData[i][1];
    
    if (param === 'LV2' && value) {
      lv2List = value;
    } else if (param === 'OPT' && value) {
      optList = value;
    } else if (param === 'ADMIN_PASSWORD' && value) {
      adminPassword = value;
    }
  }
  
  // Diviser les chaînes en tableaux
  CONFIG_LV2 = lv2List ? lv2List.split(',').map(item => item.trim()) : [];
  CONFIG_OPT = optList ? optList.split(',').map(item => item.trim()) : [];
  
  // Mettre à jour le mot de passe admin si trouvé
  if (adminPassword !== MOT_DE_PASSE_ADMIN) {
    MOT_DE_PASSE_ADMIN = adminPassword;
  }
  
  Logger.log(`Configuration chargée: LV2=[${CONFIG_LV2.join(',')}], OPT=[${CONFIG_OPT.join(',')}]`);
}

/**
 * Retourne un objet {colonne: index} pour un header donné
 * @param {string[]} header - Tableau des noms de colonnes
 * @param {string[]} colonnes - Noms à retrouver
 * @return {Object<string,number>}
 */
function getIndexesFromHeader(header, colonnes) {
  const idx = {};
  
  // Si header est vide, retourner un objet vide
  if (!header || header.length === 0) return idx;
  
  // Traiter le cas spécial de ID/ID_ELEVE
  const idColName = header.includes('ID_ELEVE') ? 'ID_ELEVE' : 'ID';
  
  colonnes.forEach(col => {
    if (col === 'ID') {
      idx[col] = header.indexOf(idColName);
    } else {
      idx[col] = header.indexOf(col);
    }
    
    // Si non trouvé, mettre -1
    if (idx[col] < 0) idx[col] = -1;
  });
  
  // S'assurer que ID est à l'index 0 si demandé
  if (colonnes.includes('ID')) idx['ID'] = 0;
  
  return idx;
}

/**
 * Prépare l'en-tête d'une feuille en ajoutant les colonnes requises
 * @param {Sheet} sheet - Feuille Google Sheets
 * @param {Array} requiredColumns - Colonnes requises
 * @return {Array} - En-tête mis à jour
 */
function prepareSheetHeader(sheet, requiredColumns) {
  const maxCol = Math.max(1, sheet.getLastColumn());
  let headers = sheet.getRange(1, 1, 1, maxCol).getValues()[0].map(h => h.toString());
  
  // Forcer ID_ELEVE en A et ajouter les colonnes requises
  if (headers[0] !== 'ID_ELEVE') {
    // Enlever ID_ELEVE et ID s'ils existent ailleurs
    headers = headers.filter(h => h !== 'ID_ELEVE' && h !== 'ID');
    headers.unshift('ID_ELEVE');
    
    // Ajouter les colonnes requises manquantes
    requiredColumns.forEach(c => { 
      if (!headers.includes(c)) headers.push(c); 
    });
    
    // Mettre à jour l'en-tête
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground("#4285f4")
      .setFontColor("white")
      .setFontWeight("bold");
    
    // S'assurer que la feuille a assez de colonnes
    if (sheet.getMaxColumns() < headers.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
    }
  }
  
  return headers;
}

// --------------------------------------------------
// Heuristiques pour détecter les LV2 (langues)
// --------------------------------------------------
const LV2_KEYWORDS = [
  'ITA','ITAL','ITALIEN','LLCA','LATIN',
  'ESP','ESPAGNOL','ALL','ALLEMAND',
  'RUS','RUSSE','CHI','CHINOIS',
  'ARABE','POR','PORTUGAIS'
];

/**
 * Retourne 'LV2' si nomOption est une langue, sinon 'OPT'
 */
function detecterTypeLV2(nomOption) {
  const nom = nomOption.toUpperCase();
  if (LV2_KEYWORDS.some(k => nom.includes(k))) return 'LV2';
  if (/^[A-Z]{3}$/.test(nom) && ['ITA','ESP','ALL','RUS','CHI','POR','JAP','ARA'].includes(nom))
    return 'LV2';
  return 'OPT';
}

/**
 * Lit l'onglet _STRUCTURE et retourne un tableau de classes
 * où chaque origine (séparée par des virgules) devient sa propre entrée.
 */
function chargerStructure() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('_STRUCTURE');
  const ui    = SpreadsheetApp.getUi();
  if (!sheet) {
    ui.alert('Configuration manquante',
             "Onglet '_STRUCTURE' introuvable.",
             ui.ButtonSet.OK);
    throw new Error("Onglet '_STRUCTURE' manquant");
  }

  // Lecture brute, première ligne = header
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    ui.alert('Structure vide',
             "Ajoutez des lignes après l'en-tête dans '_STRUCTURE'.",
             ui.ButtonSet.OK);
    throw new Error("Données _STRUCTURE insuffisantes");
  }
  const headers = values[0].map(String);
  const rows    = values.slice(1);

  // Repérage des colonnes
  let iOrig = headers.indexOf('CLASSE_ORIGINE');
  let iDest = headers.indexOf('CLASSE_DEST');
  let iEff  = headers.indexOf('EFFECTIF');
  let iOpt  = headers.indexOf('OPTIONS');
  headers.forEach((h,i)=>{
    const U=h.toUpperCase();
    if (iOrig<0 && /CLASSE|ORIGINE/.test(U))        iOrig=i;
    if (iDest<0 && /DEST/.test(U))                  iDest=i;
    if (iEff<0  && /EFFECTIF|EFF|NB_?ELEVE/.test(U)) iEff =i;
    if (iOpt<0  && /OPTION/.test(U))                iOpt =i;
  });
  if (iOrig<0||iEff<0||iOpt<0) {
    ui.alert('Colonnes manquantes',
             "Il faut au moins : CLASSE_ORIGINE, EFFECTIF, OPTIONS",
             ui.ButtonSet.OK);
    throw new Error("Colonnes clés absentes");
  }

  // On va construire la liste finale
  const classes = [];
  rows.forEach(row => {
    // explode des origines
    const rawOrig = (row[iOrig]||'').toString();
    rawOrig.split(',')
          .map(s=>s.trim())
          .filter(Boolean)
          .forEach(orig => {
      const effectif = parseInt(row[iEff],10)||0;
      const dest     = row[iDest] ? row[iDest].toString() : '';
      // options
      const optsText = row[iOpt] ? row[iOpt].toString() : '';
      const options = optsText.split(',')
                              .map(p=>p.trim())
                              .filter(Boolean)
                              .map(part=>{
        let [nom,quota] = part.includes('=') ? part.split('=')
                        : part.includes(':') ? part.split(':')
                        : [part, '0'];
        nom = nom.trim();
        quota = parseFloat(quota)||0;
        return { nom, quota, type: detecterTypeLV2(nom) };
      });
      classes.push({ origine: orig, destination: dest, effectif, options });
    });
  });

  Logger.log(`→ ${classes.length} classes chargées : `+
             classes.map(c=>c.origine).join(', '));
  return { classes };
}



/**
 * Test unitaire pour tirerProfilType
 * Vérifie que les proportions des profils générés sont proches des pondérations attendues
 */
function test_TirerProfilType() {
  const nbTests = 1000;
  const compteur = {};
  PROFILS_TYPES.forEach(p => compteur[p.nom] = 0);
  
  // Simuler 1000 tirages avec suffisamment d'absentéistes disponibles
  for (let i = 0; i < nbTests; i++) {
    const profil = tirerProfilType(99); // Valeur élevée pour ne pas limiter
    compteur[profil.nom]++;
  }
  
  // Calculer et afficher les proportions
  Logger.log("=== TEST tirerProfilType ===");
  Logger.log(`Nombre d'échantillons: ${nbTests}`);
  
  let ecartTotal = 0;
  PROFILS_TYPES.forEach(p => {
    const pourcentage = (compteur[p.nom] / nbTests) * 100;
    const attendu = p.poids * 100;
    const ecart = Math.abs(pourcentage - attendu);
    ecartTotal += ecart;
    
    Logger.log(`${p.nom}: ${compteur[p.nom]} (${pourcentage.toFixed(2)}%) - Attendu: ${attendu.toFixed(2)}% - Écart: ${ecart.toFixed(2)}%`);
  });
  
  Logger.log(`Écart total: ${ecartTotal.toFixed(2)}%`);
  Logger.log(`Test ${ecartTotal < 10 ? 'RÉUSSI' : 'ÉCHOUÉ'} (écart < 10%)`);
  
  return ecartTotal < 10; // Le test est réussi si l'écart total est inférieur à 10%
}

/**
 * Invite l'admin et déclenche la génération
 */
function ouvrirGenerationDonnees() {
  if (!verifierMotDePasse('Génération de données de test')) return;
  
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Génération données test',
    'Voulez-vous créer les données fictives pour toutes les classes ?\n' +
    'Cette opération va :\n' +
    '- Créer des élèves fictifs dans les onglets sources\n' +
    '- Respecter les options et quotas définis dans la structure\n' +
    '- Attribuer des notes de comportement et travail réalistes\n' +
    '- Générer des codes ASSO/DISSO cohérents\n\n' +
    'Les données existantes seront remplacées, mais la colonne O sera préservée.',
    ui.ButtonSet.YES_NO
  );
  
  if (resp === ui.Button.YES) {
    genererDonneesStrategiques();
  }
}

/**
 * Vérifie le mot de passe administrateur
 * @param {string} op - Description de l'opération
 * @return {boolean} - True si mot de passe correct
 */
function verifierMotDePasse(op) {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(op, 'Mot de passe administrateur :', ui.ButtonSet.OK_CANCEL);
  
  if (res.getSelectedButton() !== ui.Button.OK) return false;
  
  return res.getResponseText() === MOT_DE_PASSE_ADMIN;
}

/**
 * Crée un profil type aléatoire selon pondération
 * @param {number} absRestants - Nombre d'élèves absentéistes restant à attribuer
 * @return {Object} - Profil avec fonction de génération
 */
function tirerProfilType(absRestants) {
  // Calculer le poids total en excluant absentéiste si on a atteint la limite
  let total = PROFILS_TYPES.reduce(
    (s, p) => s + ((p.nom === 'absenteiste' && absRestants <= 0) ? 0 : p.poids), 
    0
  );
  
  // Tirage aléatoire pondéré
  let tir = Math.random() * total;
  
  for (const p of PROFILS_TYPES) {
    if (p.nom === 'absenteiste' && absRestants <= 0) continue;
    if (tir < p.poids) return p;
    tir -= p.poids;
  }
  
  return PROFILS_TYPES[0]; // Par défaut: studieux
}

/**
 * Génère la liste d’élèves pour une feuille source.
 * @param {Sheet}  sheet – onglet cible (ex. "ECOLE1", "6°1" …)
 * @param {Object} info  – { effectif:number , options:Array , destination?:string , niveau?:number }
 * @return {Array[]}     – lignes à écrire (hors en-tête)
 */
function genererPourFeuille(sheet, info) {

  /* ---------- 0. Préparation en-tête / index ----------------------- */
  const headers = prepareSheetHeader(sheet, COLONNES_REQUISES);
  const idx     = getIndexesFromHeader(headers, COLONNES_REQUISES);

  const N   = info.effectif;
  const log = m => Logger.log(`${sheet.getName()} – ${m}`);

  /* ---------- 1.  SLOTS OPT  --------------------------------------- */
  const slotsOPT = [];
  info.options
      .filter(o => o.type === 'OPT' || CONFIG_OPT.includes(o.nom))
      .forEach(o => { for (let i = 0; i < o.quota; i++) slotsOPT.push(o.nom); });
  while (slotsOPT.length < N) slotsOPT.push('');
  slotsOPT.sort(() => Math.random() - 0.5);
  log(`OPT  : ${slotsOPT.filter(Boolean).length}/${N}`);

  /* ---------- 2.  Niveau & langue implicite ------------------------ */
  const repere = String( info.niveau        ?? 
                         info.destination   ?? 
                         sheet.getName()           );
  const niveau     = parseInt(repere.replace(/[^\d]/g,''),10);   // NaN si absent
  const isSixieme  = niveau === 6;                               // vrai ⇔ 6ᵉ
  const quotasLV2  = info.options
                        .filter(o => o.type === 'LV2')
                        .map(o => o.nom.toUpperCase());
  // Première langue dans _CONFIG qui n’est PAS déjà un quota
  const fallback   = CONFIG_LV2.find(l => !quotasLV2.includes(l.toUpperCase())) || '';
  const lv2Default = (isSixieme || !Number.isFinite(niveau)) ? '' : fallback;

  /* ---------- 3.  SLOTS LV2  --------------------------------------- */
  const slotsLV2 = [];
  info.options
      .filter(o => o.type === 'LV2' || CONFIG_LV2.includes(o.nom))
      .forEach(o => { for (let i = 0; i < o.quota; i++) slotsLV2.push(o.nom); });

  while (lv2Default && slotsLV2.length < N) slotsLV2.push(lv2Default);
  slotsLV2.sort(() => Math.random() - 0.5);
  log(`LV2  : ${slotsLV2.filter(Boolean).length}/${N} (défaut «${lv2Default || 'Ø'}»)`);

  /* ---------- 4.  Génération identité / profils -------------------- */
  const nbF  = Math.floor(N/2) + (Math.random()<0.5 ? 1 : 0);
  const nbM  = N - nbF;
  const data = [];
  let idCtr  = 1;
  let absRest= MAX_ABSENTEISTES_PAR_CLASSE;

  [
    ['M', nbM, NOMS_GARCONS, PRENOMS_GARCONS],
    ['F', nbF, NOMS_FILLES , PRENOMS_FILLES ]
  ].forEach(([sx,count,noms,prens])=>{
    for (let i = 0; i < count; i++) {

      const row = Array(headers.length).fill('');

      /* -- identité ------------------------------------------------- */
      row[idx.ID_ELEVE]   = `${sheet.getName()}${String(idCtr++).padStart(3,'0')}`;
      row[idx.NOM]        = noms [Math.floor(Math.random()*noms.length)];
      row[idx.PRENOM]     = prens[Math.floor(Math.random()*prens.length)];
      row[idx.NOM_PRENOM] = `${row[idx.NOM]} ${row[idx.PRENOM]}`;
      row[idx.SEXE]       = sx;

      /* -- LV2 / OPT ------------------------------------------------ */
      if (idx.LV2 >= 0 && slotsLV2.length) row[idx.LV2] = slotsLV2.pop();
      if (idx.OPT >= 0 && slotsOPT.length) row[idx.OPT] = slotsOPT.pop();

      /* -- profil comportement / travail --------------------------- */
      const p  = tirerProfilType(absRest);
      const sc = p.generer();
      if (sc.ABS === 1) absRest--;
      ['COM','TRA','PART','ABS'].forEach(c => { if (idx[c] >= 0) row[idx[c]] = sc[c]; });

      /* -- dispositifs particuliers -------------------------------- */
      if (idx.DISPO >= 0) {
        for (const {nom,proba} of DISPOSITIFS) {
          if (Math.random() < proba) { row[idx.DISPO] = nom; break; }
        }
      }

      data.push(row);
    }
  });

  return data;
}




/**
 * Fonction générique de répartition ASSO/DISSO
 * @param {string} type - Type de groupe ('ASSO' ou 'DISSO')
 * @param {Object} groupes - Dictionnaire des groupes et compteurs
 * @param {Object|number} config - Configuration des quotas
 * @param {Function} critereFn - Fonction de filtrage des candidats
 * @param {Object} dataMap - Dictionnaire des données par classe
 * @param {Array} sourceSheets - Liste des feuilles sources
 * @param {Array} allCols - Liste des colonnes disponibles
 */
function assignerGroupes(type, groupes, config, critereFn, dataMap, sourceSheets, allCols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  for (const code in groupes) {
    // Déterminer le nombre d'élèves à placer
    let toPlace = (type === 'ASSO') ? config : config[code];
    
    // Mélanger les noms de classes pour une distribution aléatoire
    const noms = sourceSheets.map(s => s.getName()).sort(() => Math.random() - 0.5);
    
    // Utiliser au moins 2 classes, mais pas plus que disponible
    const nbCl = Math.min(Math.max(2, Math.ceil(toPlace / 2)), noms.length);
    const used = noms.slice(0, nbCl);
    
    let idxCl = 0;
    
    while (toPlace > 0) {
      const nomCl = used[idxCl % used.length];
      const data = dataMap[nomCl];
      
      // Récupérer l'en-tête et les indices
      const hdr = ss.getSheetByName(nomCl).getRange(1, 1, 1, ss.getSheetByName(nomCl).getLastColumn()).getValues()[0];
      const idx = {};
      allCols.forEach(c => idx[c] = hdr.indexOf(c));
      
      const col = idx[type];
      
      // Filtrer les candidats selon le critère
      const cands = data.map((r, i) => i).filter(i => {
        // Ne pas réattribuer un élève déjà assigné
        if (data[i][col]) return false;
        
        // Appliquer le critère adapté au type de groupe
        return (type === 'ASSO') ? critereFn(data[i]) : critereFn(data[i], code);
      });
      
      if (cands.length) {
        // Sélectionner un candidat aléatoire parmi ceux éligibles
        const pick = cands[Math.floor(Math.random() * cands.length)];
        data[pick][col] = code;
        groupes[code].total++;
        groupes[code].parClasse[nomCl]++;
        toPlace--;
      } else {
        // S'il n'y a plus de candidats idéaux, prendre parmi les non-assignés
        const free = data.map((_, i) => i).filter(i => !data[i][col]);
        
        if (!free.length) break; // Impossible d'attribuer plus d'élèves
        
        // Attribuer aléatoirement parmi les élèves restants
        while (toPlace > 0 && free.length) {
          const r = free.splice(Math.floor(Math.random() * free.length), 1)[0];
          data[r][col] = code;
          groupes[code].total++;
          groupes[code].parClasse[nomCl]++;
          toPlace--;
        }
        
        break; // Passer au code suivant
      }
      
      idxCl++; // Passer à la classe suivante pour équilibrer
    }
  }
}

/**
 * Critère pour sélection DISSO
 * @param {Array} row - Ligne de données d'un élève
 * @param {string} code - Code DISSO à appliquer
 * @return {boolean} - True si l'élève correspond au critère
 */
function critereDisso(row, code) {
  if (code === 'D1') return row.COM <= 1;        // Comportement problématique
  if (code === 'D2') return row.TRA <= 1;        // Travail insuffisant 
  if (code === 'D3') return row.PART <= 1;       // Participation insuffisante
  if (code === 'D4') return Math.random() < 0.15; // Aléatoire (15% de chance)
  return false;
}

/**
 * Critère pour sélection ASSO
 * @param {Array} row - Ligne de données d'un élève
 * @return {boolean} - True si l'élève correspond au critère
 */
function critereAsso(row) {
  return row.COM >= 3; // Bon comportement
}

/**
 * Écriture des données générées dans chaque feuille
 * @param {Sheet} sheet - Feuille Google Sheets
 * @param {Array} data - Données à écrire
 */
function ecrireDonnees(sheet, data) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const lastCol = sheet.getLastColumn();
  
  // Sauvegarde de la colonne O (index 14)
  let colO = [];
  if (lastCol >= 15 && lastRow > 1) {
    try {
      colO = sheet.getRange(2, 15, lastRow - 1, 1).getValues();
    } catch (e) {
      Logger.log(`Erreur lors de la sauvegarde de la colonne O pour ${sheet.getName()}: ${e}`);
      colO = []; // Réinitialiser en cas d'erreur
    }
  }
  
  // Suppression des validations de données pour éviter les erreurs
  try {
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).clearDataValidations();
    }
  } catch (e) {
    Logger.log(`Erreur lors de la suppression des validations pour ${sheet.getName()}: ${e}`);
  }
  
  // Effacer les données existantes
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
  
  // Préparer les lignes à écrire
  if (!data || data.length === 0) {
    Logger.log(`Aucune donnée à écrire pour ${sheet.getName()}`);
    return;
  }
  
  // S'assurer que chaque ligne a le bon nombre de colonnes
  const rows = data.map(r => {
    if (r.length < lastCol) {
      // Ajouter des cellules vides si la ligne est trop courte
      return r.concat(Array(lastCol - r.length).fill(''));
    } else if (r.length > lastCol) {
      // Tronquer si la ligne est trop longue
      return r.slice(0, lastCol);
    }
    return r;
  });
  
  // Écrire les données
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, lastCol).setValues(rows);
  }
  
  // Restaurer la colonne O
  if (colO.length && lastCol >= 15) {
    const rowsToRestore = Math.min(colO.length, rows.length);
    if (rowsToRestore > 0) {
      sheet.getRange(2, 15, rowsToRestore, 1).setValues(
        colO.slice(0, rowsToRestore)
      );
    }
  }
  
  // Réappliquer des validations de données pour LV2 et SEXE
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idx = {};
    ['SEXE', 'LV2'].forEach(col => {
      idx[col] = headers.indexOf(col);
    });
    
    // Validation SEXE
    if (idx.SEXE >= 0 && rows.length > 0) {
      const ruleSexe = SpreadsheetApp.newDataValidation()
        .requireValueInList(['M', 'F'], true)
        .setAllowInvalid(false)
        .setHelpText("Doit être 'M' ou 'F'")
        .build();
      
      sheet.getRange(2, idx.SEXE + 1, rows.length, 1).setDataValidation(ruleSexe);
    }
    
    // Validation LV2
    if (idx.LV2 >= 0 && rows.length > 0) {
      // Déterminer les options LV2 disponibles à partir des données
      const optionsLV2 = [...new Set(rows.map(r => r[idx.LV2]).filter(v => v))];
      
      if (optionsLV2.length > 0) {
        const ruleLV2 = SpreadsheetApp.newDataValidation()
          .requireValueInList(optionsLV2, true)
          .setAllowInvalid(false)
          .setHelpText(`Choisir parmi: ${optionsLV2.join(', ')}`)
          .build();
        
        sheet.getRange(2, idx.LV2 + 1, rows.length, 1).setDataValidation(ruleLV2);
      }
    }
  } catch (e) {
    Logger.log(`Erreur lors de l'application des validations pour ${sheet.getName()}: ${e}`);
  }
}

/**
 * Point d'entrée principal pour la génération de données stratégiques
 * Cette fonction coordonne tout le processus de génération de données
 */
function genererDonneesStrategiques() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Lancer le test unitaire du tirage des profils (seulement en mode développement)
    if (DEV_MODE) {
      test_TirerProfilType();
    }
    
    // 0. Charger la configuration
    chargerConfig();
    
    // 1. Charger la structure des classes
    const struct = chargerStructure().classes;
    
    // --- DÉBUT correctif : éclate "ECOLE1,ECOLE2" en deux entrées ---
const exploded = [];
struct.forEach(c => {
  c.origine.split(',').forEach(o => {
    exploded.push({ ...c, origine: o.trim() });
  });
});
struct.length = 0;            // on vide l'ancien tableau
Array.prototype.push.apply(struct, exploded); // on le remplit avec les entrées éclatées
// --- FIN correctif ------------------------------------------------

    if (!struct.length) {
      ui.alert(
        'Structure vide',
        'Aucune classe définie dans la structure. Veuillez ajouter au moins une classe dans l\'onglet _STRUCTURE.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 2. Préparer les informations pour chaque classe
    const cs = {};
    struct.forEach(c => {
      cs[c.origine] = {
  effectif : c.effectif,
  options  : c.options,
  niveau   : (c.destination || '').toString()     // ex. "6°1", "5°2", …
               .replace(/[^\d]/g,'')              // garde juste le chiffre
};
    });
    
    // 3. Récupérer les feuilles correspondantes
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = Object.keys(cs).map(n => {
      const sh = ss.getSheetByName(n);
      if (!sh) {
        ui.alert(
          'Onglet manquant', 
          `L'onglet "${n}" défini dans la structure n'existe pas. Veuillez le créer ou corriger la structure.`,
          ui.ButtonSet.OK
        );
        throw new Error(`Onglet ${n} manquant`);
      }
      return sh;
    });
    
    // 4. Générer les données pour chaque classe
    const dataMap = {};
    sheets.forEach(sh => {
      dataMap[sh.getName()] = genererPourFeuille(sh, cs[sh.getName()]);
    });
    
    // 5. Configurer et attribuer les groupes ASSO
    const groupesASSO = { 
      A1: {total: 0, parClasse: {}}, 
      A2: {total: 0, parClasse: {}}, 
      A3: {total: 0, parClasse: {}} 
    };
    
    sheets.forEach(sh => { 
      ['A1', 'A2', 'A3'].forEach(g => {
        groupesASSO[g].parClasse[sh.getName()] = 0;
      }); 
    });
    
    assignerGroupes(
      'ASSO', 
      groupesASSO, 
      ELEVES_PAR_GROUPE_ASSO, 
      critereAsso,
      dataMap, 
      sheets,
      COLONNES_REQUISES
    );
    
    // 6. Configurer et attribuer les groupes DISSO
    const groupesDISSO = { 
      D1: {total: 0, parClasse: {}},
      D2: {total: 0, parClasse: {}},
      D3: {total: 0, parClasse: {}},
      D4: {total: 0, parClasse: {}} 
    };
    
    sheets.forEach(sh => { 
      Object.keys(groupesDISSO).forEach(c => {
        groupesDISSO[c].parClasse[sh.getName()] = 0;
      }); 
    });
    
    assignerGroupes(
      'DISSO', 
      groupesDISSO, 
      QUOTA_DISSO, 
      critereDisso, 
      dataMap, 
      sheets,
      COLONNES_REQUISES
    );
    
    // 7. Écriture des données dans les feuilles
    sheets.forEach(sh => {
      ecrireDonnees(sh, dataMap[sh.getName()]);
    });
    
    // 8. Préparer le rapport final pour l'utilisateur
    let rapportGroupes = "GROUPES GÉNÉRÉS:\n\n";
    
    rapportGroupes += "Groupes ASSO (élèves à maintenir ensemble):\n";
    for (const codeASSO in groupesASSO) {
      const classesImpliquees = Object.keys(groupesASSO[codeASSO].parClasse)
        .filter(c => groupesASSO[codeASSO].parClasse[c] > 0).length;
      rapportGroupes += `- ${codeASSO}: ${groupesASSO[codeASSO].total} élèves répartis dans ${classesImpliquees} classe(s)\n`;
    }
    
    rapportGroupes += "\nGroupes DISSO (élèves à séparer):\n";
    for (const codeDISSO in groupesDISSO) {
      const classesImpliquees = Object.keys(groupesDISSO[codeDISSO].parClasse)
        .filter(c => groupesDISSO[codeDISSO].parClasse[c] > 0).length;
      rapportGroupes += `- ${codeDISSO}: ${groupesDISSO[codeDISSO].total} élèves répartis dans ${classesImpliquees} classe(s)\n`;
    }
    
    rapportGroupes += `\nPolitique des codes DISSO:\n` +
      `- D1: Élèves très perturbateurs (COM=1)\n` +
      `- D2: Élèves qui ne travaillent pas (TRA=1)\n` +
      `- D3: Élèves qui ne participent pas (PART=1)\n` +
      `- D4: Élèves à séparer (aléatoire parmi les restants)`;
    
    // 9. Afficher le message de succès
    ui.alert(
      'Génération terminée', 
      `${sheets.length} classes traitées avec ${Object.values(dataMap).reduce((s, d) => s + d.length, 0)} élèves au total.\n\n${rapportGroupes}`, 
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    // Gérer les erreurs
    ui.alert(
      'Erreur', 
      e.message + '\n\nConsultez les journaux pour plus de détails (Extensions > Apps Script > Exécutions).', 
      ui.ButtonSet.OK
    );
    Logger.log('Erreur dans genererDonneesStrategiques: ' + e.message + '\n' + e.stack);
  }
}

/**
 * Affiche les informations sur le script
 * Cette fonction est appelée depuis votre menu existant
 */
function afficherInfosGenerateur() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    'À propos du générateur', 
    'Générateur de données de test pour système scolaire\n' +
    'Version 2.1 - Mise à jour du: 06/05/2025\n\n' +
    'Ce script permet de générer des données élèves fictives cohérentes\n' +
    'avec des profils variés et des contraintes ASSO/DISSO.\n\n' +
    'Mot de passe administrateur requis pour utilisation.',
    ui.ButtonSet.OK
  );
}