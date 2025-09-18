// ==================================================================
// ElevesBackendV2.gs — Backend avec validation basée sur _STRUCTURE
// Version autonome qui lit les règles et gère les swaps
// ==================================================================

/**************************** CONFIGURATION LOCALE *********************************/
const ELEVES_MODULE_CONFIG = {
  TEST_SUFFIX: 'TEST',
  SNAPSHOT_SUFFIX: 'INT',
  STRUCTURE_SHEET: '_STRUCTURE'
};

/**************************** ALIAS DES COLONNES *********************************/
const ELEVES_ALIAS = {
  id      : ['ID_ELEVE','ID','UID','IDENTIFIANT','NUM ELÈVE'],
  nom     : ['NOM'],
  prenom  : ['PRENOM','PRÉNOM'],
  sexe    : ['SEXE','S'],
  lv2     : ['LV2','LANGUE2','L2'],
  opt     : ['OPT','OPTION'],
  disso   : ['DISSO','DISSOCIÉ','DISSOCIE'],
  asso    : ['ASSO','ASSOCIÉ','ASSOCIE'],
  com     : ['COM'],
  tra     : ['TRA'],
  part    : ['PART'],
  abs     : ['ABS'],
  source  : ['SOURCE','ORIGINE','CLASSE_ORIGINE'],
  dispo   : ['DISPO','PAI','PPRE','PAP','GEVASCO'],
  mobilite: ['MOBILITE','MOB']
};

/**************************** FONCTIONS UTILITAIRES *********************************/
const _eleves_s   = v => String(v || '').trim();
const _eleves_up  = v => _eleves_s(v).toUpperCase();
const _eleves_num = v => {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
};

function _eleves_idx(head, aliases){
  for(let i=0;i<head.length;i++)
    if(aliases.some(a=>head[i].includes(a))) return i;
  return -1;
}

function _eleves_sanitizeForSerialization(obj) {
  return JSON.parse(JSON.stringify(obj));
}

/************************ getElevesData *********************************/
/**
 * Fonction principale pour récupérer les données des élèves
 * Renvoie [{classe:"5°1",eleves:[{id,nom,…},…]}, …]
 */
function getElevesData(){
  return getElevesDataForMode('TEST');
}

/************************ getElevesDataForMode *********************************/
/**
 * Fonction pour récupérer les données des élèves selon le mode
 * @param {string} mode - Le mode ('TEST', 'CACHE', 'INT')
 * @returns {Array} Données formatées
 */
function getElevesDataForMode(mode) {
  try {
    let suffix = 'TEST'; // Par défaut
    
    switch(mode) {
      case 'TEST':
        suffix = 'TEST';
        break;
      case 'CACHE':
        suffix = 'CACHE';
        break;
      case 'INT':
        suffix = 'INT';
        break;
      default:
        suffix = 'TEST';
        console.warn(`Mode inconnu: ${mode}, utilisation de TEST par défaut`);
    }
    
    console.log(`📊 Chargement des données pour le mode: ${mode} (suffixe: ${suffix})`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = [];

    ss.getSheets().forEach(sh => {
      const name = sh.getName();

      // 1) doit finir par TEST / CACHE / INT
      if(!_eleves_up(name).endsWith(_eleves_up(suffix))) return;

      // ---------- NOUVEAU FILTRE DÉFINITIF ----------
      // 2) on rejette tout onglet dont le nom commence par level_, grp_, groupe…
      if (/^(?:level|grp|groupe|group|niv(?:eau)?)\b/i.test(name)) return;
      // ----------------------------------------------

      const data = sh.getDataRange().getValues();
      if(data.length < 2) return;

      const head = data[0].map(c => _eleves_up(c));
      const col = {};
      Object.keys(ELEVES_ALIAS).forEach(k => col[k] = _eleves_idx(head, ELEVES_ALIAS[k].map(_eleves_up)));
      if(col.id === -1) return;

      const classe = _eleves_s(name.replace(new RegExp(suffix + '$', 'i'), ''));
      const eleves = [];

      for(let r = 1; r < data.length; r++){
        const row = data[r];
        const idRaw = row[col.id];
        if(idRaw === undefined || idRaw === "" || idRaw === null) continue;

        // --- construction de l'élève ---
const eleve = {
  id      : _eleves_s(idRaw),
  nom     : col.nom  !== -1 ? _eleves_s(row[col.nom])  : '',
  prenom  : col.prenom !== -1 ? _eleves_s(row[col.prenom]) : '',
  sexe    : col.sexe !== -1 ? _eleves_up(row[col.sexe]) : '',
  lv2     : col.lv2  !== -1 ? _eleves_up(row[col.lv2])  : '',
  /* ⬇️  ON NE TOUCHE PLUS À L'OPTION — jamais vidée */
  opt     : col.opt  !== -1 ? _eleves_up(row[col.opt])  : '',
  /* le reste inchangé */
  disso   : col.disso !== -1 ? _eleves_up(row[col.disso]) : '',
  asso    : col.asso  !== -1 ? _eleves_up(row[col.asso])  : '',
  scores  : {
    C : col.com  !== -1 ? _eleves_num(row[col.com])  : 0,
    T : col.tra  !== -1 ? _eleves_num(row[col.tra])  : 0,
    P : col.part !== -1 ? _eleves_num(row[col.part]) : 0,
    A : col.abs  !== -1 ? _eleves_num(row[col.abs])  : 0
  },
  source   : col.source   !== -1 ? _eleves_s(row[col.source])   : '',
  dispo    : col.dispo    !== -1 ? _eleves_up(row[col.dispo])   : '',
  mobilite : col.mobilite !== -1 ? (_eleves_up(row[col.mobilite]) || 'LIBRE') : 'LIBRE'
};

        eleves.push(eleve);
      }

      if(eleves.length > 0) {
        result.push({
          classe: classe,
          eleves: eleves
        });
      }
    });

    console.log(`✅ ${result.length} classes trouvées pour le mode ${mode}`);
    return _eleves_sanitizeForSerialization(result);
    
  } catch (e) {
    console.error('Erreur dans getElevesDataForMode:', e);
    return [];
  }
}

/******************** getStructureRules (DEST-ONLY) ********************/
/**
 * Ne lit QUE la colonne « CLASSE_DEST ».
 * – Col. B  = CLASSE_DEST     (obligatoire)
 * – Col. C  = EFFECTIF        (facultative, défaut 28)
 * – Col. D  = OPTIONS         (facultative : "ITA=12, ESP=0")
 *
 * ➜ Renvoie un objet n'incluant **aucune** classe origine :
 * {
 *   "5°1": { capacity: 28, quotas:{ LLCA:8 } },
 *   "5°3": { capacity: 28, quotas:{ ITA:12 } },
 *   …
 * }
 */
function getStructureRules() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(ELEVES_MODULE_CONFIG.STRUCTURE_SHEET);
    if (!sh) {
      console.log('❌ Onglet _STRUCTURE absent'); 
      return {};
    }

    const raw = sh.getDataRange().getValues();
    if (raw.length < 2) return {};

    /* ---------- repérage des colonnes ---------- */
    const head = raw[0].map(h => String(h).toUpperCase().trim());
    const colDest    = head.findIndex(h => h.includes('DEST'));
    const colEff     = head.findIndex(h => h.includes('EFFECTIF'));
    const colOptions = head.findIndex(h => h.includes('OPTION'));

    if (colDest === -1) {
      console.log('❌ Pas de colonne CLASSE_DEST');
      return {};
    }

    /* ---------- constitution de l'objet rules ---------- */
    const rules = {};

    for (let r = 1; r < raw.length; r++) {
      const dest = ('' + raw[r][colDest]).trim();
      if (!dest) continue;                       // ligne vide ou commentée

      /* ---- capacité ---- */
      let capacity = 28;
      if (colEff !== -1 && raw[r][colEff] !== '' && raw[r][colEff] != null) {
        capacity = parseInt(raw[r][colEff], 10) || 28;
      }

      /* ---- quotas ---- */
      const quotas = {};
      if (colOptions !== -1 && raw[r][colOptions] !== '' && raw[r][colOptions] != null) {
        ('' + raw[r][colOptions])
          .replace(/^'/, '')      // retire éventuelle apostrophe force-texte
          .split(',')
          .map(s => s.trim())
          .filter(Boolean)
          .forEach(pair => {
            const [opt, val = '0'] = pair.split('=').map(x => x.trim().toUpperCase());
            if (opt) quotas[opt] = parseInt(val, 10) || 0;   // 0 = interdit
          });
      }

      rules[dest] = { capacity, quotas };
    }

    console.log('✅ rules (DEST-only) :', JSON.stringify(rules, null, 2));
    return _eleves_sanitizeForSerialization(rules);   // garde l'API existante

  } catch (err) {
    console.error('💥 Erreur getStructureRules :', err);
    return {};
  }
}

/******************** updateStructureRules ***********************/
/**
 * Met à jour les règles dans l'onglet _STRUCTURE
 * @param {Object} newRules - Objet avec les nouvelles règles
 * Format: {
 *   "4°1": { capacity: 28, quotas: { ESP: 12, ITA: 8, LATIN: 6 } },
 *   ...
 * }
 */
function updateStructureRules(newRules) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(ELEVES_MODULE_CONFIG.STRUCTURE_SHEET);
    
    if (!sh) {
      // Créer l'onglet s'il n'existe pas
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

/******************** getAvailableClasses ***********************/
/**
 * Retourne la liste des classes détectées automatiquement
 * Utile pour l'interface pour savoir quelles classes existent
 */
function getAvailableClasses() {
  try {
    const testSuf = _eleves_up(ELEVES_MODULE_CONFIG.TEST_SUFFIX);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classes = [];
    
    ss.getSheets().forEach(sh => {
      const name = sh.getName();
      if(_eleves_up(name).endsWith(testSuf)) {
        const classe = _eleves_s(name.replace(new RegExp(testSuf + '$', 'i'), ''));
        if(classe) classes.push(classe);
      }
    });
    
    return _eleves_sanitizeForSerialization(classes.sort());
  } catch (e) {
    console.error('Erreur dans getAvailableClasses:', e);
    return [];
  }
}
/**********************************************************
 *  BLOC 1 – INDEX GÉNÉRAL des élèves (toutes feuilles)
 *  – accepte TOUTES les variantes d'en-tête : ID_ELEVE, UID…
 **********************************************************/
function buildStudentIndex_() {

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheets  = ss.getSheets();

  // mêmes alias que dans ELEVES_ALIAS
  const ID_ALIASES = ['ID_ELEVE','ID','UID','IDENTIFIANT','NUM ELÈVE']
                     .map(s => s.toUpperCase());

  const index  = {};       //  rows[id] = [ …ligne complète… ]
  let   header = null;     //  1er header rencontré (réutilisé pour les snapshots)

  sheets.forEach(sh => {

    if (/INT$/i.test(sh.getName())) return;        // on ignore les *INT
    const data = sh.getDataRange().getValues();
    if (data.length < 2)       return;             // feuille vide

    const head = data[0].map(h => String(h).trim());
    // position de la colonne ID / ID_ELEVE / UID…
    const colId = head.findIndex(h =>
                    ID_ALIASES.includes(h.toUpperCase()));
    if (colId === -1)          return;             // pas de colonne ID, on passe

    if (!header) header = head;                    // on retient le 1er header

    // on indexe toutes les lignes
    for (let r = 1; r < data.length; r++) {
      const id = String(data[r][colId]).trim();
      if (id) index[id] = data[r];
    }
  });

  return { header, rows:index };
}


/**************************************************************************
 *  writeSnapshotSheet_ – version « mise en page 2 »
 *  – ré-affiche d'abord tout (colonnes & lignes) afin de repartir d'un état neutre
 *  – écrit l'en-tête + les données (largeur max recalculée)
 *  – masque les colonnes A-C, puis P → fin
 *  – élargit D (nom & prénom)
 *  – centre F G H I J K M N
 *  – masque toutes les lignes vides sous la dernière donnée
 **************************************************************************/
function writeSnapshotSheet_(sheetName, header, rowData) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(sheetName);

  if (sh) {
    sh.activate();               // on se place dessus → rafraîchit l'UI
    sh.showSheet();              // au cas où elle était masquée
    sh.showColumns(1, sh.getMaxColumns());
    sh.showRows(1, sh.getMaxRows());
    sh.clear();                  // efface contenu + formats
  } else {
    sh = ss.insertSheet(sheetName);
  }

  /* ---------- 1. normaliser largeur des lignes ---------- */
  const maxCols = Math.max(header.length,
                           ...rowData.map(r => r.length));

  const hdr = header.concat(Array(maxCols - header.length).fill(''))
                    .slice(0, maxCols);

  rowData = rowData.map(r =>
            r.concat(Array(maxCols - r.length).fill(''))
             .slice(0, maxCols));

  /* ---------- 2. écriture des données ---------- */
  sh.getRange(1, 1, 1, maxCols).setValues([hdr]);
  if (rowData.length) {
    sh.getRange(2, 1, rowData.length, maxCols).setValues(rowData);
  }

  /* ---------- 3. mise en forme ---------- */

  // 3-a. cacher colonnes A-C
  sh.hideColumns(1, 3);

  // 3-b. cacher colonnes P (16) → fin
  const totCols = sh.getMaxColumns();
  if (totCols > 15) sh.hideColumns(16, totCols - 15);

  // 3-c. élargir la colonne D (≈ 220 px)
  sh.setColumnWidth(4, 220);

  // 3-d. centrer F G H I J K M N
  const centerCols = [6, 7, 8, 9, 10, 11, 13, 14];
  centerCols.forEach(col =>
    sh.getRange(1, col, rowData.length + 1, 1)
      .setHorizontalAlignment('center')
  );

  // 3-e. masquer les lignes vides après la dernière donnée
  const lastDataRow = rowData.length + 1;               // +1 pour l'en-tête
  const totRows     = sh.getMaxRows();
  if (totRows > lastDataRow) {
    sh.hideRows(lastDataRow + 1, totRows - lastDataRow);
  }

  sh.hideSheet();                                       // comme avant
  SpreadsheetApp.flush();                               // force l'UI
}


/**********************************************************
 *  SAUVEGARDE AUTOMATIQUE des classes en onglets <classe>CACHE
 *  – même logique que saveElevesSnapshot mais suffixe CACHE
 **********************************************************/
function saveElevesCache(classMap) {
  const CACHE_SUFFIX = 'CACHE';
  const { header, rows } = buildStudentIndex_();
  if (!header)
    return { success:false, message:'Aucune feuille avec une colonne ID / ID_ELEVE trouvée.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(classMap).forEach(classe => {
    const ids = classMap[classe] || [];
    let rowData = ids.map(id => rows[id] || [id]);
    const maxCols = Math.max(header.length, ...rowData.map(r => r.length));
    rowData = rowData.map(r => (r.length < maxCols) ? r.concat(Array(maxCols - r.length).fill('')) : r.slice(0, maxCols));
    const hdr = (header.length < maxCols) ? header.concat(Array(maxCols - header.length).fill('')) : header.slice(0, maxCols);
    const sheetName = classe + CACHE_SUFFIX;
    let sh = ss.getSheetByName(sheetName);
    sh ? sh.clear() : sh = ss.insertSheet(sheetName);
    // sh.hideSheet(); // Supprimé pour que les onglets CACHE restent visibles
    sh.getRange(1, 1, 1, maxCols).setValues([hdr]);
    if (rowData.length)
      sh.getRange(2, 1, rowData.length, maxCols).setValues(rowData);
    // Enregistrer la date de sauvegarde en propriété de feuille
    sh.getRange(1, maxCols + 1).setValue(new Date().toISOString());
  });
  return { success:true, message:`CACHE OK – ${Object.keys(classMap).length} onglets mis à jour` };
}

/**********************************************************
 *  RESTAURATION depuis les onglets <classe>CACHE
 *  – lit les données et les renvoie pour restauration côté frontend
 **********************************************************/
function restoreElevesCache() {
  const CACHE_SUFFIX = 'CACHE';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(sh => sh.getName().endsWith(CACHE_SUFFIX));
  const result = [];
  sheets.forEach(sh => {
    const name = sh.getName();
    const classe = name.replace(new RegExp(CACHE_SUFFIX + '$', 'i'), '');
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    const head = data[0].map(c => String(c));
    const eleves = [];
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (!row[0]) continue;
      const eleve = {};
      head.forEach((col, i) => {
        eleve[col] = row[i];
      });
      eleves.push(eleve);
    }
    result.push({ classe, eleves });
  });
  return result;
}

/**********************************************************
 *  INFOS sur la dernière sauvegarde CACHE
 *  – renvoie la date la plus récente trouvée dans les onglets <classe>CACHE
 **********************************************************/
function getCacheInfo() {
  const CACHE_SUFFIX = 'CACHE';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(sh => sh.getName().endsWith(CACHE_SUFFIX));
  let lastDate = null;
  sheets.forEach(sh => {
    const data = sh.getDataRange().getValues();
    if (data.length && data[0].length > 0) {
      const dateCell = data[0][data[0].length - 1];
      if (dateCell) {
        const d = new Date(dateCell);
        if (!isNaN(d) && (!lastDate || d > lastDate)) lastDate = d;
      }
    }
  });
  return lastDate ? { exists:true, date: lastDate.toISOString() } : { exists:false };
}

/**
 * Retourne l'objet élève complet (ligne + mapping) à partir de son ID,
 * en scannant toutes les feuilles SAUF celles qui se terminent par INT.
 */
function getEleveById_(id) {

  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const idStr  = String(id).trim();

  // Parcourt chaque feuille (sauf les snapshot INT)
  for (const sh of ss.getSheets()) {
    const name = sh.getName();
    if (/INT$/.test(name)) continue;      // on saute les snapshots

    const data = sh.getDataRange().getValues();
    if (data.length === 0) continue;

    const header = data[0].map(h => String(h).trim().toUpperCase());

    const colCache = new Map();
    const col = (field) => {
      const key = String(field || '').toLowerCase();
      if (colCache.has(key)) return colCache.get(key);

      const aliases = ELEVES_ALIAS[key];
      if (!aliases) {
        colCache.set(key, -1);
        return -1;
      }

      const idx = _eleves_idx(header, aliases.map(_eleves_up));
      colCache.set(key, idx);
      return idx;
    };

    const colId = col('id');
    if (colId === -1) continue;                       // pas de colonne ID → on passe

    // Recherche de l'ID dans cette feuille
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][colId]).trim() === idStr) {

        const valueAt = (field) => {
          const idx = col(field);
          return idx !== -1 ? data[r][idx] : undefined;
        };

        return {
          id       : idStr,
          nom      : valueAt('nom'),
          prenom   : valueAt('prenom'),
          sexe     : valueAt('sexe'),
          lv2      : valueAt('lv2'),
          opt      : valueAt('opt'),
          scores   : {
            C : valueAt('com'),
            T : valueAt('tra'),
            P : valueAt('part'),
            A : valueAt('abs')
          },
          mobilite : (() => {
            const direct = valueAt('mobilite');
            return direct !== undefined ? direct : valueAt('dispo');
          })(),
          asso     : valueAt('asso'),
          disso    : valueAt('disso'),
          source   : valueAt('source')
        };
      }
    }
  }
  // ID introuvable
  return null;
}

/************************ getElevesStats ****************************/
/**
 * Calcule les statistiques des élèves
 */
function getElevesStats(){
  try {
    const groups = getElevesData();
    if(!groups || groups.length === 0) {
      return {
        global: {COM: 0, TRA: 0, PART: 0},
        parClasse: []
      };
    }
    
    // Filtrer pour ne garder que les vraies classes (exclure les groupes parasites)
    const realClasses = groups.filter(grp => {
      const className = String(grp.classe || '').trim();
      // Regex souple pour exclure tous les types de groupes parasites (espaces, tirets, underscores)
      return !className.match(/^(?:level[\s_-]?|niv(?:eau)?[\s_-]?|grp(?:oupe)?[\s_-]?|groupe[\s_-]?|group[\s_-]?|niveau[\s_-]?)/i) && 
             className.length > 0;
    });
    
    const g = {COM: 0, TRA: 0, PART: 0, count: 0};
    const list = [];
    
    realClasses.forEach(grp => {
      if (!grp.eleves || grp.eleves.length === 0) return;
      
      const s = {COM: 0, TRA: 0, PART: 0};
      grp.eleves.forEach(e => {
        if (e.scores) {
          s.COM += e.scores.C || 0;
          s.TRA += e.scores.T || 0;
          s.PART += e.scores.P || 0;
        }
      });
      
      const n = grp.eleves.length;
      list.push({
        classe: grp.classe,
        COM: Math.round(s.COM / n * 100) / 100,
        TRA: Math.round(s.TRA / n * 100) / 100,
        PART: Math.round(s.PART / n * 100) / 100
      });
      
      g.COM += s.COM;
      g.TRA += s.TRA;
      g.PART += s.PART;
      g.count += n;
    });
    
    const globalStats = g.count > 0 ? {
      COM: Math.round(g.COM / g.count * 100) / 100,
      TRA: Math.round(g.TRA / g.count * 100) / 100,
      PART: Math.round(g.PART / g.count * 100) / 100
    } : {COM: 0, TRA: 0, PART: 0};
    
    return _eleves_sanitizeForSerialization({
      global: globalStats,
      parClasse: list
    });
  } catch (e) {
    console.error('Erreur dans getElevesStats:', e);
    return {global: {COM: 0, TRA: 0, PART: 0}, parClasse: []};
  }
}

/**
 * Récupère les scores MATH et FR depuis les fichiers INT
 * @return {Object} Résultat avec les scores importés
 */
function getINTScores() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const intSheets = sheets.filter(sheet => /INT$/i.test(sheet.getName()));
    
    if (intSheets.length === 0) {
      return { success: false, error: 'Aucun fichier INT trouvé' };
    }
    
    const scores = [];
    
    intSheets.forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      
      const headers = data[0].map(h => String(h || '').toUpperCase());
      
      // Trouver les colonnes nécessaires
      const idCol = headers.findIndex(h => h === 'ID' || h === 'ID_ELEVE');
      const mathCol = headers.findIndex(h => h === 'MATH' || h === 'MATHEMATIQUES');
      const frCol = headers.findIndex(h => h === 'FR' || h === 'FRANCAIS' || h === 'FRANÇAIS');
      
      if (idCol === -1) {
        Logger.log(`Colonne ID non trouvée dans ${sheet.getName()}`);
        return;
      }
      
      // Parcourir les données
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const id = String(row[idCol] || '').trim();
        
        if (!id) continue;
        
        const score = {
          id: id,
          MATH: null,
          FR: null,
          source: sheet.getName()
        };
        
        // Récupérer les scores MATH
        if (mathCol !== -1 && row[mathCol] !== undefined && row[mathCol] !== '') {
          const mathScore = parseFloat(row[mathCol]);
          if (!isNaN(mathScore) && mathScore >= 0 && mathScore <= 20) {
            // Convertir sur 4 si nécessaire
            score.MATH = mathScore > 4 ? (mathScore / 5) : mathScore;
          }
        }
        
        // Récupérer les scores FR
        if (frCol !== -1 && row[frCol] !== undefined && row[frCol] !== '') {
          const frScore = parseFloat(row[frCol]);
          if (!isNaN(frScore) && frScore >= 0 && frScore <= 20) {
            // Convertir sur 4 si nécessaire
            score.FR = frScore > 4 ? (frScore / 5) : frScore;
          }
        }
        
        // Ne garder que si au moins un score est présent
        if (score.MATH !== null || score.FR !== null) {
          scores.push(score);
        }
      }
    });
    
    Logger.log(`Import terminé: ${scores.length} scores récupérés depuis ${intSheets.length} fichiers INT`);
    
    return {
      success: true,
      scores: scores,
      count: scores.length,
      sources: intSheets.map(s => s.getName())
    };
    
  } catch (error) {
    Logger.log(`Erreur dans getINTScores: ${error.message}`);
    return { success: false, error: error.message };
  }
}

/**
 * Algorithme d'optimisation amélioré pour les groupes
 * @param {Array} students - Liste des élèves
 * @param {number} numGroups - Nombre de groupes souhaités
 * @param {Object} options - Options d'optimisation
 * @return {Object} Résultat de l'optimisation
 */
function optimizeGroupsAdvanced(students, numGroups, options = {}) {
  const startTime = new Date();
  
  // Options par défaut
  const config = {
    iterations: options.iterations || 1000,
    populationSize: options.populationSize || 50,
    mutationRate: options.mutationRate || 0.1,
    crossoverRate: options.crossoverRate || 0.8,
    eliteSize: options.eliteSize || 5,
    constraints: options.constraints || {},
    weights: {
      balance: options.weights?.balance || 1.0,
      diversity: options.weights?.diversity || 0.5,
      constraints: options.weights?.constraints || 2.0,
      size: options.weights?.size || 0.3
    }
  };
  
  try {
    Logger.log(`Démarrage optimisation avancée: ${students.length} élèves, ${numGroups} groupes`);
    
    // Créer la population initiale
    let population = [];
    for (let i = 0; i < config.populationSize; i++) {
      population.push(createRandomSolution(students, numGroups));
    }
    
    let bestSolution = null;
    let bestScore = -Infinity;
    let generationsWithoutImprovement = 0;
    
    // Boucle d'évolution
    for (let generation = 0; generation < config.iterations; generation++) {
      // Évaluer la population
      const evaluated = population.map(solution => ({
        solution: solution,
        score: evaluateSolutionAdvanced(solution, config.weights, config.constraints)
      }));
      
      // Trier par score
      evaluated.sort((a, b) => b.score - a.score);
      
      // Garder la meilleure solution
      if (evaluated[0].score > bestScore) {
        bestScore = evaluated[0].score;
        bestSolution = JSON.parse(JSON.stringify(evaluated[0].solution));
        generationsWithoutImprovement = 0;
      } else {
        generationsWithoutImprovement++;
      }
      
      // Arrêt prématuré si pas d'amélioration
      if (generationsWithoutImprovement > 50) {
        Logger.log(`Arrêt prématuré à la génération ${generation} (pas d'amélioration)`);
        break;
      }
      
      // Créer la nouvelle population
      const newPopulation = [];
      
      // Élitisme : garder les meilleurs
      for (let i = 0; i < config.eliteSize; i++) {
        newPopulation.push(evaluated[i].solution);
      }
      
      // Croisement et mutation
      while (newPopulation.length < config.populationSize) {
        const parent1 = selectParent(evaluated);
        const parent2 = selectParent(evaluated);
        
        let child;
        if (Math.random() < config.crossoverRate) {
          child = crossover(parent1, parent2);
        } else {
          child = JSON.parse(JSON.stringify(parent1));
        }
        
        if (Math.random() < config.mutationRate) {
          child = mutate(child);
        }
        
        newPopulation.push(child);
      }
      
      population = newPopulation;
      
      // Log périodique
      if (generation % 100 === 0) {
        Logger.log(`Génération ${generation}: Meilleur score = ${bestScore.toFixed(2)}`);
      }
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log(`Optimisation terminée en ${duration.toFixed(1)}s. Score final: ${bestScore.toFixed(2)}`);
    
    return {
      success: true,
      solution: bestSolution,
      score: bestScore,
      duration: duration,
      generations: config.iterations
    };
    
  } catch (error) {
    Logger.log(`Erreur dans optimizeGroupsAdvanced: ${error.message}`);
    return { success: false, error: error.message };
  }
}

/**
 * Crée une solution aléatoire
 */
function createRandomSolution(students, numGroups) {
  const groups = Array(numGroups).fill(null).map(() => []);
  
  // Distribution aléatoire
  students.forEach(student => {
    const groupIndex = Math.floor(Math.random() * numGroups);
    groups[groupIndex].push(student);
  });
  
  return groups;
}

/**
 * Évalue une solution avec des critères avancés
 */
function evaluateSolutionAdvanced(groups, weights, constraints) {
  let score = 0;
  
  // Équilibre des tailles
  const sizes = groups.map(g => g.length);
  const avgSize = sizes.reduce((a, b) => a + b, 0) / sizes.length;
  const sizeVariance = sizes.reduce((sum, size) => sum + Math.pow(size - avgSize, 2), 0) / sizes.length;
  score += weights.size * (100 - sizeVariance * 10);
  
  // Équilibre F/M
  let balanceScore = 0;
  groups.forEach(group => {
    const fCount = group.filter(s => s.sexe === 'F').length;
    const mCount = group.filter(s => s.sexe === 'M').length;
    const total = fCount + mCount;
    if (total > 0) {
      const ratio = Math.abs(fCount - mCount) / total;
      balanceScore += (1 - ratio) * 100;
    }
  });
  score += weights.balance * (balanceScore / groups.length);
  
  // Diversité des scores
  let diversityScore = 0;
  groups.forEach(group => {
    const scores = group.map(s => (s.scores?.M || 0) + (s.scores?.F || 0));
    if (scores.length > 1) {
      const variance = calculateVariance(scores);
      diversityScore += variance * 10;
    }
  });
  score += weights.diversity * (diversityScore / groups.length);
  
  // Respect des contraintes
  let constraintScore = 100;
  if (constraints.disso) {
    // Vérifier les dissociations
    groups.forEach(group => {
      const dissoCodes = new Set();
      group.forEach(student => {
        if (student.disso) {
          if (dissoCodes.has(student.disso)) {
            constraintScore -= 20; // Pénalité pour violation
          }
          dissoCodes.add(student.disso);
        }
      });
    });
  }
  
  if (constraints.asso) {
    // Vérifier les associations
    const assoGroups = {};
    groups.forEach((group, groupIndex) => {
      group.forEach(student => {
        if (student.asso) {
          if (!assoGroups[student.asso]) {
            assoGroups[student.asso] = [];
          }
          assoGroups[student.asso].push(groupIndex);
        }
      });
    });
    
    Object.values(assoGroups).forEach(groupIndices => {
      const uniqueGroups = new Set(groupIndices);
      if (uniqueGroups.size > 1) {
        constraintScore -= 30; // Pénalité pour séparation
      }
    });
  }
  
  score += weights.constraints * constraintScore;
  
  return Math.max(0, score);
}

/**
 * Calcule la variance d'un tableau de valeurs
 */
function calculateVariance(values) {
  if (values.length === 0) return 0;
  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / values.length;
  return variance;
}

/**
 * Sélectionne un parent par roulette
 */
function selectParent(evaluated) {
  const totalScore = evaluated.reduce((sum, item) => sum + item.score, 0);
  let random = Math.random() * totalScore;
  
  for (const item of evaluated) {
    random -= item.score;
    if (random <= 0) {
      return item.solution;
    }
  }
  
  return evaluated[0].solution;
}

/**
 * Croise deux solutions
 */
function crossover(parent1, parent2) {
  const child = parent1.map(group => [...group]);
  
  // Échange aléatoire de quelques élèves
  const numSwaps = Math.floor(Math.random() * 3) + 1;
  
  for (let i = 0; i < numSwaps; i++) {
    const group1 = Math.floor(Math.random() * child.length);
    const group2 = Math.floor(Math.random() * child.length);
    
    if (child[group1].length > 0 && child[group2].length > 0) {
      const student1Index = Math.floor(Math.random() * child[group1].length);
      const student2Index = Math.floor(Math.random() * child[group2].length);
      
      const temp = child[group1][student1Index];
      child[group1][student1Index] = child[group2][student2Index];
      child[group2][student2Index] = temp;
    }
  }
  
  return child;
}

/**
 * Mute une solution
 */
function mutate(solution) {
  const mutated = solution.map(group => [...group]);
  
  // Mutation : déplacer un élève aléatoire
  const nonEmptyGroups = mutated.filter(group => group.length > 0);
  if (nonEmptyGroups.length > 1) {
    const sourceGroup = nonEmptyGroups[Math.floor(Math.random() * nonEmptyGroups.length)];
    const targetGroup = mutated[Math.floor(Math.random() * mutated.length)];
    
    if (sourceGroup.length > 0) {
      const studentIndex = Math.floor(Math.random() * sourceGroup.length);
      const student = sourceGroup.splice(studentIndex, 1)[0];
      targetGroup.push(student);
    }
  }
  
  return mutated;
}

/*************************** doGet *******************************/
/**
 * Point d'entrée pour l'application web
 */
function doGet(e) {
  // Si on demande spécifiquement l'interface V2
  if (e && e.parameter && e.parameter.page === 'interfaceV2') {
    return HtmlService.createHtmlOutputFromFile('InterfaceV2')
           .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
           .setTitle('Répartition Classes - Interface Compacte avec Swaps')
           .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }
  
  // Sinon, retourner l'interface par défaut
  return HtmlService.createHtmlOutputFromFile('InterfaceV2')
         .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
         .setTitle('Répartition Classes - Interface Compacte avec Swaps');
}

/************************* Fonctions de test **************************/
/**
 * Teste la lecture des règles depuis _STRUCTURE
 */
function testGetStructureRules() {
  const rules = getStructureRules();
  console.log('Règles depuis _STRUCTURE:', JSON.stringify(rules, null, 2));
  return rules;
}

/**
 * Teste la fonction getElevesData
 */
function testGetElevesDataV2() {
  const result = getElevesData();
  console.log('Résultat getElevesData:', JSON.stringify(result, null, 2));
  
  // Afficher un résumé
  if (result.length > 0) {
    console.log('\nRésumé:');
    result.forEach(group => {
      console.log(`- ${group.classe}: ${group.eleves.length} élèves`);
    });
  }
  
  return result;
}

/**
 * Crée un onglet _STRUCTURE de démonstration
 */
function createDemoStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "_STRUCTURE";
  
  // Vérifier si l'onglet existe déjà
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Attention',
      `L'onglet ${sheetName} existe déjà. Voulez-vous le remplacer?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      console.log('Création annulée');
      return;
    }
    
    ss.deleteSheet(sheet);
  }
  
  // Créer le nouvel onglet
  sheet = ss.insertSheet(sheetName);
  
  // En-têtes
  const headers = ["", "CLASSE_DEST", "EFFECTIF", "OPTIONS"];
  
  // Détecter automatiquement les classes existantes
  const classes = getAvailableClasses();
  
  // Données de démonstration basées sur les classes détectées
  const data = [headers];
  
  if (classes.length > 0) {
    classes.forEach(classe => {
      // Générer des quotas par défaut selon le niveau
      let options = "";
      if (classe.startsWith("5°")) {
        options = "ESP=12, ITA=8, CHAV=6, EURO=4";
      } else if (classe.startsWith("4°")) {
        options = "ESP=12, ITA=8, LATIN=6, GRECO=4";
      } else if (classe.startsWith("3°")) {
        options = "ESP=14, ITA=10, LATIN=4";
      } else {
        options = "ESP=12, ITA=8";
      }
      
      data.push(["", classe, 28, options]);
    });
  } else {
    // Si aucune classe détectée, exemple générique
    data.push(["", "4°1", 28, "ESP=12, ITA=8, LATIN=6, EURO=4"]);
    data.push(["", "4°2", 28, "ESP=12, ITA=8, LATIN=6"]);
  }
  
  // Écrire les données
  sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  
  // Formater
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#5b21b6')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
  
  // Ajouter une note explicative
  sheet.getRange(1, 4).setNote(
    "Format: OPTION=QUOTA\n" +
    "Exemple: ESP=12, ITA=8\n" +
    "Si une option n'est pas listée, elle n'est pas autorisée dans cette classe."
  );
  
  console.log(`Onglet ${sheetName} créé avec succès pour les classes: ${classes.join(', ')}`);
  
  // Retourner les règles créées
  return testGetStructureRules();
}

/**
 * Crée ou met à jour un onglet 'CACHE' dans le Google Sheet courant.
 * Copie la structure et les données des onglets de travail (TEST ou DEF).
 * Retourne {success:true} ou {success:false, error:...}
 */
function createCacheTab() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    // Chercher un onglet existant nommé 'CACHE'
    var cacheSheet = ss.getSheetByName('CACHE');
    if (!cacheSheet) {
      cacheSheet = ss.insertSheet('CACHE');
    } else {
      cacheSheet.clear();
    }
    // Trouver le premier onglet TEST ou DEF (ex: 4°1TEST, 4°2DEF...)
    var sourceSheet = null;
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (/TEST$|DEF$/.test(name)) {
        sourceSheet = sheets[i];
        break;
      }
    }
    if (!sourceSheet) {
      return {success: false, error: 'Aucun onglet TEST ou DEF trouvé'};
    }
    // Copier la structure et les données
    var data = sourceSheet.getDataRange().getValues();
    cacheSheet.getRange(1,1,data.length,data[0].length).setValues(data);
    return {success: true};
  } catch(e) {
    return {success: false, error: e.message};
  }
}
/**
 * Fonction attendue par l'interface V2 pour charger les classes et règles
 * @param {string} mode - 'TEST', 'CACHE', 'INT'
 * @return {Object} {success: true, data, rules}
 */
function getClassesData(mode) {
  try {
    const data = getElevesDataForMode(mode);
    const rules = getStructureRules();
    return {
      success: true,
      data: data,
      rules: rules
    };
  } catch (e) {
    return {
      success: false,
      error: e.toString()
    };
  }
}

// ========== AJOUT : Fonction de création des onglets INT (snapshot) ========== //
/**
 * Crée les onglets <classe>INT à partir de la disposition fournie
 * @param {Object} disposition - mapping {classe: [eleve, ...]}
 * @return {Object} {success: true/false, message: string}
 */
function saveElevesSnapshot(disposition) {
  try {
    if (!disposition || typeof disposition !== 'object') {
      return { success: false, error: 'Disposition invalide' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const header = [
      'ID_ELEVE', 'NOM', 'PRENOM', 'NOM & PRENOM', 'SEXE', 'LV2', 'OPT', 'COM', 'TRA', 'PART', 'ABS',
      'DISPO', 'ASSO', 'DISSO', 'SOURCE', 'FIXE', 'CLASSE_FINALE', 'CLASSE DEF', '', 'MOBILITE',
      'SCORE F', 'SCORE M', 'GROUP'
    ];
    // Récupérer tous les élèves en mode TEST (objet par élève, donc accès direct aux champs)
    const elevesData = getElevesDataForMode('TEST');
    // Création d'un index {id: eleve}
    const elevesIndex = {};
    elevesData.forEach(grp => {
      grp.eleves.forEach(eleve => elevesIndex[eleve.id] = eleve);
    });
    
    Object.keys(disposition).forEach(classe => {
      const eleveIds = disposition[classe] || [];
      const rowData = eleveIds.map(eleveId => {
        const eleve = elevesIndex[eleveId];
        if (!eleve) {
          console.warn(`⚠️ Élève ${eleveId} non trouvé dans l'index`);
          return [
            eleveId, '', '', eleveId, '', 'ESP', '', '', '', '', '', '', '', '', '', '', classe, classe, '', 'LIBRE', '', '', '', '', '', ''
          ];
        }
        // LV2 toujours présent sinon ESP
        let lv2 = (eleve.lv2 || '').toString().trim().toUpperCase();
        if (!lv2) lv2 = 'ESP';
        return [
          eleve.id,
          eleve.nom || '',
          eleve.prenom || '',
          (eleve.nom || '') + ' ' + (eleve.prenom || ''),
          eleve.sexe || '',
          lv2,
          eleve.opt || '',
          eleve.scores?.C ?? '',
          eleve.scores?.T ?? '',
          eleve.scores?.P ?? '',
          eleve.scores?.A ?? '',
          eleve.dispo || '',
          eleve.asso || '',
          eleve.disso || '',
          eleve.source || '',
          '', // FIXE
          classe,
          classe,
          '', // colonne S à masquer
          eleve.mobilite || 'LIBRE',
          '', // SCORE F
          '', // SCORE M
          ''  // GROUP
        ];
      });
      if (rowData.length === 0) rowData.push(Array(header.length).fill(''));
      const sheetName = classe + 'INT';
      let sh = ss.getSheetByName(sheetName);
      if (sh) sh.clear();
      else sh = ss.insertSheet(sheetName);
      sh.getRange(1, 1, 1, header.length).setValues([header]);
      sh.getRange(2, 1, rowData.length, header.length).setValues(rowData);

      // Masquer colonnes demandées
      const hiddenCols = [1,2,3,16,17,18,19];
      hiddenCols.forEach(idx => { try { sh.hideColumns(idx, 1); } catch(e){} });
      sh.setColumnWidth(4, 220); // élargir colonne D
      const centerCols = [6,7,8,9,10,11,13,14,20];
      centerCols.forEach(idx => { if (idx <= header.length) sh.getRange(1, idx, rowData.length+1, 1).setHorizontalAlignment('center'); });
      sh.hideSheet(); // optionnel : masque l’onglet INT par défaut
    });
    
    // ========== ÉTAPE 5 : FORMATAGE AUTOMATIQUE DES ONGLETS INT ==========
    try {
      console.log('🎨 Début du formatage automatique des onglets INT...');
      
      // Vérifier si le formateur INT est disponible
      if (typeof INTFormatter !== 'undefined') {
        const formatResult = INTFormatter.run({
          applyColumnFormatting: true,
          applyConditionalRules: true,
          calculateAndWriteStats: true,
          applyFiltersAndZebra: true,
          cleanUnused: true
        });
        
        if (formatResult.success) {
          console.log('✅ Formatage INT réussi:', formatResult.message);
        } else {
          console.warn('⚠️ Formatage INT avec erreurs:', formatResult.message);
        }
      } else {
        console.log('ℹ️ Formateur INT non disponible, onglets créés sans formatage');
      }
    } catch (formatError) {
      console.warn('⚠️ Erreur lors du formatage INT:', formatError.message);
      // Ne pas faire échouer la sauvegarde pour une erreur de formatage
    }
    
    return { success: true, message: `Onglets INT créés avec succès pour ${Object.keys(disposition).length} classes` };
  } catch (e) {
    console.error('❌ Erreur dans saveElevesSnapshot:', e);
    return { success: false, error: e.toString() };
  }
}
// Expose la fonction pour l'interface Google Apps Script
if (typeof global !== 'undefined') {
  global.saveElevesSnapshot = saveElevesSnapshot;
}