/**  INT Formatter v2 – PARTIE 1  (constants + utils)
 *   Dépendances : aucune
 *   À placer dans son propre fichier .gs
 */
var INTFormatter = (function () {   // ← ouverture du namespace
  'use strict';

  /* =====================  CONSTANTES  ===================== */
  const STATS_ROW_COUNT_INT        = 6;
  const CLEANUP_BUFFER_ROWS_INT    = 3;
  const CLEANUP_BUFFER_COLS_INT    = 1;

  const HEADER_BACKGROUND_COLOR_INT      = '#C6E0B4';
  const HEADER_FONT_WEIGHT_INT           = 'bold';
  const HEADER_HORIZONTAL_ALIGNMENT_INT  = 'center';
  const HEADER_VERTICAL_ALIGNMENT_INT    = 'middle';

  const EVEN_ROW_BACKGROUND_INT          = '#f8f9fa';
  const STATS_SEPARATOR_ROW_BACKGROUND_INT = '#dcdcdc';

  const SEXE_COLORS_INT   = { M: '#B3CFEC', F: '#FFD1DC' };
  const LV2_COLORS_INT    = { ESP: '#E59838', ITA: '#73C6B6', ALL: '#F4B084', AUTRE: '#A3E4D7' };
  const OPTION_FORMATS_INT = [
    { text: 'GREC',  bgColor: '#C0392B', fgColor: '#FFFFFF' },
    { text: 'LATIN', bgColor: '#641E16', fgColor: '#FFFFFF' },
    { text: 'LLCA',  bgColor: '#F4B084', fgColor: '#000000' },
    { text: 'CHAV',  bgColor: '#6C3483', fgColor: '#FFFFFF' },
    { text: 'UPE2A', bgColor: '#D5D8DC', fgColor: null      }
  ];
  const SCORE_COLORS_INT  = { 1: '#FF0000', 2: '#FFD966', 3: '#A8E4BC', 4: '#006400' };

  const DEFAULT_REQUIRED_HEADERS_CONFIG_INT = {
    ID_ELEVE   : ['ID_ELEVE', 'ID', 'IDENTIFIANT', 'IDELEVE'],
    NOM_PRENOM : ['NOM & PRENOM', 'NOM_PRENOM', 'NOM PRENOM', 'ELEVE', 'NOM ET PRENOM'],
    SEXE       : ['SEXE', 'GENRE'],
    LV2        : ['LV2', 'LANGUE', 'LANGUE VIVANTE 2'],
    OPT        : ['OPT', 'OPTION', 'OPTIONS'],
    COM        : ['COM', 'COMPORTEMENT', 'H'],
    TRA        : ['TRA', 'TRAVAIL', 'I'],
    PART       : ['PART', 'PARTICIPATION', 'J'],
    ABS        : ['ABS', 'ABSENCES', 'ASSIDUITE', 'K'],
    INDICATEUR : ['INDICATEUR', 'IND', 'L'],
    ASSO       : ['ASSO', 'ASSOCIATION', 'ASSOC'],
    DISSO      : ['DISSO', 'DISSOCIATION', 'DISSOC'],
    CLASSE_DEF : ['CLASSE DEF', 'CLASSE DEFINITIVE', 'CLASSE_DEF'],
    SOURCE     : ['SOURCE', 'ORIGINE']
  };
  const CANONICAL_ID_ELEVE_HEADER_INT = 'ID_ELEVE';

  /* ====================  UTILITAIRES  ===================== */

  /**
   * Normalise un en-tête : supprime espaces & caractères spéciaux + majuscules.
   */
  const normalizeHeaderINT = s =>
    String(s || '')
      .replace(/[\s\u00A0\u200B-\u200D]/g, '')
      .toUpperCase();

  /**
   * Applique des largeurs de colonnes (tableau d’objets {col, width}).
   */
  function setColumnWidthsINT(sheet, cfgs) {
    cfgs.forEach(c => {
      if (c.col > 0 && c.col <= sheet.getMaxColumns()) {
        try { sheet.setColumnWidth(c.col, c.width); }
        catch (e) { Logger.log(`Width col ${c.col}: ${e}`); }
      }
    });
  }

  /**
   * Cache en toute sécurité n colonnes à partir de startColumn.
   */
  function safeHideColumnsINT(sheet, startColumn, numColumns = 1) {
    if (startColumn > 0 && startColumn <= sheet.getMaxColumns()) {
      const n = Math.min(numColumns, sheet.getMaxColumns() - startColumn + 1);
      if (n > 0) {
        try { sheet.hideColumns(startColumn, n); }
        catch (e) { Logger.log(`Hide cols ${startColumn}-${startColumn+n-1}: ${e}`); }
      }
    }
  }

  /**
   * Supprime lignes/colonnes hors-zone en gardant un buffer.
   */
  function cleanUnusedRowsAndColumnsINT(sheet, lastUsedRow, lastUsedCol,
                                        bufferRows = 0, bufferCols = 0) {
    try {
      const maxR = sheet.getMaxRows();
      const maxC = sheet.getMaxColumns();
      const rowStart = lastUsedRow + 1 + bufferRows;
      if (rowStart < maxR) sheet.deleteRows(rowStart, maxR - rowStart + 1);
      const colStart = lastUsedCol + 1 + bufferCols;
      if (colStart < maxC) sheet.deleteColumns(colStart, maxC - colStart + 1);
    } catch (e) {
      Logger.log(`Clean ${sheet.getName()}: ${e.message}\n${e.stack}`);
    }
  }

  /**
   * Renvoie la liste des noms de feuilles se terminant par “INT”.
   */
  function getINTSheetsNames() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss.getSheets()
      .map(sh => sh.getName())
      .filter(name => /INT$/i.test(name));
  }

  /* =====================  PARTIE 2  (règles et statistiques)  =====================
 *  À placer APRÈS la PARTIE 1, toujours avant la fermeture du module.
 *  – Contient :
 *      • applyConditionalFormattingRules_INT()
 *      • calculateSheetStatsINT()
 *      • writeSheetStatsINT()
 *  Les constantes et utilitaires sont déjà définis dans la Partie 1.
 * =============================================================================== */


/**
 * Pose toutes les règles de mise en forme conditionnelle pour une feuille INT.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet  Feuille cible
 * @param {number} nRows  Nombre de lignes de données
 * @param {Object} colMap  Dictionnaire {CLÉ -> index de colonne}
 */
function applyConditionalFormattingRules_INT(sheet, nRows, colMap) {
  try {
    sheet.clearConditionalFormatRules();
    const rules = [];
    const R = SpreadsheetApp.newConditionalFormatRule;
    const start = 2;  // première ligne de données

    // helper interne
    const add = (range, ruleBuilder) => {
      try { rules.push(ruleBuilder.setRanges([range]).build()); }
      catch (e) { Logger.log(`Cond-format ${range.getA1Notation()}: ${e}`); }
    };

    /* ---------- SEXE ---------- */
    if (colMap.SEXE > 0 && nRows > 0) {
      const rg = sheet.getRange(start, colMap.SEXE, nRows, 1);
      add(rg, R().whenTextEqualTo('M').setBackground(SEXE_COLORS_INT.M));
      add(rg, R().whenTextEqualTo('F').setBackground(SEXE_COLORS_INT.F));
      rg.setFontWeight('bold').setHorizontalAlignment('center');
    }

    /* ---------- LV2 ---------- */
    if (colMap.LV2 > 0 && nRows > 0) {
      const rg = sheet.getRange(start, colMap.LV2, nRows, 1);
      add(rg, R().whenTextEqualTo('ESP').setBackground(LV2_COLORS_INT.ESP));
      // toute autre valeur non vide
      add(rg, R().whenFormulaSatisfied(
        `=AND(ISTEXT(INDIRECT("RC",FALSE)),INDIRECT("RC",FALSE)<>"ESP")`)
        .setBackground(LV2_COLORS_INT.AUTRE));
      rg.setFontWeight('bold').setHorizontalAlignment('center');
    }

    /* ---------- OPTIONS ---------- */
    if (colMap.OPT > 0 && nRows > 0) {
      const rg = sheet.getRange(start, colMap.OPT, nRows, 1);
      OPTION_FORMATS_INT.forEach(opt => {
        const b = R().whenTextEqualTo(opt.text).setBackground(opt.bgColor);
        if (opt.fgColor) b.setFontColor(opt.fgColor);
        add(rg, b);
      });
      rg.setFontWeight('bold').setHorizontalAlignment('center');
    }

    /* ---------- SCORES (COM/TRA/PART/ABS) ---------- */
    ['COM', 'TRA', 'PART', 'ABS'].forEach(key => {
      if (colMap[key] > 0 && nRows > 0) {
        const rg = sheet.getRange(start, colMap[key], nRows, 1);
        for (let s = 1; s <= 4; s++) {
          const fc = (s === 1 || s === 4) ? '#FFFFFF' : '#000000';
          add(rg, R().whenNumberEqualTo(s)
                     .setBackground(SCORE_COLORS_INT[s])
                     .setFontColor(fc));
        }
        rg.setFontWeight('bold').setHorizontalAlignment('center');
      }
    });

    // colonnes L, M, N toujours en gras/centré
    ['INDICATEUR', 'ASSO', 'DISSO'].forEach(k => {
      if (colMap[k] > 0 && nRows > 0)
        sheet.getRange(start, colMap[k], nRows, 1)
             .setFontWeight('bold').setHorizontalAlignment('center');
    });

    if (rules.length) sheet.setConditionalFormatRules(rules);
  } catch (e) {
    Logger.log(`applyConditionalFormattingRules_INT: ${e.message}`);
  }
}



/**
 * Calcule toutes les statistiques nécessaires pour la feuille.
 * Renvoie un objet prêt à être exploité par writeSheetStatsINT().
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} numRows  Nombre de lignes de données
 * @param {Object} colMap
 * @param {number} idCol   Index (1-based) de la colonne ID_ELEVE (souvent 1)
 * @return {Object|null}
 */
function calculateSheetStatsINT(sheet, numRows, colMap, idCol) {
  const data = sheet.getRange(2, 1, numRows, sheet.getMaxColumns()).getValues()
                    .filter(r => String(r[idCol - 1]).trim() !== '');

  if (!data.length) return null;

  // helpers internes
  const getCol = key => colMap[key] ? data.map(r => r[colMap[key]-1]) : [];
  const countEq = (arr, v) => arr.filter(x => String(x).trim().toUpperCase() === v).length;
  const countNotEmpty = arr => arr.filter(x => String(x).trim() !== '').length;
  const toNums = arr => arr.map(x => Number(String(x).replace(',', '.')) || 0);
  const avg = nums => {
    const f = nums.filter(n => n > 0);
    return f.length ? f.reduce((a,b)=>a+b,0)/f.length : 0;
  };

  // jeu de données
  const sexe = getCol('SEXE');
  const lv2  = getCol('LV2');
  const opt  = getCol('OPT');

  const res = {
    genreCounts     : [ countEq(sexe,'F'), countEq(sexe,'M') ],
    lv2Counts       : [ countEq(lv2,'ESP'), lv2.length - countEq(lv2,'ESP') - countEq(lv2,'') ],
    optionsCounts   : [ countNotEmpty(opt) ],
    criteresScores  : {},
    criteresMoyennes: []
  };

  ['COM','TRA','PART','ABS'].forEach(k => {
    const nums = toNums(getCol(k));
    res.criteresScores[k] = { 1:0,2:0,3:0,4:0 };
    nums.forEach(n => { if (n>=1 && n<=4) res.criteresScores[k][n]++; });
    res.criteresMoyennes.push( avg(nums) );
  });

  return res;
}



/**
 * Écrit (ou efface) le bloc de statistiques en bas de la feuille.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row   1-based row où commence le bloc
 * @param {Object} colMap
 * @param {Object|null} stats   Objet retourné par calculateSheetStatsINT
 */
function writeSheetStatsINT(sheet, row, colMap, stats) {

  // largeur maxi à nettoyer
  const maxCol = Math.max(
    colMap.NOM_PRENOM, colMap.SEXE, colMap.LV2, colMap.OPT,
    colMap.COM, colMap.TRA, colMap.PART, colMap.ABS
  );

  sheet.getRange(row, 1, 7, maxCol).clearContent().clearFormat();

  // petit utilitaire local
  const set = (r,c,v,fmt={}) => {
    if (c<1 || c>sheet.getMaxColumns()) return;
    const cell = sheet.getRange(r,c).setValue(v);
    if (fmt.bg)    cell.setBackground(fmt.bg);
    if (fmt.fg)    cell.setFontColor(fmt.fg);
    if (fmt.bold)  cell.setFontWeight('bold');
    if (fmt.align) cell.setHorizontalAlignment(fmt.align);
    if (fmt.fmt)   cell.setNumberFormat(fmt.fmt);
  };

  if (!stats) {
    set(row,1,'Pas de données',{italic:true});
    return;
  }

  /* ------ effectif / LV2 / options (colonnes fixes) ------ */
  set(row, 5, stats.genreCounts[0], {bg:SEXE_COLORS_INT.F, align:'center', bold:true});
  set(row+1,5, stats.genreCounts[1], {bg:SEXE_COLORS_INT.M, align:'center', fg:'#FFF', bold:true});
  set(row, 6, stats.lv2Counts[0],   {bg:LV2_COLORS_INT.ESP, align:'center', bold:true});
  set(row+1,6, stats.lv2Counts[1],   {bg:LV2_COLORS_INT.AUTRE, align:'center', bold:true});
  set(row, 7, stats.optionsCounts[0], {align:'center', bold:true});

  /* ------ scores détaillés + moyennes ------ */
  ['COM','TRA','PART','ABS'].forEach((k,i) => {
    const col = colMap[k];
    if (!col) return;
    const S = stats.criteresScores[k];
    set(row  , col, S[4], {bg:SCORE_COLORS_INT[4], fg:'#FFF', align:'center', bold:true});
    set(row+1, col, S[3], {bg:SCORE_COLORS_INT[3],             align:'center', bold:true});
    set(row+2, col, S[2], {bg:SCORE_COLORS_INT[2],             align:'center', bold:true});
    set(row+3, col, S[1], {bg:SCORE_COLORS_INT[1], fg:'#FFF',  align:'center', bold:true});
    set(row+4, col, stats.criteresMoyennes[i], {fmt:'#,##0.00', align:'center', bold:true});
  });
}

/* ====================  FIN PARTIE 2  ==================== */

/* =====================  PARTIE 3  (cœur du formatage)  =====================
 *  À placer APRÈS la PARTIE 2 et AVANT la PARTIE 4.
 *  – Contient :
 *      • miseEnFormeOngletsINT()          (fonction publique principale)
 *      • buildColumnMapINT()              (aide interne)
 *  Les constantes + utilitaires + règles + stats sont déjà dispo.
 * ========================================================================== */


/*---------------------------------------------------------------------------
 * Construit un dictionnaire {CLE HEADER -> index de colonne} pour la feuille.
 *---------------------------------------------------------------------------*/
function buildColumnMapINT(sheet) {
  const rawHdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                       .getValues()[0]
                       .map(normalizeHeaderINT);

  const idxFromAliases = aliases => {
    const list = [].concat(aliases).map(normalizeHeaderINT);
    const i = rawHdrs.findIndex(h => list.includes(h));
    return i >= 0 ? i + 1 : -1;
  };

  const map = {};
  Object.keys(DEFAULT_REQUIRED_HEADERS_CONFIG_INT).forEach(k => {
    map[k] = idxFromAliases(DEFAULT_REQUIRED_HEADERS_CONFIG_INT[k]);
  });
  return map;
}



/*----------------------------------------------------------------------------
 * Fonction PRINCIPALE : met en forme tous les onglets se terminant par “INT”.
 *----------------------------------------------------------------------------*/
function miseEnFormeOngletsINT(options) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const cfg  = Object.assign({   // toutes les options activées par défaut
    applyColumnFormatting:  true,
    applyConditionalRules:  true,
    calculateAndWriteStats: true,
    applyFiltersAndZebra:   true,
    cleanUnused:            true
  }, options || {});

  const sheets = getINTSheetsNames();
  if (!sheets.length) {
    return { success:true, message:'Aucun onglet INT trouvé.', bilan:{errors:[]} };
  }

  /* ----------- liste de validation pour CLASSE DEF ----------- */
  const classeDefList = sheets.map(n => n.replace(/INT$/i,'') + 'INT DEF').sort();
  const validationDEF = SpreadsheetApp.newDataValidation()
        .requireValueInList(classeDefList,true)
        .setAllowInvalid(false)
        .setHelpText('Sélectionnez la classe DEF')
        .build();

  /* ----------- variables de bilan ----------- */
  const bilan = { classes:[], genreCounts:[], lv2Counts:[],
                  optionsCounts:[], criteresMoyennes:[], errors:[] };

  /* ===================  BOUCLE SUR CHAQUE FEUILLE  =================== */
  sheets.forEach((name,idx) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      bilan.errors.push(`Feuille ${name} introuvable`);
      return;
    }

    try {
      /* == A. En-têtes & colonnes indispensables ===================== */
      const colMap = buildColumnMapINT(sheet);

      // forcer ID_ELEVE en colonne A
      if (colMap.ID_ELEVE !== 1) {
        const idxId = colMap.ID_ELEVE;
        if (idxId > 0) {
          sheet.moveColumns(sheet.getRange(1, idxId, sheet.getMaxRows()), 1);
        } else {
          sheet.insertColumnBefore(1);
          sheet.getRange(1,1).setValue(CANONICAL_ID_ELEVE_HEADER_INT);
        }
      }

      // style header uniforme
      sheet.getRange(1,1,1,sheet.getLastColumn())
           .setFontWeight(HEADER_FONT_WEIGHT_INT)
           .setBackground(HEADER_BACKGROUND_COLOR_INT)
           .setHorizontalAlignment(HEADER_HORIZONTAL_ALIGNMENT_INT)
           .setVerticalAlignment(HEADER_VERTICAL_ALIGNMENT_INT);

      /* == B. colonne CLASSE DEF (toujours colonne 18 = R) ============ */
      const CLASSE_DEF_COL = 18;
      if (sheet.getMaxColumns() < CLASSE_DEF_COL)
        sheet.insertColumnsAfter(sheet.getMaxColumns(),
                                 CLASSE_DEF_COL - sheet.getMaxColumns());
      if (!sheet.getRange(1,CLASSE_DEF_COL).getValue())
        sheet.getRange(1,CLASSE_DEF_COL).setValue('CLASSE DEF')
             .setFontWeight(HEADER_FONT_WEIGHT_INT)
             .setBackground(HEADER_BACKGROUND_COLOR_INT)
             .setHorizontalAlignment(HEADER_HORIZONTAL_ALIGNMENT_INT)
             .setVerticalAlignment(HEADER_VERTICAL_ALIGNMENT_INT);
      colMap.CLASSE_DEF = CLASSE_DEF_COL;

      /* == C. masquage colonnes A-C (ID, cache, etc.) ================ */
      sheet.hideColumns(1,3);

      /* == D. lignes de données ====================================== */
      const lastRow   = sheet.getLastRow();
      const ids       = sheet.getRange(2,1,Math.max(0,lastRow-1),1).getValues().flat();
      const lastDataRow = ids.map(x => String(x).trim())
                             .lastIndexOf(ids.slice().reverse().find(v => v)) + 2;
      const nRows = Math.max(0, lastDataRow-1);

      /* == E. Formatage colonnes ===================================== */
      if (cfg.applyColumnFormatting) {
        // cacher colonnes techniques
        if (colMap.SOURCE > 0)            safeHideColumnsINT(sheet,colMap.SOURCE);
        safeHideColumnsINT(sheet,16,2);   // colonnes P-Q
        safeHideColumnsINT(sheet,19,5);   // colonnes S–W

        // largeurs
        const w = [];
        if (colMap.NOM_PRENOM>0) w.push({col:colMap.NOM_PRENOM,width:200});
        [colMap.SEXE,colMap.LV2,colMap.OPT,
         colMap.INDICATEUR,colMap.ASSO,colMap.DISSO].forEach(c=>{
           if (c>0) w.push({col:c,width:70});
         });
        [colMap.COM,colMap.TRA,colMap.PART,colMap.ABS]
          .forEach(c=>{if(c>0)w.push({col:c,width:60});});
        w.push({col:CLASSE_DEF_COL,width:100});
        setColumnWidthsINT(sheet,w);
        if (sheet.getFrozenRows()===0 && sheet.getLastRow()>0) sheet.setFrozenRows(1);
      }

      /* == F. Règles conditionnelles ================================= */
      if (cfg.applyConditionalRules && nRows>0)
        applyConditionalFormattingRules_INT(sheet,nRows,colMap);

      /* == G. Zebra + filtre ========================================= */
      if (cfg.applyFiltersAndZebra && nRows>0) {
        for (let r=2;r<=lastDataRow;r+=2)
          sheet.getRange(r,1,1,sheet.getMaxColumns())
               .setBackground(EVEN_ROW_BACKGROUND_INT);

        // ligne séparatrice stats
        sheet.insertRowAfter(lastDataRow);
        sheet.getRange(lastDataRow+1,1,1,sheet.getMaxColumns())
             .setBackground(STATS_SEPARATOR_ROW_BACKGROUND_INT);

        // filtre
        if (sheet.getFilter()) sheet.getFilter().remove();
        sheet.getRange(1,1,lastDataRow,sheet.getMaxColumns()).createFilter();
      }

      /* == H. Validation liste déroulante CLASSE DEF ================= */
      if (nRows>0)
        sheet.getRange(2,CLASSE_DEF_COL,nRows,1).setDataValidation(validationDEF);

      /* == I. Statistiques ========================================== */
      let statsObj = null;
      if (cfg.calculateAndWriteStats) {
        statsObj = nRows ? calculateSheetStatsINT(sheet,nRows,colMap,1) : null;
        const statsRow = lastDataRow + 2;   // saute la ligne séparatrice
        writeSheetStatsINT(sheet,statsRow,colMap,statsObj);

        if (statsObj) {
          bilan.classes.push(name);
          bilan.genreCounts.push(statsObj.genreCounts);
          bilan.lv2Counts.push(statsObj.lv2Counts);
          bilan.optionsCounts.push(statsObj.optionsCounts);
          bilan.criteresMoyennes.push(statsObj.criteresMoyennes);
        }
      }

      /* == J. Nettoyage fin de feuille =============================== */
      if (cfg.cleanUnused) {
        const finalRow = lastDataRow + 1 + (cfg.calculateAndWriteStats?STATS_ROW_COUNT_INT:0);
        const maxUsedCol = Math.max(...Object.values(colMap).filter(c=>c>0));
        cleanUnusedRowsAndColumnsINT(sheet,finalRow,maxUsedCol,
                                     CLEANUP_BUFFER_ROWS_INT,
                                     CLEANUP_BUFFER_COLS_INT);
      }

    } catch (err) {
      bilan.errors.push(`Erreur ${name} : ${err.message}`);
    }
  }); // fin boucle feuilles

  /* ===================  RETOUR =================== */
  return {
    success : bilan.errors.length === 0,
    message : `Traité ${sheets.length} onglet(s) INT – ${bilan.errors.length} erreur(s).`,
    bilan   : bilan
  };
}

/* ====================  FIN PARTIE 3  ==================== */

/* =====================  PARTIE 4  (bilan + export)  =====================
 *  À placer APRÈS les parties 1, 2 et 3 – c’est la fin du module.
 * ======================================================================= */

/**
 * Crée ou met à jour l’onglet “BILAN INT” à partir du récapitulatif.
 *
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {Object} bilan  Objet produit par miseEnFormeOngletsINT()
 */
function createBilanINT(ss, bilan) {
  let sh = ss.getSheetByName('BILAN INT');
  if (sh) {
    sh.clear();
    sh.getCharts().forEach(c => sh.removeChart(c));
  } else {
    sh = ss.insertSheet('BILAN INT');
  }

  // --- titre ----------------------------------------------------------
  sh.getRange('A1').setValue('BILAN DES CLASSES INT')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.setRowHeight(1, 32);
  sh.getRange('A1:D1').merge();

  // --- pas de données -------------------------------------------------
  if (!bilan.classes.length) {
    sh.getRange('A3').setValue('Aucune donnée pour les classes INT.')
      .setFontStyle('italic');
    return;
  }

  // --- tableau effectifs & parité ------------------------------------
  const headers = ['CLASSE', 'TOTAL', 'FILLES', 'GARÇONS'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#ECEFF1');

  const rows = bilan.classes.map((classe, i) => {
    const filles  = bilan.genreCounts[i][0];
    const garcons = bilan.genreCounts[i][1];
    return [ classe, filles + garcons, filles, garcons ];
  });

  sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  sh.autoResizeColumns(1, headers.length);

  // --- petit histogramme ---------------------------------------------
  try {
    const chart = sh.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sh.getRange(3, 1, rows.length + 1, 4))
      .setOption('title', 'Effectifs par classe (INT)')
      .setOption('isStacked', true)
      .setPosition(2, 6, 0, 0)
      .build();
    sh.insertChart(chart);
  } catch (e) {
    Logger.log('Chart error: ' + e);
  }
}

/* ==================  EXPORT PUBLIC & FERMETURE  =================== */

/**
 * Fonction exposée : INTFormatter.run(options)
 *  - formate tous les onglets *INT*
 *  - crée le bilan s’il y a des données
 *  - renvoie l’objet résultat
 */
return {
  run: function (opts) {
    const res = miseEnFormeOngletsINT(opts);
    if (res.bilan && res.bilan.classes.length) {
      createBilanINT(SpreadsheetApp.getActiveSpreadsheet(), res.bilan);
    }
    return res;
  }
};

})();    //  ← FIN DÉFINITIVE DE L’IIFE INTFormatter



/* ------------------------------------------------------------------
 *  Point d’entrée GLOBAL (menu, déclencheur)
 * ------------------------------------------------------------------*/
function entryPointMiseEnFormeINT() {
  const outcome = INTFormatter.run();          // toutes options par défaut
  SpreadsheetApp.getUi().alert(outcome.message);
}


