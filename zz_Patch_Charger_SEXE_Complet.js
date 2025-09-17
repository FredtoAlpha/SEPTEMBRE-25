/**
 * ==================================================================
 *    FICHIER  : Patch_Charger_SEXE_Complet.gs
 *    VERSION  : 1.0
 *    OBJET    : Modifier chargerElevesEtClasses pour inclure SEXE
 * ==================================================================
 */

'use strict';

/**
 * Version PATCHÉE de chargerElevesEtClasses qui inclut le champ SEXE
 */
function chargerElevesEtClasses_AvecSEXE(config = null, mobiliteField = "MOBILITE") {
  const finalConfig = config || getConfig();
  const TEST_SUFFIX = finalConfig.TEST_SUFFIX || "TEST";
  
  // Headers standards PLUS SEXE
  const HEADER_NAMES = {
    ID_ELEVE: "ID_ELEVE",
    NOM: "NOM & PRENOM",
    SEXE: "SEXE",         // ← AJOUT CRITIQUE
    COM: "COM",
    TRA: "TRA",
    PART: "PART",
    ABS: "ABS",
    OPT: "OPT",
    LV2: "LV2",
    ASSO: "ASSO",
    DISSO: "DISSO",
    MOBILITE: mobiliteField || "MOBILITE"
  };
  
  Logger.log("--- Début chargerElevesEtClasses AVEC SEXE ---");
  Logger.log(`Headers configurés: ${JSON.stringify(HEADER_NAMES)}`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getTestSheets();
  
  if (sheets.length === 0) {
    Logger.log("ERREUR: Aucune feuille TEST trouvée");
    return { success: false, error: "ERR_NO_TEST_SHEETS", students: [], colIndexes: {} };
  }
  
  const allStudents = [];
  let globalColIndexes = {};
  
  sheets.forEach((sheet, sheetIdx) => {
    const sheetName = sheet.getName();
    Logger.log(`\nTraitement de ${sheetName}...`);
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    
    const headers = data[0];
    const colIndexes = mapHeaders(headers, HEADER_NAMES);
    
    // Vérifier que SEXE est bien mappé
    if (colIndexes.SEXE === undefined) {
      Logger.log(`⚠️ Colonne SEXE non trouvée dans ${sheetName}`);
    } else {
      Logger.log(`✓ Colonne SEXE trouvée: position ${colIndexes.SEXE + 1}`);
    }
    
    if (sheetIdx === 0) globalColIndexes = colIndexes;
    
    // Traiter chaque ligne
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const idEleve = colIndexes.ID_ELEVE !== undefined ? row[colIndexes.ID_ELEVE] : null;
      
      // PATCH CRUCIAL : On charge SEULEMENT si ID_ELEVE match bien "°" + chiffres en suffixe (ex: ECOLE°1234)
      if (idEleve && /°\d+$/.test(String(idEleve).trim())) {
        // Normal : c'est un élève
        const student = {
          CLASSE: sheetName,
          ID_ELEVE: String(idEleve).trim(),
          NOM: colIndexes.NOM !== undefined ? String(row[colIndexes.NOM] || "").trim() : "",
          SEXE: colIndexes.SEXE !== undefined ? normaliserSexeSimple(row[colIndexes.SEXE]) : "",
          COM: colIndexes.COM !== undefined ? parseFloat(row[colIndexes.COM]) || 0 : 0,
          TRA: colIndexes.TRA !== undefined ? parseFloat(row[colIndexes.TRA]) || 0 : 0,
          PART: colIndexes.PART !== undefined ? parseFloat(row[colIndexes.PART]) || 0 : 0,
          ABS: colIndexes.ABS !== undefined ? parseFloat(row[colIndexes.ABS]) || 0 : 0,
          OPT: colIndexes.OPT !== undefined ? String(row[colIndexes.OPT] || "").trim() : "",
          LV2: colIndexes.LV2 !== undefined ? String(row[colIndexes.LV2] || "").trim() : "",
          ASSO: colIndexes.ASSO !== undefined ? String(row[colIndexes.ASSO] || "").trim() : "",
          DISSO: colIndexes.DISSO !== undefined ? String(row[colIndexes.DISSO] || "").trim() : "",
          mobilite: colIndexes.MOBILITE !== undefined ? String(row[colIndexes.MOBILITE] || "LIBRE").trim() : "LIBRE"
        };
        
        // Normaliser l'option
        if (student.OPT) {
          student.optionKey = student.OPT.toUpperCase();
        }
        
        allStudents.push(student);
      } else {
        // Ligne ignorée = stats/vides/totaux
        Logger.log(`⏭️ Ligne ignorée (pas un élève): "${idEleve}"`);
      }
    }
    
    Logger.log(`🎯 ${allStudents.filter(s => s.CLASSE === sheetName).length} élèves chargés pour ${sheetName}`);
  });
  
  Logger.log(`--- Fin chargerElevesEtClasses (${allStudents.length} élèves, avec SEXE) ---`);
  
  // Statistiques rapides
  const nbF = allStudents.filter(s => s.SEXE === 'F').length;
  const nbM = allStudents.filter(s => s.SEXE === 'M').length;
  const nbAutre = allStudents.length - nbF - nbM;
  Logger.log(`Répartition globale: ${nbF}F, ${nbM}M, ${nbAutre} non définis`);
  
  return {
    success: true,
    students: allStudents,
    colIndexes: globalColIndexes
  };
}

/**
 * Fonction helper pour mapper les headers
 */
function mapHeaders(headers, headerNames) {
  const mapping = {};
  
  headers.forEach((header, idx) => {
    const h = String(header).trim().toUpperCase();
    
    Object.entries(headerNames).forEach(([key, value]) => {
      if (mapping[key] === undefined) {
        const target = String(value).toUpperCase();
        
        // Mapping exact ou partiel
        if (h === target || h.includes(target)) {
          mapping[key] = idx;
        }
        
        // Cas spéciaux pour SEXE
        if (key === 'SEXE' && mapping[key] === undefined) {
          const sexeVariants = ['SEXE', 'SEX', 'GENRE', 'H/F', 'G/F', 'M/F', 'F/M', 'F/G'];
          if (sexeVariants.some(v => h.includes(v))) {
            mapping[key] = idx;
          }
        }
      }
    });
  });
  
  return mapping;
}

/**
 * Normalisation simple du sexe
 */
function normaliserSexeSimple(raw) {
  if (!raw) return '';
  const val = String(raw).trim().toUpperCase();
  
  if (val === 'F' || val === 'FILLE' || val === 'FEMME' || val === 'FEM' || val === '2') return 'F';
  if (val === 'M' || val === 'H' || val === 'G' || val === 'GARCON' || val === 'HOMME' || val === 'MASC' || val === '1') return 'M';
  
  return val.charAt(0) === 'F' ? 'F' : val.charAt(0) === 'M' || val.charAt(0) === 'G' || val.charAt(0) === 'H' ? 'M' : '';
}

/**
 * Patch de V2_Ameliore_PreparerDonnees pour utiliser la version avec SEXE
 */
function V2_Ameliore_PreparerDonnees_AvecSEXE(config) {
  Logger.log("V2_Ameliore_PreparerDonnees_AvecSEXE: Utilisation du chargeur avec SEXE");
  
  const niveauActif = determinerNiveauActifCache();
  const structureResult = chargerStructureEtOptions(niveauActif, config);
  if (!structureResult.success) throw new Error("Échec chargement structure");
  
  const optionPools = buildOptionPools(structureResult.structure, config);
  Logger.log("Option pools: " + JSON.stringify(optionPools));
  
  // *** UTILISER LA VERSION PATCHÉE ***
  const chargeResult = chargerElevesEtClasses_AvecSEXE(config, "MOBILITE");
  if (!chargeResult.success) throw new Error("Échec chargement élèves");
  
  const { clean: elevesValides } = sanitizeStudents(chargeResult.students);
  classifierEleves(elevesValides, ['COM', 'TRA', 'PART', 'ABS']);
  
  // Organisation par classe
  const classesState = {};
  const effectifsClasses = {};
  
  elevesValides.forEach(eleve => {
    const classe = eleve.CLASSE;
    if (!classe) return;
    
    if (!classesState[classe]) classesState[classe] = [];
    classesState[classe].push(eleve);
  });
  
  Object.keys(classesState).forEach(cls => {
    effectifsClasses[cls] = classesState[cls].length;
    
    // Log rapide de la parité
    const nbF = classesState[cls].filter(e => e.SEXE === 'F').length;
    const nbM = classesState[cls].filter(e => e.SEXE === 'M').length;
    Logger.log(`${cls}: ${nbF}F/${nbM}M`);
  });
  
  // Reste du code identique...
  const distributionGlobale = V2_Ameliore_CalculerDistributionGlobale(elevesValides, ['COM', 'TRA', 'PART']);
  const totalEleves = elevesValides.length;
  const nbClasses = Object.keys(classesState).length;
  
  const ciblesParClasse = {};
  Object.keys(classesState).forEach(classe => {
    const effectifClasse = classesState[classe].length;
    ciblesParClasse[classe] = {};
    
    ['COM', 'TRA', 'PART'].forEach(critere => {
      ciblesParClasse[classe][critere] = {};
      ['1', '2', '3', '4'].forEach(score => {
        const totalScore = distributionGlobale[critere][score];
        const cible = Math.round((effectifClasse / totalEleves) * totalScore);
        ciblesParClasse[classe][critere][score] = cible;
      });
    });
  });
  
  const dataContext = {
    config: config,
    elevesValides: elevesValides,
    classesState: classesState,
    effectifsClasses: effectifsClasses,
    optionPools: optionPools,
    structureData: structureResult.structure,
    colIndexes: chargeResult.colIndexes,
    dissocMap: buildDissocCountMap(classesState),
    distributionGlobale: distributionGlobale,
    ciblesParClasse: ciblesParClasse,
    totalEleves: totalEleves,
    nbClasses: nbClasses
  };
  
  dataContext.totalEleves = elevesValides.length;
  dataContext.classeCaches = {};
  Object.keys(dataContext.classesState).forEach(cls => {
    dataContext.classeCaches[cls] = {
      dist: { COM: {1:0,2:0,3:0,4:0}, TRA: {1:0,2:0,3:0,4:0}, PART: {1:0,2:0,3:0,4:0} },
      parite: { F: 0, M: 0, total: 0 },
      score: 0
    };
  });
  dataContext.scoreGlobal = 0;
  
  return dataContext;
}

/**
 * Lancement de l'optimisation parité avec le bon chargeur
 */
function lancerPariteAgressive_AvecBonChargeur() {
  const ui = SpreadsheetApp.getUi();
  const heureDebut = new Date();
  
  try {
    Logger.log("\n=== OPTIMISATION PARITÉ AVEC CHARGEUR SEXE CORRIGÉ ===");
    
    const config = getConfig();
    
    // Configuration agressive pour la parité
    config.V2_POIDS_EQUILIBRE = {
      COM: 0.05,
      TRA: 0.05,
      PART: 0.05,
      PARITE: 0.85
    };
    
    config.TOLERANCE_EFFECTIF_MIN = 23;
    config.TOLERANCE_EFFECTIF_MAX = 29;
    config.ECART_PARITE_MAX_ACCEPTABLE = 3;
    
    // Utiliser le chargeur patché
    Logger.log("Chargement des données AVEC SEXE...");
    const dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
    
    // Vérifier que le SEXE est bien chargé
    Logger.log("\n=== VÉRIFICATION SEXE APRÈS CHARGEMENT ===");
    Object.keys(dataContext.classesState).forEach(classe => {
      const eleves = dataContext.classesState[classe];
      const nbF = eleves.filter(e => e.SEXE === 'F').length;
      const nbM = eleves.filter(e => e.SEXE === 'M').length;
      Logger.log(`${classe}: ${nbF}F, ${nbM}M (écart: ${Math.abs(nbF - nbM)})`);
    });
    
    // Maintenant lancer l'optimisation parité
    const analyseInitiale = _analyserPariteGlobale(dataContext);
    _afficherAnalyseParite("INITIALE", analyseInitiale);
    
    // Exécuter les phases d'optimisation
    const transfers = [];
    
    // Phase 1: Transferts directs
    Logger.log("\n=== PHASE 1: TRANSFERTS DIRECTS ===");
    transfers.push(..._executerTransfertsDirectsCorrigee(dataContext, config, analyseInitiale));
    
    // Phase 2: Échanges
    Logger.log("\n=== PHASE 2: ÉCHANGES ===");
    transfers.push(..._executerEchangesCorrigee(dataContext, config));
    
    // Analyse finale
    const analyseFinale = _analyserPariteGlobale(dataContext);
    _afficherAnalyseParite("FINALE", analyseFinale);
    
    // Application des modifications
    if (transfers.length > 0) {
      Logger.log(`\nApplication de ${transfers.length} modifications...`);
      const nbAppliques = V2_Ameliore_AppliquerSwaps_Flexible(transfers, dataContext, config);
      
      ui.alert('Optimisation Parité', 
        `✅ Optimisation terminée !\n\n` +
        `Modifications appliquées: ${nbAppliques}\n` +
        `Score parité: ${analyseInitiale.scoreGlobalParite.toFixed(1)} → ${analyseFinale.scoreGlobalParite.toFixed(1)}\n\n` +
        `Consultez les logs pour le détail`,
        ui.ButtonSet.OK
      );
      
      // Générer un bilan
      _genererBilanSimple(transfers);
      
    } else {
      ui.alert('Information', 'Aucune modification nécessaire ou possible.', ui.ButtonSet.OK);
    }
    
    return { success: true, nbModifications: transfers.length };
    
  } catch (e) {
    Logger.log(`ERREUR: ${e.message}\n${e.stack}`);
    ui.alert('Erreur', e.message, ui.ButtonSet.OK);
    return { success: false, error: e.message };
  }
}

/**
 * Phase 1 corrigée: Transferts directs
 */
function _executerTransfertsDirectsCorrigee(dataContext, config, analyse) {
  const transfers = [];
  
  // Classes avec gros écarts
  const classesProblematiques = analyse.classes.filter(c => 
    Math.abs(c.surplusGarcons) > config.ECART_PARITE_MAX_ACCEPTABLE
  );
  
  classesProblematiques.forEach(classeSource => {
    const eleves = dataContext.classesState[classeSource.classe];
    const genreSurplus = classeSource.surplusGarcons > 0 ? 'M' : 'F';
    
    // Chercher des élèves ESP mobiles du genre en surplus
    const elevesESPMobiles = eleves.filter(e => 
      e.OPT === 'ESP' && 
      e.mobilite === 'LIBRE' && 
      e.SEXE === genreSurplus
    );
    
    Logger.log(`${classeSource.classe}: ${elevesESPMobiles.length} élèves ESP ${genreSurplus} mobiles disponibles`);
    
    // Pour chaque élève ESP mobile
    elevesESPMobiles.forEach(eleve => {
      if (Math.abs(classeSource.surplusGarcons) <= config.ECART_PARITE_MAX_ACCEPTABLE) {
        return; // Cette classe est maintenant OK
      }
      
      // Chercher la meilleure classe cible
      const meilleureClasse = _trouverMeilleureClasseCibleCorrigee(
        eleve, classeSource, analyse, dataContext, config
      );
      
      if (meilleureClasse) {
        transfers.push({
          type: 'TRANSFERT',
          eleve1: eleve,
          eleve2: eleve,
          eleve1ID: eleve.ID_ELEVE,
          eleve2ID: eleve.ID_ELEVE,
          classe1: classeSource.classe,
          classe2: meilleureClasse,
          motif: `Transfert-Parité-${genreSurplus}-ESP`
        });
        
        // Simuler le transfert
        _appliquerTransfertDansContext(eleve, classeSource.classe, meilleureClasse, dataContext);
        
        // Mettre à jour les surplus
        if (genreSurplus === 'M') {
          classeSource.surplusGarcons--;
        } else {
          classeSource.surplusGarcons++;
        }
        
        Logger.log(`✓ Transfert planifié: ${eleve.NOM} (${genreSurplus}-ESP) de ${classeSource.classe} → ${meilleureClasse}`);
      }
    });
  });
  
  return transfers;
}

/**
 * Phase 2 corrigée: Échanges
 */
function _executerEchangesCorrigee(dataContext, config) {
  const echanges = [];
  const analyse = _analyserPariteGlobale(dataContext);
  
  // Classes avec déséquilibres opposés
  const classesGarconsPlus = analyse.classes.filter(c => c.surplusGarcons > 2);
  const classesFillesPlus = analyse.classes.filter(c => c.surplusGarcons < -2);
  
  Logger.log(`Classes avec surplus garçons: ${classesGarconsPlus.map(c => c.classe).join(', ')}`);
  Logger.log(`Classes avec surplus filles: ${classesFillesPlus.map(c => c.classe).join(', ')}`);
  
  classesGarconsPlus.forEach(classeG => {
    classesFillesPlus.forEach(classeF => {
      // Chercher des paires échangeables
      const garconsESP = dataContext.classesState[classeG.classe].filter(e => 
        e.OPT === 'ESP' && e.mobilite === 'LIBRE' && e.SEXE === 'M'
      );
      const fillesESP = dataContext.classesState[classeF.classe].filter(e => 
        e.OPT === 'ESP' && e.mobilite === 'LIBRE' && e.SEXE === 'F'
      );
      
      const nbEchanges = Math.min(
        garconsESP.length, 
        fillesESP.length,
        Math.ceil(classeG.surplusGarcons / 2),
        Math.ceil(Math.abs(classeF.surplusGarcons) / 2)
      );
      
      for (let i = 0; i < nbEchanges && i < 2; i++) { // Max 2 échanges par paire de classes
        if (garconsESP[i] && fillesESP[i]) {
          echanges.push({
            type: 'ECHANGE',
            eleve1: garconsESP[i],
            eleve2: fillesESP[i],
            eleve1ID: garconsESP[i].ID_ELEVE,
            eleve2ID: fillesESP[i].ID_ELEVE,
            classe1: classeG.classe,
            classe2: classeF.classe,
            motif: 'Échange-Parité-ESP-M↔F'
          });
          
          Logger.log(`✓ Échange planifié: ${garconsESP[i].NOM}(M) ↔ ${fillesESP[i].NOM}(F)`);
        }
      }
    });
  });
  
  return echanges;
}

function _trouverMeilleureClasseCibleCorrigee(eleve, classeSource, analyse, dataContext, config) {
  let meilleureClasse = null;
  let meilleurScore = -999;
  
  analyse.classes.forEach(classeCible => {
    if (classeCible.classe === classeSource.classe) return;
    if (classeCible.effectif >= config.TOLERANCE_EFFECTIF_MAX) return;
    
    // Vérifier que l'élève peut aller dans cette classe (option)
    if (eleve.OPT && eleve.OPT !== 'ESP') {
      const pool = dataContext.optionPools[eleve.OPT];
      if (!pool || !pool.includes(classeCible.classe.toUpperCase())) return;
    }
    
    // Calculer le bénéfice du transfert
    let score = 0;
    if (eleve.SEXE === 'M' && classeCible.surplusGarcons < 0) {
      score = Math.abs(classeCible.surplusGarcons) * 10;
    } else if (eleve.SEXE === 'F' && classeCible.surplusGarcons > 0) {
      score = classeCible.surplusGarcons * 10;
    }
    
    // Bonus si la classe cible est très déséquilibrée
    if (Math.abs(classeCible.surplusGarcons) > 4) {
      score += 50;
    }
    
    if (score > meilleurScore) {
      meilleurScore = score;
      meilleureClasse = classeCible.classe;
    }
  });
  
  return meilleureClasse;
}

function _genererBilanSimple(transfers) {
  const nomFeuille = "_BILAN_PARITE_SIMPLE";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(nomFeuille);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(nomFeuille);
  }
  
  const data = [
    ['=== BILAN CORRECTION PARITÉ ==='],
    [`Total opérations: ${transfers.length}`],
    [''],
    ['#', 'Type', 'Élève', 'Sexe', 'De', '→', 'Vers', 'Motif']
  ];
  
  transfers.forEach((t, idx) => {
    data.push([
      idx + 1,
      t.type,
      t.eleve1.NOM || t.eleve1.ID_ELEVE,
      t.eleve1.SEXE,
      t.classe1,
      '→',
      t.classe2,
      t.motif
    ]);
  });
  
  sheet.getRange(1, 1, data.length, 8).setValues(data.map(row => {
    while (row.length < 8) row.push('');
    return row.slice(0, 8);
  }));
  
  sheet.autoResizeColumns(1, 8);
}

/**
 * Test complet
 */
function testPariteAvecChargeurCorrige() {
  Logger.log("=== TEST PARITÉ AVEC CHARGEUR CORRIGÉ ===");
  
  const resultat = lancerPariteAgressive_AvecBonChargeur();
  
  if (resultat.success) {
    Logger.log(`✅ Test réussi ! ${resultat.nbModifications} modifications`);
  } else {
    Logger.log("❌ Test échoué: " + resultat.error);
  }
}

/********************************************************************
 *  SURCHARGE GLOBALE  – dernières lignes du fichier                *
 *  (garantit que TOUT le projet appelle bien la version « SEXE »)   *
 ********************************************************************/

// ① Chargeur d’élèves : on remplace la version legacy
function chargerElevesEtClasses(config, mobiliteField) {
  return chargerElevesEtClasses_AvecSEXE(config, mobiliteField);
}

// ② Préparation des données V2 : on remplace la version legacy
function V2_Ameliore_PreparerDonnees(config) {
  return V2_Ameliore_PreparerDonnees_AvecSEXE(config);
}

// ③ Parité agressive : alias pour rester compatible avec les menus
function lancerPariteAgressive() {
  return lancerPariteAgressive_AvecBonChargeur();
}
/********************************************************************
 *  SURCHARGE DES FONCTIONS ORIGINELLES AVEC LA VERSION « SEXE »
 *  (à placer tout en bas de Patch_Charger_SEXE_Complet.gs)
 ********************************************************************/

// remplace chargerElevesEtClasses() par la version qui gère SEXE
globalThis.chargerElevesEtClasses = chargerElevesEtClasses_AvecSEXE;

// remplace V2_Ameliore_PreparerDonnees() par la version qui gère SEXE
globalThis.V2_Ameliore_PreparerDonnees = V2_Ameliore_PreparerDonnees_AvecSEXE;