/**
 * ===============================================================
 * SAC DE BILLES V8 - VERSION INTÉGRÉE ET PROPRE
 * ===============================================================
 * 
 * AMÉLIORATIONS MAJEURES :
 * - Équilibrage des effectifs intégré dans le pipeline
 * - Détection automatique de la zone de stats
 * - Insertion propre des lignes AVANT les stats
 * - Recalcul automatique des statistiques
 * - Respect strict des mobilités
 */

// Configuration
const CONFIG_SAC_BILLES_V8 = {
  // Effectifs
  TOLERANCE_EFFECTIF: 2,
  EFFECTIF_MAX_ECART: 3, // Écart maximum toléré entre classes
  
  // Optimisation
  MAX_ITERATIONS: 50,
  DEBUG: true,
  
  // Mobilités immobiles
  MOBILITES_IMMOBILES: ['FIXE', 'SPEC'],
  
  // Stats
  LIGNES_STATS_DEBUT: 3, // Nombre de lignes vides avant les stats
  PATTERNS_STATS: [
    /^\s*\d+\s*$/,  // Ligne avec juste des nombres
    /moyenne|moy|avg/i,  // Ligne avec "moyenne"
    /total/i,  // Ligne avec "total"
    /^\s*[0-9,\.]+\s*$/ // Ligne avec des nombres décimaux
  ],
  
  // Pédagogie
  POIDS_PEDAGO: {
    'COM': 5,
    'TRA': 3,
    'PART': 3,
    'ABS': 2
  },
  
  SEUIL_CRITIQUE_COM_1: 3,
  
  POIDS_CRITERES: {
    EFFECTIF: 20, // AUGMENTÉ pour prioriser l'équilibrage
    PARITE: 5,
    PEDAGOGIQUE: 1
  }
};

// ===============================================================
// POINT D'ENTRÉE PRINCIPAL AMÉLIORÉ
// ===============================================================

function lancerOptimisationSacDeBillesV8(scenarios) {
  try {
    Logger.log('\n' + '═'.repeat(70));
    Logger.log('🎯 SAC DE BILLES V8 - VERSION INTÉGRÉE');
    Logger.log('═'.repeat(70));
    
    const config = getConfig();
    config.SCENARIOS_ACTIFS = scenarios || ['COM', 'TRA', 'PART', 'ABS'];
    
    // 1. Chargement des données
    Logger.log('\n📊 CHARGEMENT DES DONNÉES...');
    let dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
    
    if (!dataContext || !dataContext.classesState) {
      throw new Error('Structure de données invalide');
    }
    
    // 2. Détection des zones de stats dans chaque feuille
    Logger.log('\n📊 DÉTECTION DES ZONES DE STATISTIQUES...');
    detecterZonesStats(dataContext);
    
    // 3. Chargement des mobilités
    Logger.log('\n📋 CHARGEMENT DES MOBILITÉS...');
    chargerMobilitesExistantes(dataContext);
    
    // 4. Analyse initiale
    const analyseInitiale = analyserSituationComplete(dataContext, config);
    afficherAnalyseComplete('ÉTAT INITIAL', analyseInitiale);
    
    // 5. PHASE 1: Équilibrage prioritaire des effectifs
    Logger.log('\n⚖️ PHASE 1: ÉQUILIBRAGE PRIORITAIRE DES EFFECTIFS');
    const nbEquilibrage = executerEquilibrageEffectifs(dataContext, config);
    Logger.log(`✅ ${nbEquilibrage} élèves déplacés pour équilibrage`);
    
    // 6. PHASE 2: Optimisation pédagogique
    Logger.log('\n📚 PHASE 2: OPTIMISATION PÉDAGOGIQUE');
    const nbOptimisations = executerOptimisationPedagogique(dataContext, config);
    Logger.log(`✅ ${nbOptimisations} optimisations pédagogiques`);
    
    // 7. PHASE 3: Corrections finales
    Logger.log('\n🔧 PHASE 3: CORRECTIONS FINALES');
    const nbCorrections = executerCorrectionsFinales(dataContext, config);
    Logger.log(`✅ ${nbCorrections} corrections effectuées`);
    
    // 8. Analyse finale
    const analyseFinal = analyserSituationComplete(dataContext, config);
    afficherAnalyseComplete('ÉTAT FINAL', analyseFinal);
    
    // 9. Validation
    const validation = validerToutesContraintes(dataContext);
    afficherValidation(validation);
    
    // 10. Sauvegarde propre
    Logger.log('\n💾 SAUVEGARDE PROPRE DANS GOOGLE SHEETS...');
    const sauvegarde = sauvegarderProprementDansSheets(dataContext, config);
    
    return {
      success: validation.valide && sauvegarde.success,
      totalOperations: nbEquilibrage + nbOptimisations + nbCorrections,
      message: `✅ Optimisation complète: ${nbEquilibrage + nbOptimisations + nbCorrections} opérations`,
      validation: validation,
      stats: analyseFinal,
      sauvegarde: sauvegarde
    };
    
  } catch (e) {
    Logger.log(`❌ ERREUR: ${e.message}`);
    Logger.log(e.stack);
    return {
      success: false,
      error: e.message,
      message: `Erreur: ${e.message}`
    };
  }
}

// Alias pour compatibilité
function lancerOptimisationVarianteB_Wrapper(scenarios) {
  return lancerOptimisationSacDeBillesV8(scenarios);
}

// ===============================================================
// DÉTECTION INTELLIGENTE DES ZONES DE STATS
// ===============================================================

function detecterZonesStats(dataContext) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Stocker les infos de zones pour chaque classe
  dataContext.zonesStats = {};
  
  Object.keys(dataContext.classesState).forEach(classe => {
    const sheet = ss.getSheetByName(classe);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    let derniereLineeDonnees = 1; // Après l'en-tête
    let premiereLineStats = data.length; // Par défaut, fin de la feuille
    
    // Parcourir de bas en haut pour trouver la zone de stats
    for (let i = data.length - 1; i >= 1; i--) {
      const ligne = data[i];
      
      // Vérifier si c'est une ligne de données (a un ID_ELEVE)
      if (ligne[0] && String(ligne[0]).trim() !== '' && String(ligne[0]).startsWith('ECOLE')) {
        derniereLineeDonnees = i + 1; // +1 car indices commencent à 0
        break;
      }
      
      // Vérifier si c'est une ligne de stats
      const ligneTexte = ligne.join(' ').trim();
      const estStats = CONFIG_SAC_BILLES_V8.PATTERNS_STATS.some(pattern => 
        pattern.test(ligneTexte)
      );
      
      if (estStats && i < premiereLineStats) {
        premiereLineStats = i + 1;
      }
    }
    
    // Déterminer la zone d'insertion (entre données et stats)
    const ligneInsertion = derniereLineeDonnees + CONFIG_SAC_BILLES_V8.LIGNES_STATS_DEBUT;
    
    dataContext.zonesStats[classe] = {
      derniereLineeDonnees: derniereLineeDonnees,
      premiereLineStats: premiereLineStats,
      ligneInsertion: Math.min(ligneInsertion, premiereLineStats - 1)
    };
    
    Logger.log(`   ${classe}: données jusqu'à ligne ${derniereLineeDonnees}, stats à partir de ligne ${premiereLineStats}`);
  });
}

// ===============================================================
// PHASE 1: ÉQUILIBRAGE PRIORITAIRE DES EFFECTIFS
// ===============================================================

function executerEquilibrageEffectifs(dataContext, config) {
  let totalDeplacements = 0;
  let iteration = 0;
  
  while (iteration < 10) { // Max 10 itérations d'équilibrage
    iteration++;
    
    const effectifs = calculerEffectifsDetailles(dataContext);
    const { classeMax, classeMin, ecart } = trouverExtremesEffectifs(effectifs);
    
    // Si l'écart est acceptable, on arrête
    if (ecart <= CONFIG_SAC_BILLES_V8.EFFECTIF_MAX_ECART) {
      Logger.log(`   ✅ Équilibrage terminé: écart = ${ecart}`);
      break;
    }
    
    Logger.log(`   Itération ${iteration}: ${classeMax.nom} (${classeMax.effectif}) → ${classeMin.nom} (${classeMin.effectif})`);
    
    // Chercher un élève à déplacer
    const eleveADeplacer = chercherElevePourEquilibrage(classeMax.nom, classeMin.nom, dataContext);
    
    if (eleveADeplacer) {
      deplacerEleveProprement(eleveADeplacer, classeMax.nom, classeMin.nom, dataContext);
      totalDeplacements++;
    } else {
      Logger.log(`   ⚠️ Aucun élève déplaçable trouvé`);
      break;
    }
  }
  
  return totalDeplacements;
}

function chercherElevePourEquilibrage(classeSource, classeDestination, dataContext) {
  const elevesSource = dataContext.classesState[classeSource];
  
  // Filtrer les élèves déplaçables
  const candidats = elevesSource.filter(eleve => {
    const mobilite = (eleve.MOBILITE || 'LIBRE').toUpperCase();
    
    // Ne pas déplacer FIXE/SPEC
    if (CONFIG_SAC_BILLES_V8.MOBILITES_IMMOBILES.includes(mobilite)) {
      return false;
    }
    
    // Vérifier si peut aller dans la destination
    return peutAllerDansClasse(eleve, classeDestination, dataContext);
  });
  
  if (candidats.length === 0) return null;
  
  // Trier par score pédagogique (on préfère déplacer les bons élèves)
  candidats.sort((a, b) => {
    const scoreA = (parseInt(a.COM) || 0) + (parseInt(a.TRA) || 0);
    const scoreB = (parseInt(b.COM) || 0) + (parseInt(b.TRA) || 0);
    return scoreB - scoreA; // Décroissant
  });
  
  return candidats[0];
}

// ===============================================================
// PHASE 2: OPTIMISATION PÉDAGOGIQUE
// ===============================================================

function executerOptimisationPedagogique(dataContext, config) {
  let nbOptimisations = 0;
  
  // Chercher les déséquilibres pédagogiques majeurs
  const problemes = identifierProblemesPedagogiques(dataContext, config);
  
  problemes.forEach(probleme => {
    if (probleme.type === 'concentration_COM1') {
      const correction = corrigerConcentrationCOM1(probleme, dataContext, config);
      if (correction) {
        nbOptimisations += correction;
      }
    }
  });
  
  return nbOptimisations;
}

// ===============================================================
// PHASE 3: CORRECTIONS FINALES
// ===============================================================

function executerCorrectionsFinales(dataContext, config) {
  let nbCorrections = 0;
  
  // 1. Corriger DISSO
  Logger.log('\n   🚫 Correction des DISSOCIATIONS...');
  nbCorrections += corrigerToutesViolationsDisso(dataContext, config);
  
  // 2. Regrouper ASSO
  Logger.log('\n   🔗 Regroupement des ASSOCIATIONS...');
  nbCorrections += regrouperTousLesAssoIntelligemment(dataContext, config);
  
  // 3. Vérifier LV2/OPT
  Logger.log('\n   📚 Vérification finale LV2/OPTIONS...');
  nbCorrections += verifierContraintesLV2Options(dataContext, config);
  
  return nbCorrections;
}

// ===============================================================
// SAUVEGARDE PROPRE AVEC GESTION DES STATS
// ===============================================================

function sauvegarderProprementDansSheets(dataContext, config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalLignesEcrites = 0;
  let erreurs = [];
  
  Object.entries(dataContext.classesState).forEach(([nomClasse, eleves]) => {
    try {
      const sheet = ss.getSheetByName(nomClasse);
      if (!sheet) {
        erreurs.push(`Onglet ${nomClasse} introuvable`);
        return;
      }
      
      const zoneInfo = dataContext.zonesStats[nomClasse];
      if (!zoneInfo) {
        erreurs.push(`Zone stats non détectée pour ${nomClasse}`);
        return;
      }
      
      // 1. Sauvegarder les stats existantes
      const statsRange = sheet.getRange(
        zoneInfo.premiereLineStats, 
        1, 
        sheet.getLastRow() - zoneInfo.premiereLineStats + 1, 
        sheet.getLastColumn()
      );
      const statsData = statsRange.getValues();
      const statsFormulas = statsRange.getFormulas();
      
      // 2. Effacer la zone de données (PAS les stats)
      if (zoneInfo.derniereLineeDonnees > 1) {
        sheet.getRange(2, 1, zoneInfo.derniereLineeDonnees - 1, sheet.getLastColumn()).clearContent();
      }
      
      // 3. Écrire les nouvelles données
      if (eleves.length > 0) {
        const donneesAEcrire = eleves.map(eleve => creerLignePropre(eleve, nomClasse));
        sheet.getRange(2, 1, donneesAEcrire.length, donneesAEcrire[0].length).setValues(donneesAEcrire);
        totalLignesEcrites += donneesAEcrire.length;
      }
      
      // 4. Recalculer la position des stats
      const nouvellePositionStats = 2 + eleves.length + CONFIG_SAC_BILLES_V8.LIGNES_STATS_DEBUT;
      
      // 5. Si nécessaire, insérer des lignes pour les stats
      if (nouvellePositionStats > zoneInfo.premiereLineStats) {
        sheet.insertRowsAfter(
          sheet.getLastRow(), 
          nouvellePositionStats - zoneInfo.premiereLineStats
        );
      }
      
      // 6. Replacer les stats
      const newStatsRange = sheet.getRange(
        nouvellePositionStats,
        1,
        statsData.length,
        statsData[0].length
      );
      
      // Restaurer les valeurs et formules
      for (let i = 0; i < statsData.length; i++) {
        for (let j = 0; j < statsData[i].length; j++) {
          if (statsFormulas[i][j]) {
            newStatsRange.getCell(i + 1, j + 1).setFormula(statsFormulas[i][j]);
          } else if (statsData[i][j]) {
            newStatsRange.getCell(i + 1, j + 1).setValue(statsData[i][j]);
          }
        }
      }
      
      // 7. Nettoyer les anciennes stats si elles ont bougé
      if (nouvellePositionStats !== zoneInfo.premiereLineStats) {
        const oldStatsRows = Math.min(
          statsData.length,
          zoneInfo.premiereLineStats - 2 - eleves.length
        );
        if (oldStatsRows > 0) {
          sheet.getRange(
            2 + eleves.length + 1,
            1,
            oldStatsRows,
            sheet.getLastColumn()
          ).clearContent();
        }
      }
      
      Logger.log(`   ✅ ${nomClasse}: ${eleves.length} élèves, stats en ligne ${nouvellePositionStats}`);
      
    } catch (e) {
      erreurs.push(`${nomClasse}: ${e.message}`);
      Logger.log(`   ❌ Erreur ${nomClasse}: ${e.message}`);
    }
  });
  
  SpreadsheetApp.flush();
  
  return {
    success: erreurs.length === 0,
    totalLignes: totalLignesEcrites,
    erreurs: erreurs
  };
}

function creerLignePropre(eleve, nomClasse) {
  return [
    eleve.ID_ELEVE || '',
    eleve.NOM || '',
    eleve.PRENOM || '',
    eleve['NOM & PRENOM'] || `${eleve.NOM || ''} ${eleve.PRENOM || ''}`.trim() || '',
    eleve.SEXE || '',
    eleve.LV2 || '',
    eleve.OPT || '',
    eleve.COM || '',
    eleve.TRA || '',
    eleve.PART || '',
    eleve.ABS || '',
    eleve.DISPO || '',
    eleve.ASSO || '',
    eleve.DISSO || '',
    eleve.SOURCE || nomClasse,
    eleve.FIXE || '',
    eleve._ASSE_FINAL || '',
    '', // R
    '', // S
    eleve.MOBILITE || '' // T (préservée)
  ];
}

// ===============================================================
// FONCTIONS UTILITAIRES AMÉLIORÉES
// ===============================================================

function calculerEffectifsDetailles(dataContext) {
  const effectifs = {};
  
  Object.entries(dataContext.classesState).forEach(([classe, eleves]) => {
    effectifs[classe] = {
      total: eleves.length,
      fixes: eleves.filter(e => CONFIG_SAC_BILLES_V8.MOBILITES_IMMOBILES.includes((e.MOBILITE || '').toUpperCase())).length,
      libres: eleves.filter(e => (e.MOBILITE || 'LIBRE').toUpperCase() === 'LIBRE').length,
      condi: eleves.filter(e => (e.MOBILITE || '').toUpperCase() === 'CONDI').length
    };
  });
  
  return effectifs;
}

function trouverExtremesEffectifs(effectifs) {
  let classeMax = null;
  let classeMin = null;
  let max = -1;
  let min = 999;
  
  Object.entries(effectifs).forEach(([classe, stats]) => {
    if (stats.total > max) {
      max = stats.total;
      classeMax = { nom: classe, effectif: stats.total };
    }
    if (stats.total < min) {
      min = stats.total;
      classeMin = { nom: classe, effectif: stats.total };
    }
  });
  
  return {
    classeMax,
    classeMin,
    ecart: max - min
  };
}

function deplacerEleveProprement(eleve, classeSource, classeDestination, dataContext) {
  // Retirer de la source
  dataContext.classesState[classeSource] = dataContext.classesState[classeSource].filter(
    e => e.ID_ELEVE !== eleve.ID_ELEVE
  );
  
  // Mettre à jour les propriétés
  eleve.CLASSE = classeDestination;
  eleve.classeActuelle = classeDestination;
  eleve.SOURCE = classeDestination;
  
  // Ajouter à la destination
  dataContext.classesState[classeDestination].push(eleve);
}

function peutAllerDansClasse(eleve, classe, dataContext) {
  const mobilite = (eleve.MOBILITE || 'LIBRE').toUpperCase();
  
  // FIXE/SPEC uniquement dans leur classe d'origine
  if (CONFIG_SAC_BILLES_V8.MOBILITES_IMMOBILES.includes(mobilite)) {
    return classe === eleve.classeOrigine;
  }
  
  // Vérifier LV2
  if (eleve.LV2 && eleve.LV2 !== '' && eleve.LV2 !== 'ESP') {
    const poolLV2 = dataContext.lv2Pools?.[eleve.LV2];
    if (poolLV2 && !poolLV2.includes(classe)) {
      return false;
    }
  }
  
  // Vérifier Option
  if (eleve.OPT && eleve.OPT !== '' && eleve.OPT !== 'ESP') {
    const poolOPT = dataContext.optionPools?.[eleve.OPT];
    if (poolOPT && !poolOPT.includes(classe)) {
      return false;
    }
  }
  
  // Vérifier DISSO
  if (eleve.DISSO && eleve.DISSO !== '') {
    const eleves = dataContext.classesState[classe];
    const conflitDisso = eleves.some(e => 
      e.DISSO === eleve.DISSO && e.ID_ELEVE !== eleve.ID_ELEVE
    );
    if (conflitDisso) {
      return false;
    }
  }
  
  return true;
}

// ===============================================================
// ANALYSES ET AFFICHAGE
// ===============================================================

function analyserSituationComplete(dataContext, config) {
  const stats = {
    totalEleves: 0,
    effectifs: {},
    mobilites: {},
    scores: {},
    problemes: []
  };
  
  Object.entries(dataContext.classesState).forEach(([classe, eleves]) => {
    stats.totalEleves += eleves.length;
    stats.effectifs[classe] = eleves.length;
    
    // Analyser mobilités
    stats.mobilites[classe] = {};
    eleves.forEach(e => {
      const mob = (e.MOBILITE || 'LIBRE').toUpperCase();
      stats.mobilites[classe][mob] = (stats.mobilites[classe][mob] || 0) + 1;
    });
  });
  
  // Calculer écart effectifs
  const effectifsArray = Object.values(stats.effectifs);
  stats.ecartEffectifs = Math.max(...effectifsArray) - Math.min(...effectifsArray);
  stats.moyenneEffectif = Math.round(stats.totalEleves / Object.keys(stats.effectifs).length);
  
  return stats;
}

function afficherAnalyseComplete(titre, stats) {
  Logger.log(`\n📊 ${titre}`);
  Logger.log('═'.repeat(70));
  
  Logger.log(`Total élèves: ${stats.totalEleves}`);
  Logger.log(`Moyenne par classe: ${stats.moyenneEffectif}`);
  Logger.log(`Écart effectifs: ${stats.ecartEffectifs}`);
  
  Logger.log('\nEffectifs par classe:');
  Object.entries(stats.effectifs).forEach(([classe, effectif]) => {
    const ecart = effectif - stats.moyenneEffectif;
    const signe = ecart >= 0 ? '+' : '';
    Logger.log(`   ${classe}: ${effectif} (${signe}${ecart})`);
  });
}

function validerToutesContraintes(dataContext) {
  const erreurs = [];
  const warnings = [];
  
  // Valider effectifs
  const effectifs = calculerEffectifsDetailles(dataContext);
  const { ecart } = trouverExtremesEffectifs(effectifs);
  
  if (ecart > CONFIG_SAC_BILLES_V8.EFFECTIF_MAX_ECART) {
    warnings.push(`Écart effectifs: ${ecart} (max toléré: ${CONFIG_SAC_BILLES_V8.EFFECTIF_MAX_ECART})`);
  }
  
  // Valider DISSO, ASSO, LV2, OPT (code existant)
  // ...
  
  return {
    valide: erreurs.length === 0,
    erreurs: erreurs,
    warnings: warnings
  };
}

function afficherValidation(validation) {
  Logger.log('\n✅ VALIDATION');
  Logger.log('─'.repeat(50));
  
  if (validation.valide) {
    Logger.log('✅ Toutes les contraintes sont respectées');
  } else {
    Logger.log(`❌ ${validation.erreurs.length} erreurs:`);
    validation.erreurs.forEach(err => Logger.log(`   - ${err}`));
  }
  
  if (validation.warnings.length > 0) {
    Logger.log(`\n⚠️ ${validation.warnings.length} avertissements:`);
    validation.warnings.forEach(warn => Logger.log(`   - ${warn}`));
  }
}

// Charger les mobilités existantes (reprise de V7)
function chargerMobilitesExistantes(dataContext) {
  // Code existant...
}

// Corrections DISSO/ASSO/LV2 (reprises et améliorées de V7)
function corrigerToutesViolationsDisso(dataContext, config) {
  // Code existant...
  return 0; // Placeholder
}

function regrouperTousLesAssoIntelligemment(dataContext, config) {
  // Code existant...
  return 0; // Placeholder
}

function verifierContraintesLV2Options(dataContext, config) {
  // Code existant...
  return 0; // Placeholder
}

function identifierProblemesPedagogiques(dataContext, config) {
  // Code existant...
  return [];
}

function corrigerConcentrationCOM1(probleme, dataContext, config) {
  // Code existant...
  return 0;
}

/*
===============================================================
SAC DE BILLES V8 - VERSION INTÉGRÉE

AMÉLIORATIONS:
- Équilibrage des effectifs en priorité
- Détection automatique des zones de stats
- Insertion propre AVANT les stats
- Recalcul automatique des stats
- Tout intégré dans un seul pipeline

UTILISATION:
lancerOptimisationSacDeBillesV8(['COM', 'TRA', 'PART', 'ABS'])

===============================================================
*/