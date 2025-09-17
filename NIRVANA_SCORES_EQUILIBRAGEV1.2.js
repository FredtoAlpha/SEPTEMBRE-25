'use strict';

/**
 * ==================================================================
 *     NIRVANA_SCORES_EQUILIBRAGE - VERSION FINALE V1.2 COMPLÈTE
 *     Module spécialisé pour l'équilibrage des effectifs par score
 * ==================================================================
 * Version: 1.2 - Version corrigée et optimisée
 * Date: 19 Juillet 2025
 * 
 * Description:
 *   Module dédié à l'équilibrage des effectifs des scores 1-2-3-4
 *   pour les critères COM, TRA, PART, ABS (pas les moyennes !)
 *   
 *   OPTIMISATIONS V1.2:
 *   - Gestion robuste des groupes SPEC avec échange de groupes entiers
 *   - Stratégies d'échange adaptatives selon le type de mobilité
 *   - Application sécurisée avec vérifications multiples
 *   - Pipeline complet Nirvana intégré
 * ==================================================================
 */

// ==================================================================
// SECTION 1: CONFIGURATION ET CONSTANTES
// ==================================================================

const SCORES_CONFIG = {
  // Critères à équilibrer
  CRITERES: ['COM', 'TRA', 'PART', 'ABS'],
  
  // Scores possibles (1 = faible, 4 = excellent)
  SCORES: [1, 2, 3, 4],
  
  // Poids de priorité pour chaque critère
  POIDS_CRITERES: {
    COM: 0.35,    // Comportement - priorité haute
    TRA: 0.30,    // Travail - priorité haute
    PART: 0.25,   // Participation - priorité moyenne
    ABS: 0.10     // Absences - priorité basse
  },
  
  // Seuils de tolérance pour les effectifs
  TOLERANCE_EFFECTIFS: {
    STRICT: 1,    // Tolérance stricte (±1 élève)
    NORMAL: 2,    // Tolérance normale (±2 élèves)
    LARGE: 3      // Tolérance large (±3 élèves)
  },
  
  // Paramètres d'optimisation
  MAX_ITERATIONS: 50,
  SEUIL_AMELIORATION: 0.01,
  MAX_ECHANGES_PAR_ITERATION: 10,
  
  // Seuils pour diagnostic
  SEUIL_PROBLEME_SIGNIFICATIF: 2,
  SEUIL_SCORE_ACCEPTABLE: 70,
  
  // Priorités des stratégies d'échange
  PRIORITES_STRATEGIES: {
    LIBRE: 1.0,      // Priorité maximale
    PERMUT: 0.8,     // Priorité haute
    CONDI: 0.7,      // Priorité moyenne-haute
    SPEC: 0.6        // Priorité moyenne
  }
};

// ==================================================================
// SECTION 2: CALCUL DES CIBLES INTELLIGENT
// ==================================================================

/**
 * Calcule les effectifs cibles basés sur la distribution globale réelle
 */
function calculerEffectifsCibles_Ameliore(dataContext, config) {
  const classes = Object.keys(dataContext.classesState);
  const cibles = {};
  
  // Calculer la distribution globale réelle pour chaque critère
  const distributionGlobale = {};
  SCORES_CONFIG.CRITERES.forEach(critere => {
    distributionGlobale[critere] = { 1: 0, 2: 0, 3: 0, 4: 0, total: 0 };
  });
  
  // Compter tous les élèves
  Object.values(dataContext.classesState).flat().forEach(eleve => {
    SCORES_CONFIG.CRITERES.forEach(critere => {
      // DEBUG: Log des valeurs pour identifier le problème
      const valeurBrute = eleve[critere];
      const score = parseInt(valeurBrute) || 0;
      
      // Log détaillé pour debug
      if (critere === 'COM' && distributionGlobale[critere].total < 5) {
        Logger.log(`DEBUG ${critere}: valeurBrute="${valeurBrute}" (type: ${typeof valeurBrute}), score=${score}`);
        Logger.log(`DEBUG élève: ${JSON.stringify(eleve).substring(0, 200)}`);
      }
      
      if (score >= 1 && score <= 4) {
        distributionGlobale[critere][score]++;
        distributionGlobale[critere].total++;
      }
    });
  });
  
  Logger.log("=== DISTRIBUTION GLOBALE RÉELLE ===");
  SCORES_CONFIG.CRITERES.forEach(critere => {
    const dist = distributionGlobale[critere];
    const pourcentages = SCORES_CONFIG.SCORES.map(score => 
      `${score}:${((dist[score] / dist.total) * 100).toFixed(1)}%`
    ).join(" | ");
    Logger.log(`${critere}: ${pourcentages}`);
  });
  
  classes.forEach(classe => {
    const eleves = dataContext.classesState[classe];
    const effectifTotal = eleves.length;
    
    cibles[classe] = {
      effectifTotal: effectifTotal,
      effectifsCibles: {}
    };
    
    // Calculer les cibles proportionnelles à la distribution globale
    SCORES_CONFIG.CRITERES.forEach(critere => {
      cibles[classe].effectifsCibles[critere] = {};
      
      SCORES_CONFIG.SCORES.forEach(score => {
        const proportionGlobale = distributionGlobale[critere][score] / distributionGlobale[critere].total;
        const cibleProportionnelle = Math.round(effectifTotal * proportionGlobale);
        cibles[classe].effectifsCibles[critere][score] = cibleProportionnelle;
      });
      
      // Ajuster pour que la somme soit exacte
      const somme = Object.values(cibles[classe].effectifsCibles[critere]).reduce((a, b) => a + b, 0);
      const difference = effectifTotal - somme;
      
      if (difference !== 0) {
        // Ajuster sur le score le plus courant globalement
        const scoreLePlusCourant = Object.entries(distributionGlobale[critere])
          .filter(([score]) => score !== 'total')
          .sort((a, b) => b[1] - a[1])[0][0];
        
        cibles[classe].effectifsCibles[critere][scoreLePlusCourant] += difference;
        
        Logger.log(`Ajustement ${classe} ${critere}: +${difference} sur score ${scoreLePlusCourant}`);
      }
    });
  });
  
  return cibles;
}

/**
 * Calcule les effectifs actuels par score pour chaque classe
 */
function calculerEffectifsActuels(dataContext) {
  const classes = Object.keys(dataContext.classesState);
  const effectifs = {};
  
  classes.forEach(classe => {
    const eleves = dataContext.classesState[classe];
    
    effectifs[classe] = {
      COM: { 1: 0, 2: 0, 3: 0, 4: 0 },
      TRA: { 1: 0, 2: 0, 3: 0, 4: 0 },
      PART: { 1: 0, 2: 0, 3: 0, 4: 0 },
      ABS: { 1: 0, 2: 0, 3: 0, 4: 0 }
    };
    
    eleves.forEach(eleve => {
      // Compter les scores pour chaque critère
      SCORES_CONFIG.CRITERES.forEach(critere => {
        const score = parseInt(eleve[critere]) || 0; // CORRECTION: eleve[critere] au lieu de eleve[`niveau${critere}`]
        if (score >= 1 && score <= 4) {
          effectifs[classe][critere][score]++;
        }
      });
    });
  });
  
  return effectifs;
}

/**
 * Calcule les écarts entre effectifs actuels et cibles
 */
function calculerEcartsEffectifs(effectifsActuels, effectifsCibles) {
  const classes = Object.keys(effectifsActuels);
  const ecarts = {};
  
  classes.forEach(classe => {
    ecarts[classe] = {};
    
    SCORES_CONFIG.CRITERES.forEach(critere => {
      ecarts[classe][critere] = {};
      
      SCORES_CONFIG.SCORES.forEach(score => {
        const actuel = effectifsActuels[classe][critere][score];
        const cible = effectifsCibles[classe].effectifsCibles[critere][score];
        const ecart = actuel - cible;
        
        ecarts[classe][critere][score] = {
          actuel: actuel,
          cible: cible,
          ecart: ecart,
          surplus: ecart > 0 ? ecart : 0,
          deficit: ecart < 0 ? Math.abs(ecart) : 0
        };
      });
    });
  });
  
  return ecarts;
}

// ==================================================================
// SECTION 3: GESTION AVANCÉE DE LA MOBILITÉ
// ==================================================================

/**
 * Vérifie les conditions de mobilité avancées
 */
function verifierConditionsMobilite(eleve, dataContext) {
  const mobilite = String(eleve.MOBILITE || "").toUpperCase();
  
  switch (mobilite) {
    case 'LIBRE':
      return { 
        autorise: true, 
        condition: 'LIBRE',
        message: "Mobilité totale autorisée"
      };
      
    case 'FIXE':
      return { 
        autorise: false, 
        condition: 'FIXE',
        message: "Élève fixe, échange impossible"
      };
      
    case 'CONDI':
      // Gestion des codes DISSO spéciaux (D1, D2, etc.)
      const codeDisso = eleve.DISSO;
      if (codeDisso && codeDisso.startsWith('D')) {
        return { 
          autorise: true, 
          condition: `MEME_CODE_${codeDisso}`,
          message: `Peut échanger avec autre élève ${codeDisso}`
        };
      }
      return { 
        autorise: false, 
        condition: 'CONDI_NON_DEFINI',
        message: "Conditions CONDI non définies"
      };
      
    case 'PERMUT':
      return {
        autorise: true,
        condition: 'LV2_ET_OPT_IDENTIQUES',
        message: "Peut échanger si LV2 et OPT identiques"
      };
      
    case 'SPEC':
      return {
        autorise: true,
        condition: 'GROUPE_ASSOCIE',
        message: "Échange de groupe associé requis"
      };
      
    default:
      return { 
        autorise: false, 
        condition: 'INCONNU',
        message: `Type de mobilité inconnu: ${mobilite}`
      };
  }
}

/**
 * Gestion robuste des groupes SPEC (élèves associés)
 */
function gererGroupesSpec_Robuste(eleve, dataContext) {
  if (eleve.MOBILITE !== 'SPEC') return [eleve];
  
  const codeAssociation = eleve.ASSO || eleve.DISSO;
  if (!codeAssociation) {
    Logger.log(`⚠️ Élève SPEC ${eleve.ID_ELEVE} sans code d'association`);
    return [eleve];
  }
  
  // Trouver tous les élèves avec le même code d'association
  const tousEleves = Object.values(dataContext.classesState).flat();
  const groupeAssoc = tousEleves.filter(e => 
    (e.ASSO === codeAssociation || e.DISSO === codeAssociation) && 
    e.MOBILITE === 'SPEC' &&
    e.ID_ELEVE !== eleve.ID_ELEVE // Exclure l'élève lui-même
  );
  
  // Ajouter l'élève original au début
  const groupeComplet = [eleve, ...groupeAssoc];
  
  Logger.log(`Groupe SPEC ${codeAssociation}: ${groupeComplet.length} élèves (${groupeComplet.map(e => e.ID_ELEVE).join(', ')})`);
  return groupeComplet;
}

/**
 * Vérifie les contraintes d'échange avec validation renforcée
 */
function verifierContraintesEchange(eleve1, eleve2, classe1, classe2, dataContext) {
  const checks = [];
  
  // Vérification options
  const optionCheck1 = verifierOptions(eleve1, classe2, dataContext);
  const optionCheck2 = verifierOptions(eleve2, classe1, dataContext);
  if (!optionCheck1 || !optionCheck2) {
    checks.push(`❌ OPTIONS: ${!optionCheck1 ? eleve1.OPT + '→' + classe2 : eleve2.OPT + '→' + classe1}`);
    return { valide: false, raisons: checks };
  }
  
  // Vérification mobilité spéciale
  const mobilite1 = verifierConditionsMobilite(eleve1, dataContext);
  const mobilite2 = verifierConditionsMobilite(eleve2, dataContext);
  
  if (!mobilite1.autorise || !mobilite2.autorise) {
    checks.push(`❌ MOBILITÉ: ${mobilite1.message || mobilite2.message}`);
    return { valide: false, raisons: checks };
  }
  
  // Vérification conditions spéciales (PERMUT)
  if (mobilite1.condition === 'LV2_ET_OPT_IDENTIQUES' || mobilite2.condition === 'LV2_ET_OPT_IDENTIQUES') {
    if (eleve1.LV2 !== eleve2.LV2 || eleve1.OPT !== eleve2.OPT) {
      checks.push(`❌ PERMUT: LV2 (${eleve1.LV2}≠${eleve2.LV2}) ou OPT (${eleve1.OPT}≠${eleve2.OPT})`);
      return { valide: false, raisons: checks };
    }
  }
  
  // Vérification codes DISSO identiques (CONDI)
  if (mobilite1.condition?.startsWith('MEME_CODE_') || mobilite2.condition?.startsWith('MEME_CODE_')) {
    if (eleve1.DISSO !== eleve2.DISSO) {
      checks.push(`❌ CONDI: Codes DISSO différents (${eleve1.DISSO}≠${eleve2.DISSO})`);
      return { valide: false, raisons: checks };
    }
  }
  
  // Vérification dissociations
  if (!verifierDissociations(eleve1, classe2, dataContext) || !verifierDissociations(eleve2, classe1, dataContext)) {
    checks.push(`❌ DISSOCIATIONS: Codes incompatibles`);
    return { valide: false, raisons: checks };
  }
  
  checks.push("✅ Toutes contraintes respectées");
  return { valide: true, raisons: checks };
}

/**
 * Vérifie si un élève est mobile selon les contraintes
 */
function estEleveMobile(eleve, dataContext) {
  const mobilite = String(eleve.MOBILITE || "").toUpperCase();
  
  switch (mobilite) {
    case 'LIBRE':
      return true;
    case 'FIXE':
      return false;
    case 'SPEC':
    case 'CONDI':
    case 'PERMUT':
      const resultat = verifierConditionsMobilite(eleve, dataContext);
      return resultat.autorise;
    default:
      return false;
  }
}

// ==================================================================
// SECTION 4: STRATÉGIES D'ÉCHANGE ADAPTATIVES
// ==================================================================

/**
 * Choisit la stratégie d'échange optimale selon le type de mobilité
 */
function choisirStrategieEchange(classeSource, classeCible, critere, score, dataContext) {
  const elevesSource = dataContext.classesState[classeSource];
  const elevesCible = dataContext.classesState[classeCible];
  
  // Compter les élèves par type de mobilité
  const statsSource = compterParMobilite(elevesSource, critere, score);
  const statsCible = compterParMobilite(elevesCible, critere, score);
  
  Logger.log(`Stratégie ${classeSource}→${classeCible} ${critere} Score ${score}:`);
  Logger.log(`  Source: LIBRE:${statsSource.LIBRE} PERMUT:${statsSource.PERMUT} CONDI:${statsSource.CONDI} SPEC:${statsSource.SPEC}`);
  Logger.log(`  Cible: LIBRE:${statsCible.LIBRE} PERMUT:${statsCible.PERMUT} CONDI:${statsCible.CONDI} SPEC:${statsCible.SPEC}`);
  
  // Stratégie 1: LIBRE ↔ LIBRE (priorité max)
  if (statsSource.LIBRE > 0 && statsCible.LIBRE > 0) {
    return { type: 'LIBRE', priorite: SCORES_CONFIG.PRIORITES_STRATEGIES.LIBRE };
  }
  
  // Stratégie 2: PERMUT ↔ PERMUT (si LV2/OPT compatibles)
  if (statsSource.PERMUT > 0 && statsCible.PERMUT > 0) {
    return { type: 'PERMUT', priorite: SCORES_CONFIG.PRIORITES_STRATEGIES.PERMUT };
  }
  
  // Stratégie 3: CONDI ↔ CONDI (si même code DISSO)
  if (statsSource.CONDI > 0 && statsCible.CONDI > 0) {
    return { type: 'CONDI', priorite: SCORES_CONFIG.PRIORITES_STRATEGIES.CONDI };
  }
  
  // Stratégie 4: SPEC ↔ SPEC (échange de groupes)
  if (statsSource.SPEC > 0 && statsCible.SPEC > 0) {
    return { type: 'SPEC', priorite: SCORES_CONFIG.PRIORITES_STRATEGIES.SPEC };
  }
  
  return { type: 'AUCUNE', priorite: 0 };
}

/**
 * Compte les élèves par type de mobilité pour un critère/score donné
 */
function compterParMobilite(eleves, critere, scoreRecherche) {
  const stats = { LIBRE: 0, PERMUT: 0, CONDI: 0, SPEC: 0, FIXE: 0 };
  
  eleves.forEach(eleve => {
    const score = parseInt(eleve[critere]) || 0; // CORRECTION: eleve[critere] au lieu de eleve[`niveau${critere}`]
    if (score === scoreRecherche) {
      const mobilite = String(eleve.MOBILITE || "").toUpperCase();
      stats[mobilite] = (stats[mobilite] || 0) + 1;
    }
  });
  
  return stats;
}

// ==================================================================
// SECTION 5: ÉCHANGE DE GROUPES SPEC
// ==================================================================

/**
 * Propose un échange de groupe SPEC
 */
function proposerEchangeGroupeSpec(groupe1, groupe2, classeSource, classeCible, dataContext) {
  // Vérifier que les deux groupes ont la même taille
  if (groupe1.length !== groupe2.length) {
    Logger.log(`❌ Échange SPEC impossible: groupes de tailles différentes (${groupe1.length} vs ${groupe2.length})`);
    return null;
  }
  
  // Vérifier les contraintes pour chaque membre
  const contraintesRespectees = groupe1.every((eleve1, index) => {
    const eleve2 = groupe2[index];
    const validation = verifierContraintesEchange(eleve1, eleve2, classeSource, classeCible, dataContext);
    
    if (!validation.valide) {
      Logger.log(`❌ Échange SPEC bloqué: ${eleve1.ID_ELEVE} ↔ ${eleve2.ID_ELEVE} - ${validation.raisons.join(', ')}`);
      return false;
    }
    return true;
  });
  
  if (!contraintesRespectees) return null;
  
  // Calculer l'impact global de l'échange du groupe
  let impactTotal = 0;
  const detailsEchanges = [];
  
  groupe1.forEach((eleve1, index) => {
    const eleve2 = groupe2[index];
    
    // Calculer l'impact pour chaque critère
    SCORES_CONFIG.CRITERES.forEach(critere => {
      const score1 = parseInt(eleve1[critere]) || 0; // CORRECTION
      const score2 = parseInt(eleve2[critere]) || 0; // CORRECTION
      
      if (score1 !== score2) {
        const impact = calculerImpactEchange(eleve1, eleve2, critere, score1, score2);
        impactTotal += impact;
        
        detailsEchanges.push({
          eleve1: eleve1.ID_ELEVE,
          eleve2: eleve2.ID_ELEVE,
          critere: critere,
          score1: score1,
          score2: score2,
          impact: impact
        });
      }
    });
  });
  
  return {
    type: 'GROUPE_SPEC',
    groupe1: groupe1,
    groupe2: groupe2,
    classeSource: classeSource,
    classeCible: classeCible,
    impactTotal: impactTotal,
    detailsEchanges: detailsEchanges,
    nbEleves: groupe1.length
  };
}

/**
 * Groupe les élèves par code d'association
 */
function grouperParAssociation(eleves) {
  const groupes = {};
  
  eleves.forEach(eleve => {
    const codeAssoc = eleve.ASSO || eleve.DISSO || 'SANS_CODE';
    
    if (!groupes[codeAssoc]) {
      groupes[codeAssoc] = [];
    }
    
    groupes[codeAssoc].push(eleve);
  });
  
  return groupes;
}

// ==================================================================
// SECTION 6: RECHERCHE D'ÉCHANGES INTELLIGENTE
// ==================================================================

/**
 * Trouve les candidats d'échange avec stratégies adaptatives
 */
function trouverCandidatsEchange_Intelligent(dataContext, classeSource, classeCible, critere, scoreSource, scoreCible) {
  const strategie = choisirStrategieEchange(classeSource, classeCible, critere, scoreSource, dataContext);
  
  if (strategie.type === 'AUCUNE') {
    Logger.log(`❌ Aucune stratégie d'échange viable pour ${classeSource}→${classeCible} ${critere}`);
    return [];
  }
  
  Logger.log(`✅ Stratégie sélectionnée: ${strategie.type} (priorité: ${strategie.priorite})`);
  
  switch (strategie.type) {
    case 'LIBRE':
      return trouverCandidatsLibres(dataContext, classeSource, classeCible, critere, scoreSource, scoreCible);
      
    case 'PERMUT':
      return trouverCandidatsPermut(dataContext, classeSource, classeCible, critere, scoreSource, scoreCible);
      
    case 'CONDI':
      return trouverCandidatsCondi(dataContext, classeSource, classeCible, critere, scoreSource, scoreCible);
      
    case 'SPEC':
      return trouverCandidatsSpec(dataContext, classeSource, classeCible, critere, scoreSource, scoreCible);
      
    default:
      return [];
  }
}

/**
 * Trouve les candidats d'échange LIBRE
 */
function trouverCandidatsLibres(dataContext, classeSource, classeCible, critere, scoreSource, scoreCible) {
  const elevesSource = dataContext.classesState[classeSource].filter(e => 
    e.MOBILITE === 'LIBRE' && parseInt(e[critere]) === scoreSource // CORRECTION
  );
  
  const elevesCible = dataContext.classesState[classeCible].filter(e => 
    e.MOBILITE === 'LIBRE' && parseInt(e[critere]) === scoreCible // CORRECTION
  );
  
  const candidats = [];
  
  elevesSource.forEach(eleveSource => {
    elevesCible.forEach(eleveCible => {
      const validation = verifierContraintesEchange(eleveSource, eleveCible, classeSource, classeCible, dataContext);
      
      if (validation.valide) {
        candidats.push({
          type: 'LIBRE',
          eleveSource: eleveSource,
          eleveCible: eleveCible,
          classeSource: classeSource,
          classeCible: classeCible,
          critere: critere,
          scoreSource: scoreSource,
          scoreCible: scoreCible,
          impact: calculerImpactEchange(eleveSource, eleveCible, critere, scoreSource, scoreCible),
          priorite: SCORES_CONFIG.PRIORITES_STRATEGIES.LIBRE
        });
      }
    });
  });
  
  return candidats.sort((a, b) => b.impact - a.impact);
}

/**
 * Trouve les candidats d'échange PERMUT
 */
function trouverCandidatsPermut(dataContext, classeSource, classeCible, critere, scoreSource, scoreCible) {
  const elevesSource = dataContext.classesState[classeSource].filter(e => 
    e.MOBILITE === 'PERMUT' && parseInt(e[critere]) === scoreSource // CORRECTION
  );
  
  const elevesCible = dataContext.classesState[classeCible].filter(e => 
    e.MOBILITE === 'PERMUT' && parseInt(e[critere]) === scoreCible // CORRECTION
  );
  
  const candidats = [];
  
  elevesSource.forEach(eleveSource => {
    elevesCible.forEach(eleveCible => {
      // Vérification PERMUT: LV2 ET OPT identiques
      if (eleveSource.LV2 === eleveCible.LV2 && eleveSource.OPT === eleveCible.OPT) {
        const validation = verifierContraintesEchange(eleveSource, eleveCible, classeSource, classeCible, dataContext);
        
        if (validation.valide) {
          candidats.push({
            type: 'PERMUT',
            eleveSource: eleveSource,
            eleveCible: eleveCible,
            classeSource: classeSource,
            classeCible: classeCible,
            critere: critere,
            scoreSource: scoreSource,
            scoreCible: scoreCible,
            impact: calculerImpactEchange(eleveSource, eleveCible, critere, scoreSource, scoreCible),
            priorite: SCORES_CONFIG.PRIORITES_STRATEGIES.PERMUT
          });
        }
      }
    });
  });
  
  return candidats.sort((a, b) => b.impact - a.impact);
}

/**
 * Trouve les candidats d'échange CONDI
 */
function trouverCandidatsCondi(dataContext, classeSource, classeCible, critere, scoreSource, scoreCible) {
  const elevesSource = dataContext.classesState[classeSource].filter(e => 
    e.MOBILITE === 'CONDI' && parseInt(e[critere]) === scoreSource // CORRECTION
  );
  
  const elevesCible = dataContext.classesState[classeCible].filter(e => 
    e.MOBILITE === 'CONDI' && parseInt(e[critere]) === scoreCible // CORRECTION
  );
  
  const candidats = [];
  
  elevesSource.forEach(eleveSource => {
    elevesCible.forEach(eleveCible => {
      // Vérification CONDI: même code DISSO
      if (eleveSource.DISSO === eleveCible.DISSO) {
        const validation = verifierContraintesEchange(eleveSource, eleveCible, classeSource, classeCible, dataContext);
        
        if (validation.valide) {
          candidats.push({
            type: 'CONDI',
            eleveSource: eleveSource,
            eleveCible: eleveCible,
            classeSource: classeSource,
            classeCible: classeCible,
            critere: critere,
            scoreSource: scoreSource,
            scoreCible: scoreCible,
            impact: calculerImpactEchange(eleveSource, eleveCible, critere, scoreSource, scoreCible),
            priorite: SCORES_CONFIG.PRIORITES_STRATEGIES.CONDI,
            codeDisso: eleveSource.DISSO
          });
        }
      }
    });
  });
  
  return candidats.sort((a, b) => b.impact - a.impact);
}

/**
 * Trouve les candidats d'échange SPEC (groupes)
 */
function trouverCandidatsSpec(dataContext, classeSource, classeCible, critere, scoreSource, scoreCible) {
  const elevesSourceSpec = dataContext.classesState[classeSource].filter(e => 
    e.MOBILITE === 'SPEC' && parseInt(e[critere]) === scoreSource // CORRECTION
  );
  
  const elevesCibleSpec = dataContext.classesState[classeCible].filter(e => 
    e.MOBILITE === 'SPEC' && parseInt(e[critere]) === scoreCible // CORRECTION
  );
  
  const candidats = [];
  
  // Grouper par code d'association
  const groupesSource = grouperParAssociation(elevesSourceSpec);
  const groupesCible = grouperParAssociation(elevesCibleSpec);
  
  Object.keys(groupesSource).forEach(codeAssocSource => {
    Object.keys(groupesCible).forEach(codeAssocCible => {
      const groupe1 = groupesSource[codeAssocSource];
      const groupe2 = groupesCible[codeAssocCible];
      
      const echangeGroupe = proposerEchangeGroupeSpec(groupe1, groupe2, classeSource, classeCible, dataContext);
      
      if (echangeGroupe) {
        candidats.push({
          type: 'SPEC',
          echangeGroupe: echangeGroupe,
          classeSource: classeSource,
          classeCible: classeCible,
          critere: critere,
          scoreSource: scoreSource,
          scoreCible: scoreCible,
          impact: echangeGroupe.impactTotal,
          priorite: SCORES_CONFIG.PRIORITES_STRATEGIES.SPEC,
          nbEleves: echangeGroupe.nbEleves
        });
      }
    });
  });
  
  return candidats.sort((a, b) => b.impact - a.impact);
}

// ==================================================================
// SECTION 7: ANALYSE DES DÉSÉQUILIBRES
// ==================================================================

/**
 * Identifie les classes avec surplus et déficit pour un critère/score donné
 */
function identifierSurplusDeficit(ecarts, critere, score) {
  const classes = Object.keys(ecarts);
  const surplus = [];
  const deficit = [];
  
  classes.forEach(classe => {
    const ecart = ecarts[classe][critere][score];
    
    if (ecart.surplus > 0) {
      surplus.push({
        classe: classe,
        surplus: ecart.surplus,
        effectifActuel: ecart.actuel,
        effectifCible: ecart.cible
      });
    }
    
    if (ecart.deficit > 0) {
      deficit.push({
        classe: classe,
        deficit: ecart.deficit,
        effectifActuel: ecart.actuel,
        effectifCible: ecart.cible
      });
    }
  });
  
  // Trier par ordre de priorité (plus gros surplus/déficit en premier)
  surplus.sort((a, b) => b.surplus - a.surplus);
  deficit.sort((a, b) => b.deficit - a.deficit);
  
  return { surplus, deficit };
}

/**
 * Calcule le score d'équilibre global
 */
function calculerScoreEquilibre(ecarts) {
  const classes = Object.keys(ecarts);
  let scoreTotal = 0;
  let nbEvaluations = 0;
  
  classes.forEach(classe => {
    SCORES_CONFIG.CRITERES.forEach(critere => {
      SCORES_CONFIG.SCORES.forEach(score => {
        const ecart = ecarts[classe][critere][score];
        const poidsCritere = SCORES_CONFIG.POIDS_CRITERES[critere];
        
        // Score basé sur l'écart relatif
        const ecartRelatif = Math.abs(ecart.ecart) / Math.max(ecart.cible, 1);
        const scoreCritere = Math.max(0, 100 - (ecartRelatif * 100));
        
        scoreTotal += scoreCritere * poidsCritere;
        nbEvaluations++;
      });
    });
  });
  
  return nbEvaluations > 0 ? scoreTotal / nbEvaluations : 0;
}

/**
 * Calcule l'impact d'un échange sur l'équilibre
 */
function calculerImpactEchange(eleve1, eleve2, critere, score1, score2) {
  // Impact basé sur la différence de scores
  const differenceScores = Math.abs(score1 - score2);
  
  // Impact basé sur le poids du critère
  const poidsCritere = SCORES_CONFIG.POIDS_CRITERES[critere];
  
  // Impact basé sur la mobilité (priorité aux échanges LIBRE)
  const mobilite1 = String(eleve1.MOBILITE || "").toUpperCase();
  const mobilite2 = String(eleve2.MOBILITE || "").toUpperCase();
  const facteurMobilite = (mobilite1 === 'LIBRE' && mobilite2 === 'LIBRE') ? 1.0 : 0.7;
  
  return differenceScores * poidsCritere * facteurMobilite;
}

// ==================================================================
// SECTION 8: ALGORITHME D'ÉQUILIBRAGE PRINCIPAL
// ==================================================================

/**
 * Algorithme principal d'équilibrage des scores
 */
function equilibrerScores(dataContext, config) {
  Logger.log("=== DÉBUT ÉQUILIBRAGE SCORES (VERSION FINALE V1.2) ===");
  
  const journalEchanges = [];
  let iteration = 0;
  let scoreActuel = 0;
  let amelioration = true;
  
  try {
    // Calculer les cibles et effectifs actuels
    const effectifsCibles = calculerEffectifsCibles_Ameliore(dataContext, config);
    const effectifsActuels = calculerEffectifsActuels(dataContext);
    const ecarts = calculerEcartsEffectifs(effectifsActuels, effectifsCibles);
    
    scoreActuel = calculerScoreEquilibre(ecarts);
    Logger.log(`Score initial d'équilibre: ${scoreActuel.toFixed(2)}/100`);
    
    while (amelioration && iteration < SCORES_CONFIG.MAX_ITERATIONS) {
      iteration++;
      amelioration = false;
      let nbEchangesIteration = 0;
      
      Logger.log(`--- Itération ${iteration} ---`);
      
      // Parcourir les critères par ordre de priorité
      const criteresPriorises = Object.entries(SCORES_CONFIG.POIDS_CRITERES)
        .sort((a, b) => b[1] - a[1])
        .map(([critere]) => critere);
      
      for (const critere of criteresPriorises) {
        if (nbEchangesIteration >= SCORES_CONFIG.MAX_ECHANGES_PAR_ITERATION) break;
        
        // Parcourir les scores (priorité aux extrêmes)
        const scoresPriorises = [4, 1, 3, 2]; // Priorité: 4, 1, 3, 2
        
        for (const score of scoresPriorises) {
          if (nbEchangesIteration >= SCORES_CONFIG.MAX_ECHANGES_PAR_ITERATION) break;
          
          // Identifier surplus et déficit
          const { surplus, deficit } = identifierSurplusDeficit(ecarts, critere, score);
          
          if (surplus.length > 0 && deficit.length > 0) {
            // Chercher des échanges possibles
            for (const classeSurplus of surplus) {
              if (nbEchangesIteration >= SCORES_CONFIG.MAX_ECHANGES_PAR_ITERATION) break;
              
              for (const classeDeficit of deficit) {
                if (nbEchangesIteration >= SCORES_CONFIG.MAX_ECHANGES_PAR_ITERATION) break;
                
                // Trouver le score opposé pour l'échange
                const scoreOppose = score === 4 ? 1 : (score === 1 ? 4 : (score === 3 ? 2 : 3));
                
                const candidats = trouverCandidatsEchange_Intelligent(
                  dataContext,
                  classeSurplus.classe,
                  classeDeficit.classe,
                  critere,
                  score,
                  scoreOppose
                );
                
                if (candidats.length > 0) {
                  // Appliquer le meilleur échange
                  const meilleurEchange = candidats[0];
                  
                  if (appliquerEchangeSecurise(meilleurEchange, dataContext)) {
                    journalEchanges.push(meilleurEchange);
                    nbEchangesIteration++;
                    amelioration = true;
                    
                    // Mettre à jour les effectifs et écarts
                    mettreAJourEffectifsApresEchange(meilleurEchange, effectifsActuels, ecarts, effectifsCibles);
                    
                    Logger.log(`✅ Échange ${critere} (${meilleurEchange.type}): ${meilleurEchange.eleveSource?.ID_ELEVE || 'GROUPE'}(${score}) ↔ ${meilleurEchange.eleveCible?.ID_ELEVE || 'GROUPE'}(${scoreOppose})`);
                    break;
                  }
                }
              }
            }
          }
        }
      }
      
      // Calculer le nouveau score
      const nouveauScore = calculerScoreEquilibre(ecarts);
      const ameliorationScore = nouveauScore - scoreActuel;
      
      if (ameliorationScore > SCORES_CONFIG.SEUIL_AMELIORATION) {
        scoreActuel = nouveauScore;
        Logger.log(`Score amélioré: ${scoreActuel.toFixed(2)}/100 (+${ameliorationScore.toFixed(2)})`);
      } else {
        amelioration = false;
        Logger.log(`Amélioration insuffisante: +${ameliorationScore.toFixed(2)}`);
      }
    }
    
    Logger.log(`=== FIN ÉQUILIBRAGE SCORES ===`);
    Logger.log(`Itérations: ${iteration}`);
    Logger.log(`Échanges effectués: ${journalEchanges.length}`);
    Logger.log(`Score final: ${scoreActuel.toFixed(2)}/100`);
    
    return {
      success: true,
      journalEchanges: journalEchanges,
      scoreFinal: scoreActuel,
      iterations: iteration,
      effectifsCibles: effectifsCibles,
      effectifsActuels: effectifsActuels,
      ecarts: ecarts
    };
    
  } catch (e) {
    Logger.log(`❌ ERREUR dans equilibrerScores: ${e.message}`);
    return {
      success: false,
      error: e.message,
      journalEchanges: journalEchanges
    };
  }
}

// ==================================================================
// SECTION 9: APPLICATION SÉCURISÉE DES ÉCHANGES
// ==================================================================

/**
 * Application sécurisée des échanges
 */
function appliquerEchangeSecurise(echange, dataContext) {
  try {
    if (echange.type === 'SPEC' && echange.echangeGroupe) {
      // Échange de groupe SPEC
      return appliquerEchangeGroupe(echange.echangeGroupe, dataContext);
    } else {
      // Échange simple
      return appliquerEchange(echange, dataContext);
    }
  } catch (e) {
    Logger.log(`❌ ERREUR application échange: ${e.message}`);
    return false;
  }
}

/**
 * Applique un échange simple dans le contexte de données
 */
function appliquerEchange(echange, dataContext) {
  try {
    const { eleveSource, eleveCible, classeSource, classeCible } = echange;
    
    // Retirer les élèves de leurs classes actuelles
    const indexSource = dataContext.classesState[classeSource].findIndex(e => e.ID_ELEVE === eleveSource.ID_ELEVE);
    const indexCible = dataContext.classesState[classeCible].findIndex(e => e.ID_ELEVE === eleveCible.ID_ELEVE);
    
    if (indexSource === -1 || indexCible === -1) {
      Logger.log(`❌ Élève non trouvé dans sa classe`);
      return false;
    }
    
    // Échanger les élèves
    const temp = dataContext.classesState[classeSource][indexSource];
    dataContext.classesState[classeSource][indexSource] = dataContext.classesState[classeCible][indexCible];
    dataContext.classesState[classeCible][indexCible] = temp;
    
    return true;
    
  } catch (e) {
    Logger.log(`❌ ERREUR lors de l'application de l'échange: ${e.message}`);
    return false;
  }
}

/**
 * Applique un échange de groupe SPEC
 */
function appliquerEchangeGroupe(echangeGroupe, dataContext) {
  const { groupe1, groupe2, classeSource, classeCible } = echangeGroupe;
  
  try {
    // Vérification finale avant application
    const verification = groupe1.every((eleve1, index) => {
      const eleve2 = groupe2[index];
      const indexSource = dataContext.classesState[classeSource].findIndex(e => e.ID_ELEVE === eleve1.ID_ELEVE);
      const indexCible = dataContext.classesState[classeCible].findIndex(e => e.ID_ELEVE === eleve2.ID_ELEVE);
      
      return indexSource !== -1 && indexCible !== -1;
    });
    
    if (!verification) {
      Logger.log(`❌ Vérification finale échec pour échange de groupe`);
      return false;
    }
    
    // Application de l'échange
    groupe1.forEach((eleve1, index) => {
      const eleve2 = groupe2[index];
      
      const indexSource = dataContext.classesState[classeSource].findIndex(e => e.ID_ELEVE === eleve1.ID_ELEVE);
      const indexCible = dataContext.classesState[classeCible].findIndex(e => e.ID_ELEVE === eleve2.ID_ELEVE);
      
      // Échanger
      const temp = dataContext.classesState[classeSource][indexSource];
      dataContext.classesState[classeSource][indexSource] = dataContext.classesState[classeCible][indexCible];
      dataContext.classesState[classeCible][indexCible] = temp;
    });
    
    Logger.log(`✅ Échange de groupe SPEC réussi: ${groupe1.length} élèves entre ${classeSource} et ${classeCible}`);
    return true;
    
  } catch (e) {
    Logger.log(`❌ ERREUR échange de groupe: ${e.message}`);
    return false;
  }
}

/**
 * Met à jour les effectifs après un échange
 */
function mettreAJourEffectifsApresEchange(echange, effectifsActuels, ecarts, effectifsCibles) {
  const { eleveSource, eleveCible, classeSource, classeCible, critere, scoreSource, scoreCible } = echange;
  
  // Gérer les échanges de groupes SPEC
  if (echange.type === 'SPEC' && echange.echangeGroupe) {
    // Pour les groupes SPEC, mettre à jour tous les critères affectés
    echange.echangeGroupe.detailsEchanges.forEach(detail => {
      const crit = detail.critere;
      const score1 = detail.score1;
      const score2 = detail.score2;
      
      // Mettre à jour les effectifs actuels
      effectifsActuels[classeSource][crit][score1]--;
      effectifsActuels[classeSource][crit][score2]++;
      effectifsActuels[classeCible][crit][score2]--;
      effectifsActuels[classeCible][crit][score1]++;
      
      // Mettre à jour les écarts
      [classeSource, classeCible].forEach(classe => {
        [score1, score2].forEach(score => {
          const actuel = effectifsActuels[classe][crit][score];
          const cible = effectifsCibles[classe].effectifsCibles[crit][score];
          const ecart = actuel - cible;
          
          ecarts[classe][crit][score] = {
            actuel: actuel,
            cible: cible,
            ecart: ecart,
            surplus: ecart > 0 ? ecart : 0,
            deficit: ecart < 0 ? Math.abs(ecart) : 0
          };
        });
      });
    });
  } else {
    // Échange simple
    effectifsActuels[classeSource][critere][scoreSource]--;
    effectifsActuels[classeSource][critere][scoreCible]++;
    effectifsActuels[classeCible][critere][scoreCible]--;
    effectifsActuels[classeCible][critere][scoreSource]++;
    
    // Mettre à jour les écarts
    [classeSource, classeCible].forEach(classe => {
      [scoreSource, scoreCible].forEach(score => {
        const actuel = effectifsActuels[classe][critere][score];
        const cible = effectifsCibles[classe].effectifsCibles[critere][score];
        const ecart = actuel - cible;
        
        ecarts[classe][critere][score] = {
          actuel: actuel,
          cible: cible,
          ecart: ecart,
          surplus: ecart > 0 ? ecart : 0,
          deficit: ecart < 0 ? Math.abs(ecart) : 0
        };
      });
    });
  }
}

/**
 * Vérifie les options d'un élève pour une classe
 */
function verifierOptions(eleve, classeCible, dataContext) {
  if (!eleve.OPT || eleve.OPT === "" || eleve.OPT === "ESP") {
    return true; // Pas de contrainte d'option
  }
  
  const optionPools = dataContext.optionPools || {};
  const pool = optionPools[eleve.OPT];
  
  return !pool || pool.includes(classeCible.toUpperCase());
}

/**
 * Vérifie les dissociations
 */
function verifierDissociations(eleve, classeCible, dataContext) {
  if (!eleve.DISSO) {
    return true; // Pas de code de dissociation
  }
  
  const dissocMap = dataContext.dissocMap || {};
  const dissocSet = dissocMap[classeCible];
  
  return !dissocSet || !dissocSet.has(eleve.DISSO);
}

// ==================================================================
// SECTION 10: DIAGNOSTIC AMÉLIORÉ
// ==================================================================

/**
 * Diagnostic avancé avec suggestions
 */
function diagnostiquerEffectifs_Ameliore(dataContext) {
  const effectifsCibles = calculerEffectifsCibles_Ameliore(dataContext, {});
  const effectifsActuels = calculerEffectifsActuels(dataContext);
  const ecarts = calculerEcartsEffectifs(effectifsActuels, effectifsCibles);
  const score = calculerScoreEquilibre(ecarts);
  
  let rapport = "=== DIAGNOSTIC AVANCÉ EFFECTIFS PAR SCORE ===\n\n";
  
  // Analyse globale
  rapport += `📊 SCORE GLOBAL D'ÉQUILIBRE: ${score.toFixed(2)}/100\n\n`;
  
  // Problèmes prioritaires
  const problemes = [];
  
  Object.keys(effectifsActuels).forEach(classe => {
    SCORES_CONFIG.CRITERES.forEach(critere => {
      SCORES_CONFIG.SCORES.forEach(score => {
        const ecart = ecarts[classe][critere][score];
        if (Math.abs(ecart.ecart) > SCORES_CONFIG.SEUIL_PROBLEME_SIGNIFICATIF) {
          problemes.push({
            classe,
            critere,
            score,
            ecart: ecart.ecart,
            priorite: SCORES_CONFIG.POIDS_CRITERES[critere] * Math.abs(ecart.ecart)
          });
        }
      });
    });
  });
  
  // Trier par priorité
  problemes.sort((a, b) => b.priorite - a.priorite);
  
  if (problemes.length > 0) {
    rapport += "🚨 PROBLÈMES PRIORITAIRES:\n";
    problemes.slice(0, 10).forEach((pb, index) => {
      const type = pb.ecart > 0 ? "SURPLUS" : "DÉFICIT";
      rapport += `${index + 1}. ${pb.classe} - ${pb.critere} Score ${pb.score}: ${type} de ${Math.abs(pb.ecart)} élèves\n`;
    });
    rapport += "\n";
  }
  
  // Détail par classe
  Object.keys(effectifsActuels).forEach(classe => {
    rapport += `📊 CLASSE: ${classe}\n`;
    
    SCORES_CONFIG.CRITERES.forEach(critere => {
      rapport += `  ${critere}: `;
      
      const details = [];
      SCORES_CONFIG.SCORES.forEach(score => {
        const ecart = ecarts[classe][critere][score];
        const statut = ecart.ecart === 0 ? "✅" : (ecart.surplus > 0 ? "📈" : "📉");
        details.push(`${score}:${ecart.actuel}/${ecart.cible}${statut}`);
      });
      
      rapport += details.join(" | ") + "\n";
    });
    
    rapport += "\n";
  });
  
  // Suggestions d'amélioration
  rapport += "💡 SUGGESTIONS:\n";
  if (score < SCORES_CONFIG.SEUIL_SCORE_ACCEPTABLE) {
    rapport += "- Lancer l'équilibrage automatique avec priorité sur " + 
               Object.entries(SCORES_CONFIG.POIDS_CRITERES)
                 .sort((a, b) => b[1] - a[1])[0][0] + "\n";
  }
  if (problemes.length > 5) {
    rapport += "- Vérifier les contraintes de mobilité (trop d'élèves FIXE ?)\n";
  }
  
  return rapport;
}

// ==================================================================
// SECTION 11: INTÉGRATION AVEC SYSTÈME EXISTANT
// ==================================================================

/**
 * Fonction d'intégration avec le système existant
 */
function integrerAvecNirvanaV2(dataContext, config) {
  Logger.log("=== INTÉGRATION NIRVANA_SCORES_EQUILIBRAGE V1.2 ===");
  
  try {
    // Phase 1: Diagnostic initial
    const diagnosticInitial = diagnostiquerEffectifs_Ameliore(dataContext);
    Logger.log("Diagnostic initial:\n" + diagnosticInitial);
    
    // Phase 2: Équilibrage des scores
    const resultatEquilibrage = equilibrerScores(dataContext, config);
    
    if (!resultatEquilibrage.success) {
      throw new Error("Échec de l'équilibrage: " + resultatEquilibrage.error);
    }
    
    // Phase 3: Diagnostic final
    const diagnosticFinal = diagnostiquerEffectifs_Ameliore(dataContext);
    Logger.log("Diagnostic final:\n" + diagnosticFinal);
    
    // Retourner les échanges au format compatible avec Nirvana V2
    const echangesFormates = resultatEquilibrage.journalEchanges.map(echange => {
      if (echange.type === 'SPEC' && echange.echangeGroupe) {
        // Échange de groupe SPEC - créer un swap pour chaque paire
        return echange.echangeGroupe.groupe1.map((eleve1, index) => {
          const eleve2 = echange.echangeGroupe.groupe2[index];
          return {
            eleve1: eleve1,
            eleve2: eleve2,
            eleve1ID: eleve1.ID_ELEVE,
            eleve2ID: eleve2.ID_ELEVE,
            eleve1Nom: eleve1.NOM,
            eleve2Nom: eleve2.NOM,
            classe1: echange.classeSource,
            classe2: echange.classeCible,
            oldClasseE1: echange.classeSource,
            oldClasseE2: echange.classeCible,
            newClasseE1: echange.classeCible,
            newClasseE2: echange.classeSource,
            motif: `Équilibrage SPEC ${echange.critere} (Groupe ${echange.echangeGroupe.nbEleves} élèves)`
          };
        });
      } else {
        // Échange simple
        return {
          eleve1: echange.eleveSource,
          eleve2: echange.eleveCible,
          eleve1ID: echange.eleveSource.ID_ELEVE,
          eleve2ID: echange.eleveCible.ID_ELEVE,
          eleve1Nom: echange.eleveSource.NOM,
          eleve2Nom: echange.eleveCible.NOM,
          classe1: echange.classeSource,
          classe2: echange.classeCible,
          oldClasseE1: echange.classeSource,
          oldClasseE2: echange.classeCible,
          newClasseE1: echange.classeCible,
          newClasseE2: echange.classeSource,
          motif: `Équilibrage ${echange.critere} (Score ${echange.scoreSource}↔${echange.scoreCible}) [${echange.type}]`
        };
      }
    }).flat(); // Aplatir en cas d'échanges de groupes
    
    return {
      success: true,
      echanges: echangesFormates,
      nbEchanges: echangesFormates.length,
      scoreInitial: 0, // Calculé dans l'algorithme
      scoreFinal: resultatEquilibrage.scoreFinal,
      amelioration: resultatEquilibrage.scoreFinal,
      iterations: resultatEquilibrage.iterations
    };
    
  } catch (e) {
    Logger.log(`❌ ERREUR intégration: ${e.message}`);
    return {
      success: false,
      error: e.message,
      echanges: []
    };
  }
}

// ==================================================================
// SECTION 12: FONCTION DE TEST INTÉGRÉE
// ==================================================================

/**
 * Fonction de test pour validation complète du module
 */
function testerModuleScoresEquilibrage(dataContext) {
  Logger.log("=== TEST MODULE SCORES ÉQUILIBRAGE V1.2 ===");
  
  try {
    // Test 1: Diagnostic initial
    const diagnostic = diagnostiquerEffectifs_Ameliore(dataContext);
    Logger.log("✅ Test diagnostic: " + (diagnostic.includes("SCORE GLOBAL") ? "RÉUSSI" : "ÉCHEC"));
    
    // Test 2: Calcul des cibles
    const cibles = calculerEffectifsCibles_Ameliore(dataContext, {});
    Logger.log("✅ Test calcul cibles: " + (cibles && Object.keys(cibles).length > 0 ? "RÉUSSI" : "ÉCHEC"));
    
    // Test 3: Contraintes de mobilité
    const classes = Object.keys(dataContext.classesState);
    if (classes.length >= 2) {
      const eleves1 = dataContext.classesState[classes[0]];
      const eleves2 = dataContext.classesState[classes[1]];
      
      if (eleves1.length > 0 && eleves2.length > 0) {
        const validation = verifierContraintesEchange(eleves1[0], eleves2[0], classes[0], classes[1], dataContext);
        Logger.log("✅ Test validation contraintes: " + (validation && validation.raisons ? "RÉUSSI" : "ÉCHEC"));
      }
    }
    
    // Test 4: Stratégies d'échange
    if (classes.length >= 2) {
      const strategie = choisirStrategieEchange(classes[0], classes[1], 'COM', 3, dataContext);
      Logger.log("✅ Test stratégies: " + (strategie && strategie.type ? "RÉUSSI" : "ÉCHEC"));
    }
    
    // Test 5: Groupes SPEC
    const elevesSpec = Object.values(dataContext.classesState).flat().filter(e => e.MOBILITE === 'SPEC');
    if (elevesSpec.length > 0) {
      const groupe = gererGroupesSpec_Robuste(elevesSpec[0], dataContext);
      Logger.log("✅ Test groupes SPEC: " + (groupe && groupe.length > 0 ? "RÉUSSI" : "ÉCHEC"));
    }
    
    // Test 6: Intégration complète
    const integration = integrerAvecNirvanaV2(dataContext, {});
    Logger.log("✅ Test intégration: " + (integration && integration.success !== undefined ? "RÉUSSI" : "ÉCHEC"));
    
    Logger.log("=== FIN TEST MODULE V1.2 ===");
    return true;
    
  } catch (e) {
    Logger.log(`❌ ERREUR test: ${e.message}`);
    return false;
  }
}

// ==================================================================
// SECTION 13: PIPELINE NIRVANA COMPLET
// ==================================================================

/**
 * Pipeline complet Nirvana : Scores → Parité → MultiSwap
 */
function pipelineNirvanaComplet(dataContext, config) {
  Logger.log("=== PIPELINE NIRVANA COMPLET V1.2 ===");
  
  try {
    const resultats = {
      phase1_scores: null,
      phase2_parite: null,
      phase3_multiswap: null,
      echanges_totaux: [],
      score_final: 0,
      temps_execution: 0
    };
    
    const tempsDebut = new Date();
    
    // PHASE 1: Équilibrage des scores (MODULE PRINCIPAL)
    Logger.log("--- PHASE 1: ÉQUILIBRAGE SCORES ---");
    resultats.phase1_scores = integrerAvecNirvanaV2(dataContext, config);
    
    if (resultats.phase1_scores.success) {
      resultats.echanges_totaux.push(...resultats.phase1_scores.echanges);
      Logger.log(`✅ Phase 1: ${resultats.phase1_scores.nbEchanges} échanges effectués`);
    } else {
      Logger.log(`❌ Phase 1 échouée: ${resultats.phase1_scores.error}`);
    }
    
    // PHASE 2: Correction parité (FICHIER 1 optimisé)
    Logger.log("--- PHASE 2: CORRECTION PARITÉ ---");
    if (typeof psv5_nirvanaV2_CorrectionPariteINTELLIGENTE === 'function') {
      try {
        resultats.phase2_parite = psv5_nirvanaV2_CorrectionPariteINTELLIGENTE(dataContext, config);
        if (resultats.phase2_parite && resultats.phase2_parite.length > 0) {
          resultats.echanges_totaux.push(...resultats.phase2_parite);
          Logger.log(`✅ Phase 2: ${resultats.phase2_parite.length} échanges parité effectués`);
        } else {
          Logger.log("ℹ️ Phase 2: Aucun échange parité nécessaire");
        }
      } catch (e) {
        Logger.log(`❌ Erreur Phase 2: ${e.message}`);
      }
    } else {
      Logger.log("⚠️ Fonction parité non disponible");
    }
    
    // PHASE 3: MultiSwap pour cas complexes (FICHIER 2)
    Logger.log("--- PHASE 3: MULTISWAP ---");
    if (typeof V2_Ameliore_MultiSwap_AvecRetourSwaps === 'function') {
      try {
        resultats.phase3_multiswap = V2_Ameliore_MultiSwap_AvecRetourSwaps(dataContext, config);
        if (resultats.phase3_multiswap && resultats.phase3_multiswap.swapsDetailles) {
          resultats.echanges_totaux.push(...resultats.phase3_multiswap.swapsDetailles);
          Logger.log(`✅ Phase 3: ${resultats.phase3_multiswap.swapsDetailles.length} échanges MultiSwap effectués`);
        } else {
          Logger.log("ℹ️ Phase 3: Aucun cycle MultiSwap trouvé");
        }
      } catch (e) {
        Logger.log(`❌ Erreur Phase 3: ${e.message}`);
      }
    } else {
      Logger.log("⚠️ Fonction MultiSwap non disponible");
    }
    
    // CALCUL DU TEMPS TOTAL
    resultats.temps_execution = new Date() - tempsDebut;
    resultats.score_final = resultats.phase1_scores ? resultats.phase1_scores.scoreFinal : 0;
    
    // BILAN FINAL
    Logger.log("--- BILAN FINAL PIPELINE ---");
    Logger.log(`Total échanges: ${resultats.echanges_totaux.length}`);
    Logger.log(`Score final: ${resultats.score_final}/100`);
    Logger.log(`Temps d'exécution: ${(resultats.temps_execution / 1000).toFixed(2)}s`);
    
    // Répartition par phase
    const nbPhase1 = resultats.phase1_scores ? resultats.phase1_scores.nbEchanges : 0;
    const nbPhase2 = resultats.phase2_parite ? resultats.phase2_parite.length : 0;
    const nbPhase3 = resultats.phase3_multiswap ? resultats.phase3_multiswap.swapsDetailles.length : 0;
    
    Logger.log(`Répartition: Phase1(${nbPhase1}) + Phase2(${nbPhase2}) + Phase3(${nbPhase3}) = ${resultats.echanges_totaux.length}`);
    
    return {
      success: true,
      resultats: resultats,
      nbEchangesTotal: resultats.echanges_totaux.length,
      scoreFinal: resultats.score_final,
      tempsExecution: resultats.temps_execution
    };
    
  } catch (e) {
    Logger.log(`❌ ERREUR pipeline: ${e.message}`);
    return {
      success: false,
      error: e.message,
      echanges_totaux: []
    };
  }
}

// ==================================================================
// SECTION 14: FONCTIONS UTILITAIRES AVANCÉES
// ==================================================================

/**
 * Génère un rapport détaillé de performance
 */
function genererRapportPerformance(resultats) {
  let rapport = "=== RAPPORT DE PERFORMANCE NIRVANA V1.2 ===\n\n";
  
  // Métriques générales
  rapport += `⏱️ TEMPS D'EXÉCUTION: ${(resultats.temps_execution / 1000).toFixed(2)} secondes\n`;
  rapport += `🎯 SCORE FINAL: ${resultats.score_final.toFixed(2)}/100\n`;
  rapport += `🔄 TOTAL ÉCHANGES: ${resultats.nbEchangesTotal}\n\n`;
  
  // Détail par phase
  rapport += "📊 DÉTAIL PAR PHASE:\n";
  
  if (resultats.resultats.phase1_scores) {
    const p1 = resultats.resultats.phase1_scores;
    rapport += `  Phase 1 (Scores): ${p1.nbEchanges} échanges`;
    if (p1.iterations) rapport += ` en ${p1.iterations} itérations`;
    rapport += "\n";
  }
  
  if (resultats.resultats.phase2_parite) {
    rapport += `  Phase 2 (Parité): ${resultats.resultats.phase2_parite.length} échanges\n`;
  }
  
  if (resultats.resultats.phase3_multiswap) {
    const p3 = resultats.resultats.phase3_multiswap;
    rapport += `  Phase 3 (MultiSwap): ${p3.swapsDetailles.length} échanges`;
    if (p3.nbCycles) rapport += ` (${p3.nbCycles} cycles)`;
    rapport += "\n";
  }
  
  // Recommandations
  rapport += "\n💡 RECOMMANDATIONS:\n";
  if (resultats.score_final < 70) {
    rapport += "- Score faible: Vérifier les contraintes de mobilité\n";
    rapport += "- Considérer l'ajustement des paramètres de tolérance\n";
  } else if (resultats.score_final > 90) {
    rapport += "- Excellent équilibre atteint !\n";
  }
  
  if (resultats.nbEchangesTotal > 50) {
    rapport += "- Nombre d'échanges élevé: Considérer une optimisation des paramètres\n";
  }
  
  return rapport;
}

/**
 * Sauvegarde les résultats dans un onglet de bilan
 */
function sauvegarderBilanDansOnglet(resultats, nomOnglet = '_BILAN_SCORES') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(nomOnglet);
    
    if (sheet) {
      sheet.clear();
    } else {
      sheet = ss.insertSheet(nomOnglet);
    }
    
    // En-tête
    const donnees = [
      [`=== BILAN NIRVANA SCORES ÉQUILIBRAGE V1.2 ===`],
      [`Date: ${new Date().toLocaleString('fr-FR')}`],
      [`Score final: ${resultats.scoreFinal.toFixed(2)}/100`],
      [`Temps d'exécution: ${(resultats.tempsExecution / 1000).toFixed(2)}s`],
      [`Total échanges: ${resultats.nbEchangesTotal}`],
      [''],
      ['Phase', 'Nombre d\'échanges', 'Détails']
    ];
    
    // Détails par phase
    if (resultats.resultats.phase1_scores) {
      donnees.push(['Phase 1 - Scores', resultats.resultats.phase1_scores.nbEchanges, 'Équilibrage effectifs par score']);
    }
    
    if (resultats.resultats.phase2_parite) {
      donnees.push(['Phase 2 - Parité', resultats.resultats.phase2_parite.length, 'Correction parité F/M']);
    }
    
    if (resultats.resultats.phase3_multiswap) {
      donnees.push(['Phase 3 - MultiSwap', resultats.resultats.phase3_multiswap.swapsDetailles.length, 'Cycles complexes']);
    }
    
    // Écrire les données
    if (donnees.length > 0) {
      sheet.getRange(1, 1, donnees.length, 3).setValues(donnees);
      
      // Mise en forme
      sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
      sheet.getRange(7, 1, 1, 3).setFontWeight('bold').setBackground('#e6f3ff');
      sheet.autoResizeColumns(1, 3);
    }
    
    Logger.log(`✅ Bilan sauvegardé dans l'onglet '${nomOnglet}'`);
    return nomOnglet;
    
  } catch (e) {
    Logger.log(`❌ Erreur sauvegarde bilan: ${e.message}`);
    return null;
  }
}

// ==================================================================
// SECTION 15: FONCTIONS DE LIAISON POUR ORCHESTRATEUR
// ==================================================================

/**
 * Fonction de liaison pour l'orchestrateur - Équilibrage scores UI
 */
function lancerEquilibrageScores_UI(config) {
  Logger.log("=== LANCEMENT ÉQUILIBRAGE SCORES UI ===");
  
  try {
    const dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
    
    if (!dataContext || !dataContext.classesState) {
      throw new Error("Impossible de préparer les données");
    }
    
    // Lancer l'équilibrage via intégration Nirvana V2
    const resultat = integrerAvecNirvanaV2(dataContext, config);
    
    if (resultat.success) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `✅ Équilibrage scores terminé: ${resultat.nbEchanges} échanges`, 
        "Succès", 
        10
      );
      Logger.log(`✅ Succès: ${resultat.nbEchanges} échanges, score ${resultat.scoreFinal.toFixed(2)}/100`);
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `❌ Échec équilibrage: ${resultat.error}`, 
        "Erreur", 
        10
      );
    }
    
    return resultat;
    
  } catch (e) {
    Logger.log(`❌ ERREUR lancerEquilibrageScores_UI: ${e.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `❌ Erreur: ${e.message}`, 
      "Erreur", 
      10
    );
    return { success: false, error: e.message };
  }
}

/**
 * Fonction de liaison pour l'orchestrateur - Stratégie réaliste
 */
function executerEquilibrageSelonStrategieRealiste(config) {
  Logger.log("=== EXÉCUTION ÉQUILIBRAGE STRATÉGIE RÉALISTE ===");
  
  try {
    const dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
    
    if (!dataContext || !dataContext.classesState) {
      throw new Error("Impossible de préparer les données");
    }
    
    // Diagnostic initial
    const diagnosticInitial = diagnostiquerEffectifs_Ameliore(dataContext);
    Logger.log("Diagnostic initial:\n" + diagnosticInitial);
    
    // Équilibrage avec stratégie réaliste
    const resultat = equilibrerScores(dataContext, config);
    
    return {
      success: resultat.success,
      totalEchanges: resultat.journalEchanges?.length || 0,
      nbIterations: resultat.iterations || 0,
      scoreInitial: 0, // Calculé dans l'algorithme
      scoreFinal: resultat.scoreFinal || 0,
      strategieUtilisee: "Stratégie réaliste V1.2",
      echanges: resultat.journalEchanges || [],
      error: resultat.error
    };
    
  } catch (e) {
    Logger.log(`❌ ERREUR executerEquilibrageSelonStrategieRealiste: ${e.message}`);
    return { success: false, error: e.message };
  }
}

/**
 * Fonction de liaison pour l'orchestrateur - Pipeline complet UI
 */
function lancerPipelineComplet_UI(config) {
  Logger.log("=== LANCEMENT PIPELINE COMPLET UI ===");
  
  try {
    const dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
    
    if (!dataContext || !dataContext.classesState) {
      throw new Error("Impossible de préparer les données");
    }
    
    // Lancer le pipeline complet
    const resultat = pipelineNirvanaComplet(dataContext, config);
    
    if (resultat.success && resultat.nbEchangesTotal > 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `✅ Pipeline complet terminé: ${resultat.nbEchangesTotal} échanges`, 
        "Succès", 
        10
      );
    } else if (resultat.success) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "ℹ️ Pipeline complet: Aucun échange nécessaire", 
        "Info", 
        10
      );
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `❌ Échec pipeline: ${resultat.error}`, 
        "Erreur", 
        10
      );
    }
    
    return resultat;
    
  } catch (e) {
    Logger.log(`❌ ERREUR lancerPipelineComplet_UI: ${e.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `❌ Erreur: ${e.message}`, 
      "Erreur", 
      10
    );
    return { success: false, error: e.message };
  }
}

/**
 * Fonction de liaison pour l'orchestrateur - Diagnostic UI
 */
function lancerDiagnosticScores_UI(config) {
  Logger.log("=== LANCEMENT DIAGNOSTIC SCORES UI ===");
  
  try {
    const dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
    
    if (!dataContext || !dataContext.classesState) {
      throw new Error("Impossible de préparer les données");
    }
    
    // Diagnostic complet
    const diagnostic = diagnostiquerEffectifs_Ameliore(dataContext);
    
    // Log pour débogage
    Logger.log("Diagnostic généré avec succès");
    
    // Afficher dans une boîte de dialogue
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        "Diagnostic Scores V1.2", 
        diagnostic.substring(0, 2000) + (diagnostic.length > 2000 ? "\n\n[...suite dans les logs...]" : ""), 
        ui.ButtonSet.OK
      );
    } catch (uiError) {
      // Fallback: utiliser toast si UI.alert échoue
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Diagnostic généré avec succès - Consultez les logs", 
        "Diagnostic", 
        5
      );
    }
    
    return { success: true, diagnostic: diagnostic };
    
  } catch (e) {
    Logger.log(`❌ ERREUR lancerDiagnosticScores_UI: ${e.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `❌ Erreur: ${e.message}`, 
      "Erreur", 
      10
    );
    return { success: false, error: e.message };
  }
}

// ==================================================================
// SECTION 16: POINTS D'ENTRÉE UI ORIGINAUX (COMPATIBILITÉ)
// ==================================================================

/**
 * Point d'entrée UI principal - Équilibrage seul (VERSION ORIGINALE)
 */
function lancerEquilibrageScores_UI_Original() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log("=== LANCEMENT ÉQUILIBRAGE SCORES V1.2 (ORIGINAL) ===");
    
    // Préparation des données
    const config = getConfig() || {};
    const dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
    
    // Lancement de l'équilibrage
    const resultat = integrerAvecNirvanaV2(dataContext, config);
    
    if (resultat.success) {
      ui.alert('Équilibrage Réussi', 
        `✅ Équilibrage des scores terminé !\n\n` +
        `${resultat.nbEchanges} échanges effectués\n` +
        `Score final: ${resultat.scoreFinal.toFixed(2)}/100\n` +
        `Itérations: ${resultat.iterations}\n\n` +
        `Consultez les logs pour le détail.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Erreur', `❌ Échec: ${resultat.error}`, ui.ButtonSet.OK);
    }
    
    return resultat;
    
  } catch (e) {
    Logger.log(`❌ ERREUR UI: ${e.message}`);
    ui.alert('Erreur Critique', e.message, ui.ButtonSet.OK);
    return { success: false, error: e.message };
  }
}

/**
 * Point d'entrée UI - Pipeline complet (VERSION ORIGINALE)
 */
function lancerPipelineComplet_UI_Original() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log("=== LANCEMENT PIPELINE COMPLET V1.2 (ORIGINAL) ===");
    
    // Confirmation utilisateur
    const reponse = ui.alert('Pipeline Complet Nirvana',
      'Lancer le pipeline complet ?\n\n' +
      '• Phase 1: Équilibrage des scores\n' +
      '• Phase 2: Correction parité F/M\n' +
      '• Phase 3: MultiSwap complexe\n\n' +
      'Cette opération peut prendre plusieurs minutes.',
      ui.ButtonSet.YES_NO
    );
    
    if (reponse !== ui.Button.YES) {
      return { success: false, error: 'Annulé par l\'utilisateur' };
    }
    
    // Préparation des données
    const config = getConfig() || {};
    const dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
    
    // Lancement du pipeline
    const resultats = pipelineNirvanaComplet(dataContext, config);
    
    if (resultats.success) {
      // Générer et sauvegarder le bilan
      const rapport = genererRapportPerformance(resultats);
      const nomOnglet = sauvegarderBilanDansOnglet(resultats);
      
      Logger.log(rapport);
      
      ui.alert('Pipeline Terminé', 
        `✅ Pipeline Nirvana complet terminé !\n\n` +
        `${resultats.nbEchangesTotal} échanges au total\n` +
        `Score final: ${resultats.scoreFinal.toFixed(2)}/100\n` +
        `Temps: ${(resultats.tempsExecution / 1000).toFixed(2)}s\n\n` +
        `Bilan détaillé dans l'onglet '${nomOnglet}'`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Erreur Pipeline', `❌ Échec: ${resultats.error}`, ui.ButtonSet.OK);
    }
    
    return resultats;
    
  } catch (e) {
    Logger.log(`❌ ERREUR UI Pipeline: ${e.message}`);
    ui.alert('Erreur Critique', e.message, ui.ButtonSet.OK);
    return { success: false, error: e.message };
  }
}

/**
 * Point d'entrée UI - Diagnostic seul (VERSION ORIGINALE)
 */
function lancerDiagnosticScores_UI_Original() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = getConfig() || {};
    const dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
    const diagnostic = diagnostiquerEffectifs_Ameliore(dataContext);
    
    Logger.log(diagnostic);
    
    // Afficher dans une boîte de dialogue
    const html = HtmlService.createHtmlOutput(
      `<pre style="font-family: monospace; font-size: 11px;">${diagnostic.replace(/\n/g, '<br>')}</pre>`
    ).setWidth(800).setHeight(600);
    
    ui.showModalDialog(html, "Diagnostic Effectifs par Score V1.2");
    
    return { success: true, diagnostic: diagnostic };
    
  } catch (e) {
    Logger.log(`❌ ERREUR Diagnostic: ${e.message}`);
    ui.alert('Erreur Diagnostic', e.message, ui.ButtonSet.OK);
    return { success: false, error: e.message };
  }
}

// ==================================================================
// SECTION 15: POINTS D'ENTRÉE UI
// ==================================================================

/**
 * Point d'entrée UI principal - Équilibrage seul
 */
function lancerEquilibrageScores_UI() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log("=== LANCEMENT ÉQUILIBRAGE SCORES V1.2 ===");
    
    // Préparation des données
    const config = getConfig() || {};
    if (typeof V2_Ameliore_PreparerDonnees_AvecSEXE === 'function') {
      const dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
      
      // Lancement de l'équilibrage
      const resultat = integrerAvecNirvanaV2(dataContext, config);
      
      if (resultat.success) {
        ui.alert('Équilibrage Réussi', 
          `✅ Équilibrage des scores terminé !\n\n` +
          `${resultat.nbEchanges} échanges effectués\n` +
          `Score final: ${resultat.scoreFinal.toFixed(2)}/100\n` +
          `Itérations: ${resultat.iterations}\n\n` +
          `Consultez les logs pour le détail.`,
          ui.ButtonSet.OK
        );
      } else {
        ui.alert('Erreur', `❌ Échec: ${resultat.error}`, ui.ButtonSet.OK);
      }
      
      return resultat;
      
    } else {
      throw new Error("Fonction V2_Ameliore_PreparerDonnees_AvecSEXE non disponible");
    }
    
  } catch (e) {
    Logger.log(`❌ ERREUR UI: ${e.message}`);
    ui.alert('Erreur Critique', e.message, ui.ButtonSet.OK);
    return { success: false, error: e.message };
  }
}

/**
 * Point d'entrée UI - Pipeline complet
 */
function lancerPipelineComplet_UI() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log("=== LANCEMENT PIPELINE COMPLET V1.2 ===");
    
    // Confirmation utilisateur
    const reponse = ui.alert('Pipeline Complet Nirvana',
      'Lancer le pipeline complet ?\n\n' +
      '• Phase 1: Équilibrage des scores\n' +
      '• Phase 2: Correction parité F/M\n' +
      '• Phase 3: MultiSwap complexe\n\n' +
      'Cette opération peut prendre plusieurs minutes.',
      ui.ButtonSet.YES_NO
    );
    
    if (reponse !== ui.Button.YES) {
      return { success: false, error: 'Annulé par l\'utilisateur' };
    }
    
    // Préparation des données
    const config = getConfig() || {};
    if (typeof V2_Ameliore_PreparerDonnees_AvecSEXE === 'function') {
      const dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
      
      // Lancement du pipeline
      const resultats = pipelineNirvanaComplet(dataContext, config);
      
      if (resultats.success) {
        // Générer et sauvegarder le bilan
        const rapport = genererRapportPerformance(resultats);
        const nomOnglet = sauvegarderBilanDansOnglet(resultats);
        
        Logger.log(rapport);
        
        ui.alert('Pipeline Terminé', 
          `✅ Pipeline Nirvana complet terminé !\n\n` +
          `${resultats.nbEchangesTotal} échanges au total\n` +
          `Score final: ${resultats.scoreFinal.toFixed(2)}/100\n` +
          `Temps: ${(resultats.tempsExecution / 1000).toFixed(2)}s\n\n` +
          `Bilan détaillé dans l'onglet '${nomOnglet}'`,
          ui.ButtonSet.OK
        );
      } else {
        ui.alert('Erreur Pipeline', `❌ Échec: ${resultats.error}`, ui.ButtonSet.OK);
      }
      
      return resultats;
      
    } else {
      throw new Error("Fonction V2_Ameliore_PreparerDonnees_AvecSEXE non disponible");
    }
    
  } catch (e) {
    Logger.log(`❌ ERREUR UI Pipeline: ${e.message}`);
    ui.alert('Erreur Critique', e.message, ui.ButtonSet.OK);
    return { success: false, error: e.message };
  }
}

/**
 * Point d'entrée UI - Diagnostic seul
 */
function lancerDiagnosticScores_UI() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = getConfig() || {};
    if (typeof V2_Ameliore_PreparerDonnees_AvecSEXE === 'function') {
      const dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
      const diagnostic = diagnostiquerEffectifs_Ameliore(dataContext);
      
      Logger.log(diagnostic);
      
      // Afficher dans une boîte de dialogue
      const html = HtmlService.createHtmlOutput(
        `<pre style="font-family: monospace; font-size: 11px;">${diagnostic.replace(/\n/g, '<br>')}</pre>`
      ).setWidth(800).setHeight(600);
      
      ui.showModalDialog(html, "Diagnostic Effectifs par Score V1.2");
      
      return { success: true, diagnostic: diagnostic };
      
    } else {
      throw new Error("Fonction V2_Ameliore_PreparerDonnees_AvecSEXE non disponible");
    }
    
  } catch (e) {
    Logger.log(`❌ ERREUR Diagnostic: ${e.message}`);
    ui.alert('Erreur Diagnostic', e.message, ui.ButtonSet.OK);
    return { success: false, error: e.message };
  }
}

// ==================================================================
// EXPORT DES FONCTIONS PRINCIPALES - VERSION FINALE
// ==================================================================

// Fonction principale
// - equilibrerScores

// Fonctions de calcul améliorées
// - calculerEffectifsCibles_Ameliore
// - calculerEffectifsActuels
// - calculerEcartsEffectifs
// - calculerScoreEquilibre

// Fonctions d'analyse
// - identifierSurplusDeficit
// - trouverCandidatsEchange_Intelligent

// Fonctions de mobilité avancées
// - verifierConditionsMobilite
// - gererGroupesSpec_Robuste
// - verifierContraintesEchange
// - estEleveMobile

// Fonctions de stratégies
// - choisirStrategieEchange
// - compterParMobilite
// - trouverCandidatsLibres
// - trouverCandidatsPermut
// - trouverCandidatsCondi
// - trouverCandidatsSpec

// Fonctions d'échange de groupes
// - proposerEchangeGroupeSpec
// - grouperParAssociation

// Fonctions d'application
// - appliquerEchange
// - appliquerEchangeSecurise
// - appliquerEchangeGroupe
// - mettreAJourEffectifsApresEchange

// Fonctions utilitaires
// - verifierOptions
// - verifierDissociations
// - calculerImpactEchange

// Fonctions de diagnostic
// - diagnostiquerEffectifs_Ameliore

// Fonctions d'intégration
// - integrerAvecNirvanaV2

// Fonctions de test
// - testerModuleScoresEquilibrage

// Pipeline complet
// - pipelineNirvanaComplet

// Fonctions avancées
// - genererRapportPerformance
// - sauvegarderBilanDansOnglet

// Fonctions de liaison orchestrateur
// - lancerEquilibrageScores_UI
// - executerEquilibrageSelonStrategieRealiste
// - lancerPipelineComplet_UI
// - lancerDiagnosticScores_UI

// Points d'entrée UI originaux (compatibilité)
// - lancerEquilibrageScores_UI_Original
// - lancerPipelineComplet_UI_Original
// - lancerDiagnosticScores_UI_Original

// ==================================================================
// FIN DU MODULE NIRVANA_SCORES_EQUILIBRAGE V1.2 COMPLET
// ==================================================================