function verifierEquilibreActuel() {
    const config = getConfig();
    const result = chargerElevesEtClassesCorrige(config, "MOBILITE");
    
    if (!result.success) return;
    
    // Calculer les stats
    classifierEleves(result.students, []);
    const stats = calculerStatistiquesDistribution(result.classesMap, result.students.length, []);
    
    Logger.log("=== ÉTAT ACTUEL DE L'ÉQUILIBRE ===");
    Logger.log(`Écart-type têtes: ${stats.tetesDeClasse.ecartType.toFixed(4)}`);
    Logger.log(`Écart-type niveau1: ${stats.niveau1.ecartType.toFixed(4)}`);
    Logger.log(`Écart distribution: ${stats.distribution.ecartMoyen.toFixed(4)}`);
    
    if (stats.tetesDeClasse.ecartType < 0.5 && 
        stats.niveau1.ecartType < 0.5 && 
        stats.distribution.ecartMoyen < 0.1) {
        Logger.log("\n✅ Les classes sont DÉJÀ très bien équilibrées !");
        Logger.log("C'est normal qu'aucun swap ne soit trouvé.");
    } else {
        Logger.log("\n⚠️ Il y a du déséquilibre, des swaps devraient être possibles");
    }
}
function testMoteurSeuilTresBas() {
    Logger.log("=== TEST AVEC SEUIL ULTRA-BAS ===");
    
    // Dans V11_OptimisationDistribution_Combined, trouvez et modifiez:
    // const SEUIL_IMPACT_MINIMAL = 1e-6;
    // Remplacez temporairement par:
    // const SEUIL_IMPACT_MINIMAL = -1000; // Accepte même les impacts négatifs pour tester
    
    const poids = {
        tetesDeClasse: 10.0,  // Très élevé
        niveau1: 10.0,        // Très élevé
        distribution: 10.0,    // Très élevé
        com1: 0, tra4: 0, part4: 0,
        garantieTete: 1000
    };
    
    const resultat = V11_OptimisationDistribution_Combined(null, poids, 30);
    
    Logger.log("Résultat avec seuil très bas:");
    Logger.log("- Success: " + resultat.success);
    Logger.log("- Nb Swaps: " + resultat.nbSwaps);
    Logger.log("- Message: " + resultat.message);
    
    return resultat;
}

function verifierVariabilitesDonnees() {
    Logger.log("=== VÉRIFICATION VARIABILITÉ DES DONNÉES ===");
    
    const config = getConfig();
    const result = chargerElevesEtClassesCorrige(config, "MOBILITE");
    
    if (!result.success || !result.students) {
        Logger.log("Erreur chargement");
        return;
    }
    
    // Analyser la distribution des valeurs
    const distribution = {
        COM: { 1: 0, 2: 0, 3: 0, 4: 0, vides: 0 },
        TRA: { 1: 0, 2: 0, 3: 0, 4: 0, vides: 0 },
        PART: { 1: 0, 2: 0, 3: 0, 4: 0, vides: 0 }
    };
    
    result.students.forEach(eleve => {
        // COM
        if (eleve.COM === '') distribution.COM.vides++;
        else if ([1,2,3,4].includes(eleve.COM)) distribution.COM[eleve.COM]++;
        
        // TRA
        if (eleve.TRA === '') distribution.TRA.vides++;
        else if ([1,2,3,4].includes(eleve.TRA)) distribution.TRA[eleve.TRA]++;
        
        // PART
        if (eleve.PART === '') distribution.PART.vides++;
        else if ([1,2,3,4].includes(eleve.PART)) distribution.PART[eleve.PART]++;
    });
    
    // Afficher la distribution
    ['COM', 'TRA', 'PART'].forEach(critere => {
        Logger.log(`\n${critere}:`);
        [1,2,3,4].forEach(score => {
            Logger.log(`  Score ${score}: ${distribution[critere][score]} élèves`);
        });
        Logger.log(`  Vides: ${distribution[critere].vides}`);
    });
    
    // Classifier et compter
    classifierEleves(result.students, ["com1", "tra4", "part4"]);
    
    const tetesDeClasse = result.students.filter(e => e.estTeteDeClasse).length;
    const niveau1 = result.students.filter(e => e.estNiveau1).length;
    
    Logger.log(`\n=== RÉSUMÉ ===`);
    Logger.log(`Têtes de classe (au moins un 4): ${tetesDeClasse}`);
    Logger.log(`Niveau 1 (au moins un 1): ${niveau1}`);
    
    // Vérifier l'équilibre par classe
    Logger.log("\n=== ÉQUILIBRE PAR CLASSE ===");
    Object.entries(result.classesMap).forEach(([classe, eleves]) => {
        const tetes = eleves.filter(e => {
            e.niveauCOM = getNiveau(e.COM);
            e.niveauTRA = getNiveau(e.TRA);
            e.niveauPART = getNiveau(e.PART);
            return (e.niveauCOM === 4 || e.niveauTRA === 4 || e.niveauPART === 4);
        }).length;
        
        Logger.log(`${classe}: ${eleves.length} élèves, ${tetes} têtes`);
    });
}

/**
 * OPTIMISATION PARITÉ M/F POUR MOTEUR V14 - VERSION 2 AMÉLIORÉE
 * Intègre les retours : swaps asymétriques, meilleure gestion des effectifs,
 * paramètres configurables, et optimisations de performance
 * À ajouter dans Optimisation_V14.gs
 */

// =================================================
// 0. PARAMÈTRES CONFIGURABLES
// =================================================
const PARITE_CONFIG = {
  // Seuils de tolérance
  TOLERANCE_EFFECTIF: 1.5,      // ±1.5 élève par classe
  TOLERANCE_PARITE: 0.05,       // ±5% du ratio F idéal
  SEUIL_SIMILARITE: 0.5,        // Similarité minimale des scores (abaissé)
  
  // Pondérations pour le calcul d'impact
  POIDS_PARITE: 0.7,            // Importance de l'équilibrage M/F
  POIDS_EFFECTIFS: 0.3,         // Importance de l'équilibrage des effectifs
  
  // Pénalités similarité
  PENALITE_NIVEAU: 0.08,        // Par niveau de différence (réduit de 0.05 à 0.08)
  PENALITE_MOYENNE: 0.05,       // Par point de différence moyenne (réduit)
  PENALITE_TETE: 0.10,          // Si statuts tête différents (réduit de 0.15)
  PENALITE_NIV1: 0.05,          // Si statuts niveau1 différents (réduit de 0.10)
  BONUS_IDENTIQUE: 0.15,        // Si niveaux identiques
  
  // Options de swaps
  AUTORISER_SWAPS_ASYMETRIQUES: true,  // Permet 1↔2 ou 2↔1
  MAX_ASYMETRIE: 1,                     // Différence max dans un swap asymétrique
  
  // Performance
  MAX_ITERATIONS_PAR_PAIRE: 50,         // Limite de candidats à examiner
  SEUIL_IMPACT_MINIMAL: 0.05            // Impact minimal pour considérer un swap
};

// Cache pour éviter les appels répétés
let _pariteConfigCache = null;

function getPariteConfig() {
  if (_pariteConfigCache) {
    return _pariteConfigCache;
  }
  
  try {
    const config = getConfig();
    if (config && config.PARITE_CONFIG) {
      _pariteConfigCache = { ...PARITE_CONFIG, ...config.PARITE_CONFIG };
    } else {
      _pariteConfigCache = PARITE_CONFIG;
    }
  } catch (e) {
    _pariteConfigCache = PARITE_CONFIG;
  }
  
  return _pariteConfigCache;
}

// =================================================
// 1. FONCTION PRINCIPALE D'OPTIMISATION PARITÉ V2
// =================================================
function optimiserPariteMF_V2(maxSwapsParite = 50) {
  Logger.log("==========================================================");
  Logger.log("OPTIMISATION PARITÉ M/F V2 - SWAPS SYMÉTRIQUES ET ASYMÉTRIQUES");
  Logger.log("==========================================================");
  
  const startTime = new Date();
  const config = getPariteConfig();
  
  try {
    // 1. Charger et préparer les données
    const donnees = preparerDonneesParite();
    if (!donnees.success) {
      return donnees;
    }
    
    const { students, structure, optionsNiveau, optionPools, dissocMap, niveau } = donnees;
    let { classesMap } = donnees; // Sera modifié dans la boucle
    
    // 2. Analyser la situation initiale
    Logger.log("\n=== ANALYSE INITIALE ===");
    const statsInitiales = analyserPariteEtEffectifs(classesMap);
    afficherStatsCompletes(statsInitiales);
    
    // 3. Calculer les cibles et écarts
    const cibles = calculerCiblesIdeales(statsInitiales);
    afficherCiblesIdeales(cibles);
    
    // 4. Boucle itérative de génération des swaps
    Logger.log("\n=== GÉNÉRATION ITÉRATIVE DES SWAPS ===");
    
    const swapsAccumules = [];
    const elevesDejaSwappes = new Set();
    let iteration = 0;
    let continuer = true;
    
    while (continuer && swapsAccumules.length < maxSwapsParite && iteration < 10) {
      iteration++;
      Logger.log(`\n--- Itération ${iteration} ---`);
      
      // Recalculer les écarts avec la situation actuelle
      const statsActuelles = analyserPariteEtEffectifs(classesMap);
      const ecartsActuels = identifierEcartsCibles(statsActuelles, cibles);
      
      // Vérifier si on a encore des déséquilibres importants
      const maxEcartF = Math.max(...ecartsActuels.map(e => Math.abs(e.ecartF)));
      Logger.log(`Écart F maximum actuel: ${maxEcartF.toFixed(1)}`);
      
      if (maxEcartF < 2) {
        Logger.log("✅ Équilibre satisfaisant atteint !");
        continuer = false;
        break;
      }
      
      const tousLesSwaps = [];
      
      // Swaps symétriques (1↔1)
      if (ecartsActuels.length >= 2) {
        const swapsSymetriques = genererSwapsSymetriques(
          students, classesMap, ecartsActuels, cibles,
          structure, optionsNiveau, optionPools, dissocMap
        );
        tousLesSwaps.push(...swapsSymetriques);
        Logger.log(`${swapsSymetriques.length} swaps symétriques trouvés`);
      }
      
      // Swaps asymétriques si autorisés
      if (config.AUTORISER_SWAPS_ASYMETRIQUES) {
        const swapsAsymetriques = genererSwapsAsymetriques(
          students, classesMap, ecartsActuels, cibles,
          structure, optionsNiveau, optionPools, dissocMap
        );
        tousLesSwaps.push(...swapsAsymetriques);
        Logger.log(`${swapsAsymetriques.length} swaps asymétriques trouvés`);
      }
      
      // Filtrer les swaps qui n'utilisent pas d'élèves déjà swappés
      const swapsDisponibles = tousLesSwaps.filter(swap => {
        const ids = [
          ...swap.eleves1.map(e => e.ID_ELEVE),
          ...swap.eleves2.map(e => e.ID_ELEVE)
        ];
        return !ids.some(id => elevesDejaSwappes.has(id));
      });
      
      if (swapsDisponibles.length === 0) {
        Logger.log("Plus de swaps disponibles sans conflit");
        continuer = false;
        break;
      }
      
      // Sélectionner les meilleurs pour cette itération
      const swapsIterations = selectionnerMeilleursSwaps(
        swapsDisponibles, 
        Math.min(10, maxSwapsParite - swapsAccumules.length) // Max 10 par itération
      );
      
      if (swapsIterations.length === 0) {
        continuer = false;
        break;
      }
      
      // Simuler l'application des swaps pour la prochaine itération
      swapsIterations.forEach(swap => {
        // Mettre à jour classesMap temporairement
        swap.eleves1.forEach(e => {
          classesMap[swap.classe1] = classesMap[swap.classe1].filter(el => el.ID_ELEVE !== e.ID_ELEVE);
          e.CLASSE = swap.classe2;
          classesMap[swap.classe2].push(e);
        });
        swap.eleves2.forEach(e => {
          classesMap[swap.classe2] = classesMap[swap.classe2].filter(el => el.ID_ELEVE !== e.ID_ELEVE);
          e.CLASSE = swap.classe1;
          classesMap[swap.classe1].push(e);
        });
        
        // Marquer les élèves comme utilisés
        [...swap.eleves1, ...swap.eleves2].forEach(e => elevesDejaSwappes.add(e.ID_ELEVE));
      });
      
      swapsAccumules.push(...swapsIterations);
      Logger.log(`${swapsIterations.length} swaps ajoutés (total: ${swapsAccumules.length})`);
    }
    
    Logger.log(`\n=== FIN GÉNÉRATION: ${swapsAccumules.length} swaps en ${iteration} itérations ===`);
    
    if (swapsAccumules.length === 0) {
      Logger.log("❌ Aucun swap possible trouvé");
      return {
        success: true,
        nbSwaps: 0,
        message: "Aucun swap possible pour améliorer la répartition",
        statsInitiales,
        statsFinales: statsInitiales
      };
    }
    
    const swapsOptimaux = swapsAccumules;
    
    // 6. Afficher et confirmer
    afficherSwapsProposés(swapsOptimaux);
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Optimisation Parité M/F V2',
      `${swapsOptimaux.length} échanges proposés :\n` +
      `- ${swapsOptimaux.filter(s => s.type === 'symetrique').length} symétriques (1↔1)\n` +
      `- ${swapsOptimaux.filter(s => s.type === 'asymetrique').length} asymétriques\n\n` +
      `Voulez-vous les appliquer ?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      Logger.log("Optimisation annulée par l'utilisateur");
      return {
        success: true,
        nbSwaps: 0,
        message: "Optimisation annulée",
        statsInitiales,
        statsFinales: statsInitiales
      };
    }
    
    // 7. Exécuter les swaps (ATTENTION: recharger les données originales car on a modifié classesMap)
    Logger.log("\n=== EXÉCUTION DES SWAPS ===");
    // Recharger les données originales
    const donneesOriginales = preparerDonneesParite();
    const resultatExecution = executerSwapsOptimises(swapsOptimaux, niveau);
    
    if (!resultatExecution.success) {
      return resultatExecution;
    }
    
    // 8. Analyser les résultats
    const newClassesMap = rechargerClassesMap();
    const statsFinales = analyserPariteEtEffectifs(newClassesMap);
    
    Logger.log("\n=== RÉSULTATS OPTIMISATION ===");
    afficherResultatsOptimisation(statsInitiales, statsFinales, cibles);
    
    const endTime = new Date();
    const duree = (endTime.getTime() - startTime.getTime()) / 1000;
    
    Logger.log(`\n✅ Optimisation terminée en ${duree.toFixed(1)}s`);
    
    ui.alert(
      'Optimisation Parité Terminée',
      `${resultatExecution.nbSwaps} échanges effectués :\n` +
      `- Réduction écart-type parité : ${((statsInitiales.global.ecartTypeParite - statsFinales.global.ecartTypeParite) * 100).toFixed(1)}%\n` +
      `- Réduction écart-type effectifs : ${(statsInitiales.global.ecartTypeEffectifs - statsFinales.global.ecartTypeEffectifs).toFixed(2)}\n\n` +
      `La répartition M/F et les effectifs ont été optimisés.`,
      ui.ButtonSet.OK
    );
    
    return {
      success: true,
      nbSwaps: resultatExecution.nbSwaps,
      message: `${resultatExecution.nbSwaps} swaps effectués`,
      statsInitiales,
      statsFinales,
      journal: resultatExecution.journal,
      dureeMs: endTime.getTime() - startTime.getTime()
    };
    
  } catch (error) {
    Logger.log(`❌ Erreur optimisation parité V2: ${error.message}`);
    Logger.log(error.stack);
    return { success: false, message: "Erreur: " + error.message };
  }
}

// =================================================
// 2. PRÉPARATION DES DONNÉES
// =================================================
function preparerDonneesParite() {
  const config = getConfig();
  const niveau = determinerNiveauActifCache();
  
  // Charger les élèves
  const elevesResult = chargerElevesEtClasses(config, "MOBILITE");
  if (!elevesResult.success) {
    return { success: false, message: "Erreur chargement élèves" };
  }
  
  // Patch classesMap si nécessaire
  let { students, classesMap } = elevesResult;
  if (!classesMap && students) {
    classesMap = {};
    students.forEach(eleve => {
      if (eleve && eleve.CLASSE) {
        if (!classesMap[eleve.CLASSE]) classesMap[eleve.CLASSE] = [];
        classesMap[eleve.CLASSE].push(eleve);
      }
    });
  }
  
  // Classifier les élèves
  classifierEleves(students, ["com1", "tra4", "part4"]);
  
  // Charger la structure et options
  const structureResult = chargerStructureEtOptions(niveau, config);
  if (!structureResult.success) {
    return { success: false, message: "Erreur chargement structure" };
  }
  
  const { structure, optionsNiveau } = structureResult;
  const optionPools = buildOptionPools(structure, config);
  const dissocMap = buildDissocCountMap(classesMap);
  
  return {
    success: true,
    students,
    classesMap,
    structure,
    optionsNiveau,
    optionPools,
    dissocMap,
    niveau  // AJOUT du niveau
  };
}

// =================================================
// 3. GÉNÉRATION DES SWAPS SYMÉTRIQUES (1↔1)
// =================================================
function genererSwapsSymetriques(students, classesMap, ecarts, cibles, structure, optionsNiveau, optionPools, dissocMap) {
  const swaps = [];
  const config = getPariteConfig();
  
  // Optimisation : limiter les paires à examiner
  const classesPrioritaires = ecarts.slice(0, 8); // Top 8 des plus écartées
  
  for (let i = 0; i < classesPrioritaires.length - 1; i++) {
    for (let j = i + 1; j < classesPrioritaires.length; j++) {
      const ecart1 = classesPrioritaires[i];
      const ecart2 = classesPrioritaires[j];
      
      // Déterminer les scénarios d'échange bénéfiques
      const scenarios = determinerScenariosEchange(ecart1, ecart2);
      
      for (const scenario of scenarios) {
        const candidatsSwap = trouverCandidatsSymetriques(
          ecart1, ecart2, scenario,
          classesMap, students,
          structure, optionsNiveau, optionPools, dissocMap,
          cibles  // AJOUT du paramètre cibles
        );
        
        // Garder seulement les meilleurs
        const meilleurs = candidatsSwap
          .filter(swap => swap.impactGlobal > config.SEUIL_IMPACT_MINIMAL)
          .sort((a, b) => b.impactGlobal - a.impactGlobal)
          .slice(0, 3); // Top 3 par paire de classes
        
        swaps.push(...meilleurs);
      }
    }
  }
  
  return swaps;
}

// =================================================
// 4. GÉNÉRATION DES SWAPS ASYMÉTRIQUES (1↔2, 2↔1)
// =================================================
function genererSwapsAsymetriques(students, classesMap, ecarts, cibles, structure, optionsNiveau, optionPools, dissocMap) {
  const swaps = [];
  const config = getPariteConfig();
  
  // Chercher les classes avec gros écarts d'effectifs
  const classesGrosEcart = ecarts.filter(e => 
    Math.abs(e.ecartEffectif) > config.TOLERANCE_EFFECTIF
  );
  
  if (classesGrosEcart.length < 2) {
    return swaps; // Pas assez de déséquilibre pour justifier l'asymétrie
  }
  
  // Séparer classes trop grandes et trop petites
  const classesGrandes = classesGrosEcart.filter(e => e.ecartEffectif > config.TOLERANCE_EFFECTIF);
  const classesPetites = classesGrosEcart.filter(e => e.ecartEffectif < -config.TOLERANCE_EFFECTIF);
  
  // Scénario 1 : 2 élèves d'une grande classe → 1 élève d'une petite classe
  for (const grande of classesGrandes) {
    for (const petite of classesPetites) {
      const candidats2vers1 = trouverCandidatsAsymetriques(
        grande, petite, "2vers1",
        classesMap, students,
        structure, optionsNiveau, optionPools, dissocMap,
        cibles
      );
      
      swaps.push(...candidats2vers1.filter(s => s.impactGlobal > config.SEUIL_IMPACT_MINIMAL));
    }
  }
  
  // Scénario 2 : 1 élève d'une petite classe → 2 élèves d'une grande classe
  for (const petite of classesPetites) {
    for (const grande of classesGrandes) {
      const candidats1vers2 = trouverCandidatsAsymetriques(
        petite, grande, "1vers2",
        classesMap, students,
        structure, optionsNiveau, optionPools, dissocMap,
        cibles
      );
      
      swaps.push(...candidats1vers2.filter(s => s.impactGlobal > config.SEUIL_IMPACT_MINIMAL));
    }
  }
  
  return swaps;
}

// =================================================
// 5. TROUVER CANDIDATS POUR SWAPS SYMÉTRIQUES
// =================================================
function trouverCandidatsSymetriques(ecart1, ecart2, scenario, classesMap, students, structure, optionsNiveau, optionPools, dissocMap, cibles) {
  const candidats = [];
  const config = getPariteConfig(); // Une seule fois au début
  
  const eleves1 = classesMap[ecart1.classe].filter(e =>
    e.SEXE === scenario.sexe1 &&
    (e.mobilite || 'LIBRE') !== 'FIXE' &&
    (e.mobilite || 'LIBRE') !== 'SPEC'
  );
  
  const eleves2 = classesMap[ecart2.classe].filter(e =>
    e.SEXE === scenario.sexe2 &&
    (e.mobilite || 'LIBRE') !== 'FIXE' &&
    (e.mobilite || 'LIBRE') !== 'SPEC'
  );
  
  // Limiter le nombre de combinaisons pour performance
  const maxCandidats1 = Math.min(eleves1.length, config.MAX_ITERATIONS_PAR_PAIRE);
  const maxCandidats2 = Math.min(eleves2.length, config.MAX_ITERATIONS_PAR_PAIRE);
  
  for (let i = 0; i < maxCandidats1; i++) {
    for (let j = 0; j < maxCandidats2; j++) {
      const e1 = eleves1[i];
      const e2 = eleves2[j];
      
      if (!respecteContraintes(e1, e2, students, structure, optionsNiveau, optionPools, dissocMap)) {
        continue;
      }
      
      // Passer la config à calculerSimilariteScoresV2
      const similarite = calculerSimilariteScoresV2(e1, e2, config);
      if (similarite < config.SEUIL_SIMILARITE) continue;
      
      const impact = calculerImpactSwapSymetrique(
        ecart1, ecart2, e1.SEXE, e2.SEXE, cibles, similarite
      );
      
      candidats.push({
        type: 'symetrique',
        eleves1: [e1],
        eleves2: [e2],
        classe1: ecart1.classe,
        classe2: ecart2.classe,
        ...impact,
        similariteScore: similarite
      });
    }
  }
  
  return candidats;
}

// =================================================
// 6. TROUVER CANDIDATS POUR SWAPS ASYMÉTRIQUES
// =================================================
function trouverCandidatsAsymetriques(ecartSource, ecartDest, direction, classesMap, students, structure, optionsNiveau, optionPools, dissocMap, cibles) {
  const candidats = [];
  const config = getPariteConfig(); // Une seule fois au début
  
  if (direction === "2vers1") {
    // Chercher 2 élèves dans source qui peuvent aller vers 1 élève dans dest
    const pairesSource = trouverMeilleuresPaires(
      classesMap[ecartSource.classe],
      ecartSource,
      cibles
    );
    
    const elevesDest = classesMap[ecartDest.classe].filter(e =>
      (e.mobilite || 'LIBRE') !== 'FIXE' &&
      (e.mobilite || 'LIBRE') !== 'SPEC'
    );
    
    for (const paire of pairesSource.slice(0, 10)) { // Top 10 paires
      for (const eDest of elevesDest.slice(0, 20)) { // Top 20 candidats
        // Vérifier contraintes pour les deux échanges
        const contraintes1 = respecteContraintes(paire.eleve1, eDest, students, structure, optionsNiveau, optionPools, dissocMap);
        const contraintes2 = respecteContraintes(paire.eleve2, eDest, students, structure, optionsNiveau, optionPools, dissocMap);
        
        if (!contraintes1 || !contraintes2) continue;
        
        // ... dans les boucles, passer config :
  const sim1 = calculerSimilariteScoresV2(paire.eleve1, eDest, config);
  const sim2 = calculerSimilariteScoresV2(paire.eleve2, eDest, config);
        const similariteMoyenne = (sim1 + sim2) / 2;
        
        if (similariteMoyenne < config.SEUIL_SIMILARITE) continue;
        
        const impact = calculerImpactSwapAsymetrique(
          ecartSource, ecartDest,
          [paire.eleve1.SEXE, paire.eleve2.SEXE],
          [eDest.SEXE],
          cibles, similariteMoyenne
        );
        
        candidats.push({
          type: 'asymetrique',
          direction: '2vers1',
          eleves1: [paire.eleve1, paire.eleve2],
          eleves2: [eDest],
          classe1: ecartSource.classe,
          classe2: ecartDest.classe,
          ...impact,
          similariteScore: similariteMoyenne
        });
      }
    }
  }
  
  // Direction "1vers2" : similaire mais inversé
  // (code similaire avec les rôles inversés)
  else if (direction === "1vers2") {
    // Chercher 1 élève dans source qui peut aller vers 2 élèves dans dest
    const elevesSource = classesMap[ecartSource.classe].filter(e =>
      (e.mobilite || 'LIBRE') !== 'FIXE' &&
      (e.mobilite || 'LIBRE') !== 'SPEC'
    );
    
    const pairesDest = trouverMeilleuresPaires(
      classesMap[ecartDest.classe],
      ecartDest,
      cibles
    );
    
    for (const eSource of elevesSource.slice(0, 20)) { // Top 20 candidats
      for (const paire of pairesDest.slice(0, 10)) { // Top 10 paires
        // Vérifier contraintes pour les deux échanges
        const contraintes1 = respecteContraintes(eSource, paire.eleve1, students, structure, optionsNiveau, optionPools, dissocMap);
        const contraintes2 = respecteContraintes(eSource, paire.eleve2, students, structure, optionsNiveau, optionPools, dissocMap);
        
        if (!contraintes1 || !contraintes2) continue;
        
        // Similarité moyenne
        const sim1 = calculerSimilariteScoresV2(eSource, paire.eleve1);
        const sim2 = calculerSimilariteScoresV2(eSource, paire.eleve2);
        const similariteMoyenne = (sim1 + sim2) / 2;
        
        if (similariteMoyenne < config.SEUIL_SIMILARITE) continue;
        
        const impact = calculerImpactSwapAsymetrique(
          ecartSource, ecartDest,
          [eSource.SEXE],
          [paire.eleve1.SEXE, paire.eleve2.SEXE],
          cibles, similariteMoyenne
        );
        
        candidats.push({
          type: 'asymetrique',
          direction: '1vers2',
          eleves1: [eSource],
          eleves2: [paire.eleve1, paire.eleve2],
          classe1: ecartSource.classe,
          classe2: ecartDest.classe,
          ...impact,
          similariteScore: similariteMoyenne
        });
      }
    }
  }
  
  return candidats;
}

// =================================================
// 7. CALCUL DE SIMILARITÉ V2 (AVEC CONFIG)
// =================================================
function calculerSimilariteScoresV2(eleve1, eleve2, config = null) {
  // Utiliser la config passée ou la récupérer une seule fois
  if (!config) {
    config = getPariteConfig();
  }

  // Niveaux
  const niv1 = {
    COM: eleve1.niveauCOM || 1,
    TRA: eleve1.niveauTRA || 1,
    PART: eleve1.niveauPART || 1
  };
  
  const niv2 = {
    COM: eleve2.niveauCOM || 1,
    TRA: eleve2.niveauTRA || 1,
    PART: eleve2.niveauPART || 1
  };
  
  // Différences
  const diffCOM = Math.abs(niv1.COM - niv2.COM);
  const diffTRA = Math.abs(niv1.TRA - niv2.TRA);
  const diffPART = Math.abs(niv1.PART - niv2.PART);
  const diffTotale = diffCOM + diffTRA + diffPART;
  
  // Moyennes
  const moy1 = (niv1.COM + niv1.TRA + niv1.PART) / 3;
  const moy2 = (niv2.COM + niv2.TRA + niv2.PART) / 3;
  const diffMoy = Math.abs(moy1 - moy2);
  
  // Calcul avec pénalités configurables
  let similarite = 1.0;
  
  // Pénalité pour différences de niveaux
  similarite -= diffTotale * config.PENALITE_NIVEAU;
  
  // Pénalité pour différence de moyenne
  similarite -= diffMoy * config.PENALITE_MOYENNE;
  
  // Pénalités pour statuts différents
  if (eleve1.estTeteDeClasse !== eleve2.estTeteDeClasse) {
    similarite -= config.PENALITE_TETE;
  }
  if (eleve1.estNiveau1 !== eleve2.estNiveau1) {
    similarite -= config.PENALITE_NIV1;
  }
  
  // Bonus si identiques
  if (diffTotale === 0) {
    similarite += config.BONUS_IDENTIQUE;
  }
  
  return Math.max(0, Math.min(1, similarite));
}

// =================================================
// 8. CALCUL D'IMPACT POUR SWAP SYMÉTRIQUE
// =================================================
function calculerImpactSwapSymetrique(ecart1, ecart2, sexe1, sexe2, cibles, similarite) {
  const config = getPariteConfig();
  
  // État actuel
  const avant = {
    ecartF1: Math.abs(ecart1.ecartF),
    ecartF2: Math.abs(ecart2.ecartF),
    ecartEff1: Math.abs(ecart1.ecartEffectif),
    ecartEff2: Math.abs(ecart2.ecartEffectif),
    ecartTotal: Math.abs(ecart1.ecartF) + Math.abs(ecart2.ecartF)
  };
  
  // Éviter division par zéro
  if (avant.ecartTotal < 0.001) {
    return { impactGlobal: 0, ameliorationParite: 0, ameliorationEffectifs: 0 };
  }
  
  // Simuler l'échange
  const deltaF1 = (sexe2 === 'F' ? 1 : 0) - (sexe1 === 'F' ? 1 : 0);
  const deltaF2 = (sexe1 === 'F' ? 1 : 0) - (sexe2 === 'F' ? 1 : 0);
  
  // État après
  const apres = {
    ecartF1: Math.abs(ecart1.ecartF + deltaF1),
    ecartF2: Math.abs(ecart2.ecartF + deltaF2),
    ecartEff1: avant.ecartEff1, // Inchangé pour swap 1↔1
    ecartEff2: avant.ecartEff2
  };
  
  // Améliorations
  const ameliorationParite = (avant.ecartF1 + avant.ecartF2 - apres.ecartF1 - apres.ecartF2) / avant.ecartTotal;
  const ameliorationEffectifs = 0; // Pas d'amélioration pour swap symétrique
  
  // Impact global pondéré
  const impactGlobal = (
    ameliorationParite * config.POIDS_PARITE +
    ameliorationEffectifs * config.POIDS_EFFECTIFS
  ) * similarite;
  
  return {
    impactGlobal,
    ameliorationParite,
    ameliorationEffectifs,
    deltaF1,
    deltaF2
  };
}

// =================================================
// 9. CALCUL D'IMPACT POUR SWAP ASYMÉTRIQUE
// =================================================
function calculerImpactSwapAsymetrique(ecartSource, ecartDest, sexesSource, sexesDest, cibles, similarite) {
  const config = getPariteConfig();
  
  // Compter les F échangées
  const nbFSource = sexesSource.filter(s => s === 'F').length;
  const nbFDest = sexesDest.filter(s => s === 'F').length;
  
  // Delta effectifs
  const deltaEffSource = -sexesSource.length + sexesDest.length;
  const deltaEffDest = -sexesDest.length + sexesSource.length;
  
  // Delta F
  const deltaFSource = nbFDest - nbFSource;
  const deltaFDest = nbFSource - nbFDest;
  
  // État avant
  const avant = {
    ecartFSource: Math.abs(ecartSource.ecartF),
    ecartFDest: Math.abs(ecartDest.ecartF),
    ecartEffSource: Math.abs(ecartSource.ecartEffectif),
    ecartEffDest: Math.abs(ecartDest.ecartEffectif),
    totalEcartF: Math.abs(ecartSource.ecartF) + Math.abs(ecartDest.ecartF),
    totalEcartEff: Math.abs(ecartSource.ecartEffectif) + Math.abs(ecartDest.ecartEffectif)
  };
  
  // État après
  const apres = {
    ecartFSource: Math.abs(ecartSource.ecartF + deltaFSource),
    ecartFDest: Math.abs(ecartDest.ecartF + deltaFDest),
    ecartEffSource: Math.abs(ecartSource.ecartEffectif + deltaEffSource),
    ecartEffDest: Math.abs(ecartDest.ecartEffectif + deltaEffDest)
  };
  
  // Améliorations (éviter division par zéro)
  const ameliorationParite = avant.totalEcartF > 0.001 ?
    (avant.ecartFSource + avant.ecartFDest - apres.ecartFSource - apres.ecartFDest) / avant.totalEcartF : 0;
    
  const ameliorationEffectifs = avant.totalEcartEff > 0.001 ?
    (avant.ecartEffSource + avant.ecartEffDest - apres.ecartEffSource - apres.ecartEffDest) / avant.totalEcartEff : 0;
  
  // Impact global
  const impactGlobal = (
    ameliorationParite * config.POIDS_PARITE +
    ameliorationEffectifs * config.POIDS_EFFECTIFS
  ) * similarite;
  
  return {
    impactGlobal,
    ameliorationParite,
    ameliorationEffectifs,
    deltaFSource,
    deltaFDest,
    deltaEffSource,
    deltaEffDest
  };
}

// =================================================
// 10. SÉLECTION DES MEILLEURS SWAPS SANS CONFLIT
// =================================================
function selectionnerMeilleursSwaps(tousLesSwaps, maxSwaps) {
  // Trier par impact décroissant
  tousLesSwaps.sort((a, b) => b.impactGlobal - a.impactGlobal);
  
  const swapsSelectionnes = [];
  const elevesUtilises = new Set();
  
  for (const swap of tousLesSwaps) {
    // Vérifier qu'aucun élève n'est déjà utilisé
    const idsSwap = [
      ...swap.eleves1.map(e => e.ID_ELEVE),
      ...swap.eleves2.map(e => e.ID_ELEVE)
    ];
    
    const conflit = idsSwap.some(id => elevesUtilises.has(id));
    
    if (!conflit) {
      swapsSelectionnes.push(swap);
      idsSwap.forEach(id => elevesUtilises.add(id));
      
      if (swapsSelectionnes.length >= maxSwaps) break;
    }
  }
  
  return swapsSelectionnes;
}

// =================================================
// 11. TROUVER LES MEILLEURES PAIRES (POUR ASYMÉTRIQUE)
// =================================================
function trouverMeilleuresPaires(eleves, ecartClasse, cibles) {
  const paires = [];
  const candidats = eleves.filter(e =>
    (e.mobilite || 'LIBRE') !== 'FIXE' &&
    (e.mobilite || 'LIBRE') !== 'SPEC'
  );
  
  // Pour chaque paire possible
  for (let i = 0; i < candidats.length - 1; i++) {
    for (let j = i + 1; j < candidats.length; j++) {
      const e1 = candidats[i];
      const e2 = candidats[j];
      
      // Calculer le bénéfice de déplacer cette paire
      const benefice = calculerBeneficePaire(e1, e2, ecartClasse, cibles);
      
      paires.push({
        eleve1: e1,
        eleve2: e2,
        benefice
      });
    }
  }
  
  // Trier par bénéfice décroissant
  paires.sort((a, b) => b.benefice - a.benefice);
  
  return paires;
}

function calculerBeneficePaire(e1, e2, ecartClasse, cibles) {
  // Bénéfice basé sur la réduction des écarts
  let benefice = 0;
  
  // Si la classe est trop grande, bénéfice de retirer 2 élèves
  if (ecartClasse.tropGrande) {
    benefice += 2; // Base pour réduire l'effectif
  }
  
  // Si on retire des F d'une classe qui en a trop
  const nbFRetires = (e1.SEXE === 'F' ? 1 : 0) + (e2.SEXE === 'F' ? 1 : 0);
  if (ecartClasse.besoinMoinsF && nbFRetires > 0) {
    benefice += nbFRetires * 1.5;
  }
  
  // Bonus si les deux ont des scores similaires (facilite l'échange)
  const similarite = calculerSimilariteScoresV2(e1, e2);
  benefice += similarite;
  
  return benefice;
}

// =================================================
// 12. EXÉCUTION DES SWAPS OPTIMISÉS
// =================================================
function executerSwapsOptimises(swapsOptimaux, niveau) {
  const journalSwaps = [];
  
  try {
    for (const swap of swapsOptimaux) {
      if (swap.type === 'symetrique') {
        // Swap classique 1↔1
        executerSwapsDansOnglets([{
          eleve1ID: swap.eleves1[0].ID_ELEVE,
          eleve1Nom: swap.eleves1[0].NOM,
          eleve2ID: swap.eleves2[0].ID_ELEVE,
          eleve2Nom: swap.eleves2[0].NOM,
          classe1: swap.classe1,
          classe2: swap.classe2,
          impact: swap.impactGlobal
        }]);
        
        journalSwaps.push({
          type: 'symetrique',
          eleve1ID: swap.eleves1[0].ID_ELEVE,
          eleve1Nom: swap.eleves1[0].NOM,
          eleve2ID: swap.eleves2[0].ID_ELEVE,
          eleve2Nom: swap.eleves2[0].NOM,
          classe1: swap.classe1,
          classe2: swap.classe2,
          impact: swap.impactGlobal
        });
        
      } else if (swap.type === 'asymetrique') {
        // Swap asymétrique : plusieurs mouvements
        const mouvements = [];
        
        // ========== SUITE DE executerSwapsOptimises ==========
        // Élèves de classe1 vers classe2
        swap.eleves1.forEach(e => {
          mouvements.push({
            eleveID: e.ID_ELEVE,
            eleveNom: e.NOM,
            classeOrigine: swap.classe1,
            classeDestination: swap.classe2
          });
        });
        
        // Élèves de classe2 vers classe1
        swap.eleves2.forEach(e => {
          mouvements.push({
            eleveID: e.ID_ELEVE,
            eleveNom: e.NOM,
            classeOrigine: swap.classe2,
            classeDestination: swap.classe1
          });
        });
        
        // Exécuter les mouvements asymétriques
        executerMouvementsAsymetriques(mouvements);
        
        // Enregistrer dans le journal
        journalSwaps.push({
          type: 'asymetrique',
          direction: swap.direction,
          eleves1: swap.eleves1.map(e => ({ id: e.ID_ELEVE, nom: e.NOM })),
          eleves2: swap.eleves2.map(e => ({ id: e.ID_ELEVE, nom: e.NOM })),
          classe1: swap.classe1,
          classe2: swap.classe2,
          impact: swap.impactGlobal
        });
      }
    }
    
    // Sauvegarder le journal
    if (journalSwaps.length > 0) {
      sauvegarderJournalSwaps(journalSwaps, niveau, "PARITE_MF_V2");
    }
    
    return {
      success: true,
      nbSwaps: journalSwaps.length,
      journal: journalSwaps
    };
    
  } catch (e) {
    Logger.log(`❌ Erreur exécution swaps: ${e.message}`);
    return {
      success: false,
      message: "Erreur exécution: " + e.message
    };
  }
}

// =================================================
// FONCTION UTILITAIRE POUR TROUVER UN ÉLÈVE
// =================================================
function findStudentRowInSheet_Robust(sheet, targetId) {
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;
  
  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const idColIndex = headers.indexOf("ID_ELEVE");
  if (idColIndex === -1) {
    Logger.log(`findStudentRowInSheet_Robust: Colonne ID_ELEVE non trouvée dans '${sheet.getName()}'.`);
    return null;
  }
  
  const targetIdStr = String(targetId).trim();
  
  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1;
    const rowData = data[i];
    if (rowData.length <= idColIndex) continue;
    
    const idSheetRaw = rowData[idColIndex];
    const idSheetStr = String(idSheetRaw).trim();
    
    if (idSheetStr === targetIdStr) {
      return { rowNum: rowNum, rowData: rowData };
    }
  }
  return null;
}

// =================================================
// 13. EXÉCUTION DES MOUVEMENTS ASYMÉTRIQUES
// =================================================
function executerMouvementsAsymetriques(mouvements) {
  Logger.log(`Exécution de ${mouvements.length} mouvements asymétriques...`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const updates = {}; // Par feuille
  
  // Préparer les updates par feuille
  mouvements.forEach(mouv => {
    const sheetOrigine = ss.getSheetByName(mouv.classeOrigine);
    const sheetDest = ss.getSheetByName(mouv.classeDestination);
    
    if (!sheetOrigine || !sheetDest) {
      throw new Error(`Feuille ${mouv.classeOrigine} ou ${mouv.classeDestination} non trouvée`);
    }
    
    // Trouver l'élève dans la feuille origine
    const locOrigine = findStudentRowInSheet_Robust(sheetOrigine, mouv.eleveID);
    if (!locOrigine) {
      throw new Error(`Élève ${mouv.eleveID} non trouvé dans ${mouv.classeOrigine}`);
    }
    
    // Supprimer de l'origine
    if (!updates[mouv.classeOrigine]) updates[mouv.classeOrigine] = { toDelete: [], toAdd: [] };
    updates[mouv.classeOrigine].toDelete.push(locOrigine.rowNum);
    
    // Ajouter à la destination
    if (!updates[mouv.classeDestination]) updates[mouv.classeDestination] = { toDelete: [], toAdd: [] };
    updates[mouv.classeDestination].toAdd.push(locOrigine.rowData);
  });
  
  // Appliquer les updates
  Object.entries(updates).forEach(([sheetName, actions]) => {
    const sheet = ss.getSheetByName(sheetName);
    
    // Supprimer les lignes (en ordre inverse pour ne pas décaler)
    actions.toDelete.sort((a, b) => b - a).forEach(rowNum => {
      sheet.deleteRow(rowNum);
    });
    
    // Ajouter les nouvelles lignes en gardant l'ordre
    if (actions.toAdd.length > 0) {
      // Insérer après la ligne d'en-tête
      const currentLastRow = sheet.getLastRow();
      actions.toAdd.forEach((rowData, index) => {
        // Insérer à la position 2 (après l'en-tête)
        sheet.insertRowAfter(1);
        sheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);
      });
      
      // Optionnel : trier par ID_ELEVE après insertion
      try {
        const range = sheet.getDataRange();
        range.sort(1); // Trier par colonne A (ID_ELEVE)
      } catch (e) {
        Logger.log("Tri automatique échoué : " + e.message);
      }
    }
  });
}

// =================================================
// 14. AFFICHAGES AMÉLIORÉS
// =================================================
function afficherCiblesIdeales(cibles) {
  Logger.log("\n=== CIBLES IDÉALES ===");
  Logger.log(`Effectif moyen: ${cibles.effectifMoyen.toFixed(1)} élèves (±${cibles.toleranceEffectif})`);
  Logger.log(`Ratio F global: ${(cibles.ratioFGlobal * 100).toFixed(1)}%`);
  Logger.log(`Répartition idéale par classe:`);
  Logger.log(`  - ${cibles.nbFIdeal.toFixed(1)} filles`);
  Logger.log(`  - ${cibles.nbMIdeal.toFixed(1)} garçons`);
  Logger.log(`  - Total: ${cibles.effectifMin} à ${cibles.effectifMax} élèves`);
}

function afficherSwapsProposés(swaps) {
  Logger.log("\n=== SWAPS PROPOSÉS ===");
  
  const symetriques = swaps.filter(s => s.type === 'symetrique');
  const asymetriques = swaps.filter(s => s.type === 'asymetrique');
  
  if (symetriques.length > 0) {
    Logger.log(`\n📊 SWAPS SYMÉTRIQUES (${symetriques.length}):`);
    symetriques.forEach((swap, i) => {
      const e1 = swap.eleves1[0];
      const e2 = swap.eleves2[0];
      Logger.log(`${i + 1}. ${e1.NOM} (${e1.SEXE}, ${swap.classe1}) ↔ ${e2.NOM} (${e2.SEXE}, ${swap.classe2})`);
      Logger.log(`   Impact: ${swap.impactGlobal.toFixed(3)} (parité: ${swap.ameliorationParite.toFixed(2)}, similarité: ${swap.similariteScore.toFixed(2)})`);
    });
  }
  
  if (asymetriques.length > 0) {
    Logger.log(`\n📊 SWAPS ASYMÉTRIQUES (${asymetriques.length}):`);
    asymetriques.forEach((swap, i) => {
      const nbEleves1 = swap.eleves1.length;
      const nbEleves2 = swap.eleves2.length;
      const sexes1 = swap.eleves1.map(e => e.SEXE).join('/');
      const sexes2 = swap.eleves2.map(e => e.SEXE).join('/');
      
      Logger.log(`${i + 1}. ${nbEleves1} élève(s) [${sexes1}] de ${swap.classe1} ↔ ${nbEleves2} élève(s) [${sexes2}] de ${swap.classe2}`);
      Logger.log(`   Impact: ${swap.impactGlobal.toFixed(3)} (parité: ${swap.ameliorationParite.toFixed(2)}, effectifs: ${swap.ameliorationEffectifs.toFixed(2)})`);
      
      // Détailler les élèves
      Logger.log(`   Élèves déplacés:`);
      swap.eleves1.forEach(e => Logger.log(`     - ${e.NOM} (${e.SEXE}, niv:${e.niveauCOM}/${e.niveauTRA}/${e.niveauPART})`));
      Logger.log(`   En échange de:`);
      swap.eleves2.forEach(e => Logger.log(`     - ${e.NOM} (${e.SEXE}, niv:${e.niveauCOM}/${e.niveauTRA}/${e.niveauPART})`));
    });
  }
}

function afficherResultatsOptimisation(statsInitiales, statsFinales, cibles) {
  Logger.log("\n📊 RÉSULTATS DE L'OPTIMISATION:");
  
  // Améliorations globales
  const reductionEcartTypeParite = (statsInitiales.global.ecartTypeParite - statsFinales.global.ecartTypeParite) * 100;
  const reductionEcartTypeEffectifs = statsInitiales.global.ecartTypeEffectifs - statsFinales.global.ecartTypeEffectifs;
  
  Logger.log("\nAMÉLIORATIONS GLOBALES:");
  Logger.log(`  Écart-type parité: ${(statsInitiales.global.ecartTypeParite * 100).toFixed(1)}% → ${(statsFinales.global.ecartTypeParite * 100).toFixed(1)}% (${reductionEcartTypeParite > 0 ? '-' : '+'}${Math.abs(reductionEcartTypeParite).toFixed(1)}%)`);
  Logger.log(`  Écart-type effectifs: ${statsInitiales.global.ecartTypeEffectifs.toFixed(2)} → ${statsFinales.global.ecartTypeEffectifs.toFixed(2)} (${reductionEcartTypeEffectifs > 0 ? '-' : '+'}${Math.abs(reductionEcartTypeEffectifs).toFixed(2)})`);
  
  // Classes les plus améliorées
  const ameliorations = [];
  Object.keys(statsInitiales.parClasse).forEach(classe => {
    const avant = statsInitiales.parClasse[classe];
    const apres = statsFinales.parClasse[classe];
    
    const ecartFAvant = Math.abs(avant.nbF - cibles.nbFIdeal);
    const ecartFApres = Math.abs(apres.nbF - cibles.nbFIdeal);
    const ameliorationF = ecartFAvant - ecartFApres;
    
    const ecartEffAvant = Math.abs(avant.total - cibles.effectifMoyen);
    const ecartEffApres = Math.abs(apres.total - cibles.effectifMoyen);
    const ameliorationEff = ecartEffAvant - ecartEffApres;
    
    if (ameliorationF !== 0 || ameliorationEff !== 0) {
      ameliorations.push({
        classe,
        avant,
        apres,
        ameliorationF,
        ameliorationEff,
        scoreAmelioration: ameliorationF * 2 + ameliorationEff
      });
    }
  });
  
  ameliorations.sort((a, b) => b.scoreAmelioration - a.scoreAmelioration);
  
  Logger.log("\nCLASSES LES PLUS AMÉLIORÉES:");
  ameliorations.slice(0, 5).forEach(a => {
    Logger.log(`\n${a.classe}:`);
    Logger.log(`  Avant: ${a.avant.total} élèves (${a.avant.nbF}F/${a.avant.nbM}M)`);
    Logger.log(`  Après: ${a.apres.total} élèves (${a.apres.nbF}F/${a.apres.nbM}M)`);
    Logger.log(`  Amélioration: ${a.ameliorationF > 0 ? '+' : ''}${a.ameliorationF.toFixed(1)}F, ${a.ameliorationEff > 0 ? '+' : ''}${a.ameliorationEff.toFixed(1)} effectif`);
  });
  
  // Statistiques finales
  const nbClassesAmeliorees = ameliorations.filter(a => a.scoreAmelioration > 0).length;
  const nbClassesDegradees = ameliorations.filter(a => a.scoreAmelioration < 0).length;
  
  Logger.log("\nBILAN:");
  Logger.log(`  Classes améliorées: ${nbClassesAmeliorees}`);
  Logger.log(`  Classes dégradées: ${nbClassesDegradees}`);
  Logger.log(`  Classes inchangées: ${Object.keys(statsInitiales.parClasse).length - ameliorations.length}`);
}

// =================================================
// 15. EXPORT DES STATS VERS UN ONGLET (OPTIONNEL)
// =================================================
function exporterStatsPariteVersOnglet(stats, nomOnglet = "STATS_PARITE") {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(nomOnglet);
    
    // Créer l'onglet s'il n'existe pas
    if (!sheet) {
      sheet = ss.insertSheet(nomOnglet);
    }
    
    // Nettoyer
    sheet.clear();
    
    // En-têtes
    const headers = ["Classe", "Effectif", "Filles", "Garçons", "%F", "Écart F idéal", "Écart Effectif"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.getRange(1, 1, 1, headers.length).setBackground("#E8F5E9");
    
    // Données par classe
    const data = [];
    const cibles = calculerCiblesIdeales(stats);
    
    Object.entries(stats.parClasse)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .forEach(([classe, s]) => {
        const ecartF = s.nbF - cibles.nbFIdeal;
        const ecartEff = s.total - cibles.effectifMoyen;
        
        data.push([
          classe,
          s.total,
          s.nbF,
          s.nbM,
          s.pourcentF.toFixed(1) + "%",
          ecartF > 0 ? "+" + ecartF.toFixed(1) : ecartF.toFixed(1),
          ecartEff > 0 ? "+" + ecartEff.toFixed(1) : ecartEff.toFixed(1)
        ]);
      });
    
    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      
      // Mise en forme conditionnelle
      const rangeEcartF = sheet.getRange(2, 6, data.length, 1);
      const rangeEcartEff = sheet.getRange(2, 7, data.length, 1);
      
      // Colorer les écarts importants
      const rules = [];
      
      // Écart F : rouge si > ±2
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(2)
        .setBackground("#FFCDD2")
        .setRanges([rangeEcartF])
        .build());
      
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(-2)
        .setBackground("#FFCDD2")
        .setRanges([rangeEcartF])
        .build());
      
      // Écart Effectif : orange si > ±2
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(2)
        .setBackground("#FFE0B2")
        .setRanges([rangeEcartEff])
        .build());
      
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(-2)
        .setBackground("#FFE0B2")
        .setRanges([rangeEcartEff])
        .build());
      
      sheet.setConditionalFormatRules(rules);
    }
    
    // Ajouter les statistiques globales
    const rowStats = data.length + 4;
    sheet.getRange(rowStats, 1).setValue("STATISTIQUES GLOBALES");
    sheet.getRange(rowStats, 1).setFontWeight("bold");
    
    const globalStats = [
      ["Total élèves:", stats.global.total],
      ["Ratio F global:", (stats.global.ratioF * 100).toFixed(1) + "%"],
      ["Effectif moyen:", stats.global.effectifMoyen.toFixed(1)],
      ["Écart-type parité:", (stats.global.ecartTypeParite * 100).toFixed(1) + "%"],
      ["Écart-type effectifs:", stats.global.ecartTypeEffectifs.toFixed(2)]
    ];
    
    sheet.getRange(rowStats + 1, 1, globalStats.length, 2).setValues(globalStats);
    
    // Auto-resize
    sheet.autoResizeColumns(1, headers.length);
    
    Logger.log(`✅ Stats exportées vers l'onglet ${nomOnglet}`);
    
  } catch (e) {
    Logger.log(`Erreur export stats: ${e.message}`);
  }
}

// =================================================
// 16. FONCTIONS RÉUTILISÉES (À COPIER DU MODULE V1)
// =================================================
// =================================================
// 2. ANALYSE COMPLÈTE (PARITÉ + EFFECTIFS)
// =================================================
function analyserPariteEtEffectifs(classesMap) {
  const stats = {
    global: { 
      total: 0, 
      nbM: 0, 
      nbF: 0, 
      ratioF: 0,
      effectifMoyen: 0,
      ecartTypeEffectifs: 0,
      ecartTypeParite: 0
    },
    parClasse: {},
    classeMinEffectif: null,
    classeMaxEffectif: null,
    classeMinF: null,
    classeMaxF: null
  };
  
  const effectifs = [];
  const ratiosF = [];
  
  // Analyser chaque classe
  Object.entries(classesMap).forEach(([classe, eleves]) => {
    const nbM = eleves.filter(e => e.SEXE === 'M').length;
    const nbF = eleves.filter(e => e.SEXE === 'F').length;
    const total = eleves.length;
    const ratioF = total > 0 ? nbF / total : 0;
    
    stats.parClasse[classe] = {
      total,
      nbM,
      nbF,
      ratioF,
      pourcentF: ratioF * 100,
      ecartEffectif: 0, // Sera calculé après
      ecartParite: 0    // Sera calculé après
    };
    
    stats.global.total += total;
    stats.global.nbM += nbM;
    stats.global.nbF += nbF;
    
    effectifs.push(total);
    ratiosF.push(ratioF);
    
    // Trouver les extrêmes
    if (!stats.classeMinEffectif || total < stats.parClasse[stats.classeMinEffectif].total) {
      stats.classeMinEffectif = classe;
    }
    if (!stats.classeMaxEffectif || total > stats.parClasse[stats.classeMaxEffectif].total) {
      stats.classeMaxEffectif = classe;
    }
    if (!stats.classeMinF || nbF < stats.parClasse[stats.classeMinF].nbF) {
      stats.classeMinF = classe;
    }
    if (!stats.classeMaxF || nbF > stats.parClasse[stats.classeMaxF].nbF) {
      stats.classeMaxF = classe;
    }
  });
  
  // Calculer les moyennes et ratios globaux
  const nbClasses = Object.keys(classesMap).length;
  stats.global.effectifMoyen = stats.global.total / nbClasses;
  stats.global.ratioF = stats.global.total > 0 ? stats.global.nbF / stats.global.total : 0;
  
  // Calculer les écarts-types
  const moyenneEffectif = stats.global.effectifMoyen;
  const varianceEffectif = effectifs.reduce((sum, eff) => sum + Math.pow(eff - moyenneEffectif, 2), 0) / nbClasses;
  stats.global.ecartTypeEffectifs = Math.sqrt(varianceEffectif);
  
  const moyenneRatioF = ratiosF.reduce((sum, r) => sum + r, 0) / nbClasses;
  const varianceRatioF = ratiosF.reduce((sum, r) => sum + Math.pow(r - moyenneRatioF, 2), 0) / nbClasses;
  stats.global.ecartTypeParite = Math.sqrt(varianceRatioF);
  
  // Calculer les écarts par classe
  Object.entries(stats.parClasse).forEach(([classe, classeStats]) => {
    classeStats.ecartEffectif = Math.abs(classeStats.total - moyenneEffectif);
    classeStats.ecartParite = Math.abs(classeStats.ratioF - stats.global.ratioF);
  });
  
  return stats;
}

// =================================================
// 3. CALCULER LES CIBLES IDÉALES
// =================================================
function calculerCiblesIdeales(stats) {
  const nbClasses = Object.keys(stats.parClasse).length;
  const effectifMoyen = stats.global.effectifMoyen;
  const ratioFGlobal = stats.global.ratioF;
  
  // Tolérance pour les effectifs (±1 ou ±2 élèves)
  const toleranceEffectif = 1.5;
  
  // Calculer le nombre idéal de F et M par classe
  const nbFIdeal = effectifMoyen * ratioFGlobal;
  const nbMIdeal = effectifMoyen * (1 - ratioFGlobal);
  
  return {
    effectifMoyen,
    effectifMin: Math.floor(effectifMoyen - toleranceEffectif),
    effectifMax: Math.ceil(effectifMoyen + toleranceEffectif),
    ratioFGlobal,
    nbFIdeal,
    nbMIdeal,
    toleranceEffectif,
    toleranceParite: 0.05 // 5% de tolérance sur le ratio F
  };
}

// =================================================
// 4. IDENTIFIER LES ÉCARTS AUX CIBLES
// =================================================
function identifierEcartsCibles(stats, cibles) {
  const ecarts = [];
  
  Object.entries(stats.parClasse).forEach(([classe, classeStats]) => {
    const ecartEffectif = classeStats.total - cibles.effectifMoyen;
    const ecartF = classeStats.nbF - cibles.nbFIdeal;
    const ecartM = classeStats.nbM - cibles.nbMIdeal;
    
    // Score d'écart global (plus c'est élevé, plus la classe s'écarte de l'idéal)
    const scoreEcart = Math.abs(ecartEffectif) + Math.abs(ecartF) * 2; // Pondération plus forte sur la parité
    
    ecarts.push({
      classe,
      stats: classeStats,
      ecartEffectif,
      ecartF,
      ecartM,
      scoreEcart,
      besoinPlusF: ecartF < -0.5,
      besoinMoinsF: ecartF > 0.5,
      tropGrande: ecartEffectif > cibles.toleranceEffectif,
      tropPetite: ecartEffectif < -cibles.toleranceEffectif
    });
  });
  
  // Trier par score d'écart décroissant
  ecarts.sort((a, b) => b.scoreEcart - a.scoreEcart);
  
  return ecarts;
}
// =================================================
// 7. DÉTERMINER LES SCÉNARIOS D'ÉCHANGE
// =================================================
function determinerScenariosEchange(ecart1, ecart2) {
  const scenarios = [];
  
  // Analyser ce dont chaque classe a besoin
  const besoins1 = {
    plusF: ecart1.besoinPlusF,
    moinsF: ecart1.besoinMoinsF,
    plusEleves: ecart1.tropPetite,
    moinsEleves: ecart1.tropGrande
  };
  
  const besoins2 = {
    plusF: ecart2.besoinPlusF,
    moinsF: ecart2.besoinMoinsF,
    plusEleves: ecart2.tropPetite,
    moinsEleves: ecart2.tropGrande
  };
  
  // Scénario 1: Échanger pour améliorer la parité
  if (besoins1.plusF && besoins2.moinsF) {
    scenarios.push({ sexe1: 'M', sexe2: 'F', priorite: 1 });
  }
  if (besoins1.moinsF && besoins2.plusF) {
    scenarios.push({ sexe1: 'F', sexe2: 'M', priorite: 1 });
  }
  
  // Scénario 2: Échanger même sexe pour équilibrer les effectifs
  if ((besoins1.plusEleves && besoins2.moinsEleves) || 
      (besoins1.moinsEleves && besoins2.plusEleves)) {
    scenarios.push({ sexe1: 'M', sexe2: 'M', priorite: 2 });
    scenarios.push({ sexe1: 'F', sexe2: 'F', priorite: 2 });
  }
  
  // Scénario 3: Tout échange qui rapprocherait les deux classes de leurs cibles
  if (scenarios.length === 0) {
    // Essayer tous les échanges possibles
    scenarios.push({ sexe1: 'M', sexe2: 'F', priorite: 3 });
    scenarios.push({ sexe1: 'F', sexe2: 'M', priorite: 3 });
    scenarios.push({ sexe1: 'M', sexe2: 'M', priorite: 3 });
    scenarios.push({ sexe1: 'F', sexe2: 'F', priorite: 3 });
  }
  
  // Trier par priorité
  scenarios.sort((a, b) => a.priorite - b.priorite);
  
  return scenarios;
}

// =================================================
// 10. AFFICHAGE DES STATISTIQUES
// =================================================
function afficherStatsCompletes(stats) {
  Logger.log("\n📊 STATISTIQUES GLOBALES:");
  Logger.log(`Total: ${stats.global.total} élèves (${stats.global.nbF}F/${stats.global.nbM}M)`);
  Logger.log(`Ratio F global: ${(stats.global.ratioF * 100).toFixed(1)}%`);
  Logger.log(`Effectif moyen: ${stats.global.effectifMoyen.toFixed(1)} élèves/classe`);
  Logger.log(`Écart-type effectifs: ${stats.global.ecartTypeEffectifs.toFixed(2)}`);
  Logger.log(`Écart-type parité: ${(stats.global.ecartTypeParite * 100).toFixed(1)}%`);
  
  Logger.log("\n📋 DÉTAIL PAR CLASSE:");
  const classes = Object.entries(stats.parClasse)
    .sort((a, b) => b[1].ecartEffectif - a[1].ecartEffectif);
  
  classes.forEach(([classe, s]) => {
    const indicateurs = [];
    if (s.ecartEffectif > 2) indicateurs.push("⚠️EFFECTIF");
    if (s.ecartParite > 0.1) indicateurs.push("⚠️PARITÉ");
    
    Logger.log(`${classe}: ${s.total} élèves (${s.nbF}F/${s.nbM}M = ${s.pourcentF.toFixed(0)}%F) ${indicateurs.join(' ')}`);
  });
  
  Logger.log("\n🔍 CLASSES EXTRÊMES:");
  if (stats.classeMaxEffectif) {
    const max = stats.parClasse[stats.classeMaxEffectif];
    Logger.log(`Plus grande: ${stats.classeMaxEffectif} (${max.total} élèves)`);
  }
  if (stats.classeMinEffectif) {
    const min = stats.parClasse[stats.classeMinEffectif];
    Logger.log(`Plus petite: ${stats.classeMinEffectif} (${min.total} élèves)`);
  }
  if (stats.classeMaxF) {
    const maxF = stats.parClasse[stats.classeMaxF];
    Logger.log(`Plus de F: ${stats.classeMaxF} (${maxF.nbF}F/${maxF.nbM}M)`);
  }
  if (stats.classeMinF) {
    const minF = stats.parClasse[stats.classeMinF];
    Logger.log(`Moins de F: ${stats.classeMinF} (${minF.nbF}F/${minF.nbM}M)`);
  }
}

function afficherEcartsCibles(ecarts) {
  Logger.log("\n🎯 ÉCARTS AUX CIBLES:");
  const plusEcartes = ecarts.slice(0, 5); // Top 5 des plus écartés
  
  plusEcartes.forEach(ecart => {
    Logger.log(`\n${ecart.classe}:`);
    Logger.log(`  Effectif: ${ecart.stats.total} (écart: ${ecart.ecartEffectif > 0 ? '+' : ''}${ecart.ecartEffectif.toFixed(1)})`);
    Logger.log(`  Filles: ${ecart.stats.nbF} (écart: ${ecart.ecartF > 0 ? '+' : ''}${ecart.ecartF.toFixed(1)})`);
    Logger.log(`  Garçons: ${ecart.stats.nbM} (écart: ${ecart.ecartM > 0 ? '+' : ''}${ecart.ecartM.toFixed(1)})`);
    Logger.log(`  Score d'écart: ${ecart.scoreEcart.toFixed(2)}`);
  });
}

function afficherComparaisonComplete(statsInitiales, statsFinales, cibles) {
  Logger.log("\n📊 COMPARAISON AVANT/APRÈS:");
  
  // Statistiques globales
  Logger.log("\nGLOBAL:");
  Logger.log(`  Écart-type parité: ${(statsInitiales.global.ecartTypeParite * 100).toFixed(1)}% → ${(statsFinales.global.ecartTypeParite * 100).toFixed(1)}%`);
  Logger.log(`  Écart-type effectifs: ${statsInitiales.global.ecartTypeEffectifs.toFixed(2)} → ${statsFinales.global.ecartTypeEffectifs.toFixed(2)}`);
  
  // Changements par classe
  Logger.log("\nCHANGEMENTS PAR CLASSE:");
  Object.keys(statsInitiales.parClasse).forEach(classe => {
    const avant = statsInitiales.parClasse[classe];
    const apres = statsFinales.parClasse[classe];
    
    if (avant.nbF !== apres.nbF || avant.nbM !== apres.nbM) {
      const ecartFAvant = Math.abs(avant.nbF - cibles.nbFIdeal);
      const ecartFApres = Math.abs(apres.nbF - cibles.nbFIdeal);
      const amelioration = ecartFAvant > ecartFApres;
      
      Logger.log(`\n${amelioration ? '✅' : '❌'} ${classe}:`);
      Logger.log(`  Avant: ${avant.nbF}F/${avant.nbM}M (écart F: ${ecartFAvant.toFixed(1)})`);
      Logger.log(`  Après: ${apres.nbF}F/${apres.nbM}M (écart F: ${ecartFApres.toFixed(1)})`);
    }
  });
  
  // Résumé des améliorations
  Logger.log("\n📈 RÉSUMÉ DES AMÉLIORATIONS:");
  const nbClassesAmeliorees = Object.keys(statsInitiales.parClasse).filter(classe => {
    const avant = statsInitiales.parClasse[classe];
    const apres = statsFinales.parClasse[classe];
    const ecartAvant = Math.abs(avant.nbF - cibles.nbFIdeal);
    const ecartApres = Math.abs(apres.nbF - cibles.nbFIdeal);
    return ecartApres < ecartAvant;
  }).length;
  
  Logger.log(`  Classes améliorées: ${nbClassesAmeliorees}/${Object.keys(statsInitiales.parClasse).length}`);
  Logger.log(`  Réduction écart-type parité: ${((statsInitiales.global.ecartTypeParite - statsFinales.global.ecartTypeParite) * 100).toFixed(1)}%`);
}

// =================================================
// 11. FONCTION UTILITAIRE POUR RECHARGER LES DONNÉES
// =================================================
function rechargerClassesMap() {
  const config = getConfig();
  const elevesResult = chargerElevesEtClasses(config, "MOBILITE");
  
  if (!elevesResult.success) {
    throw new Error("Erreur rechargement classes");
  }
  
  let { students, classesMap } = elevesResult;
  if (!classesMap && students) {
    classesMap = {};
    students.forEach(eleve => {
      if (eleve && eleve.CLASSE) {
        if (!classesMap[eleve.CLASSE]) classesMap[eleve.CLASSE] = [];
        classesMap[eleve.CLASSE].push(eleve);
      }
    });
  }
  
  classifierEleves(students, ["com1", "tra4", "part4"]);
  
  return classesMap;
}

// =================================================
// 12. FONCTION D'ANALYSE SANS MODIFICATION
// =================================================
function analyserPariteSansModifier() {
  Logger.log("=== ANALYSE PARITÉ ET EFFECTIFS (SANS MODIFICATION) ===");
  
  const config = getConfig();
  const elevesResult = chargerElevesEtClasses(config, "MOBILITE");
  
  if (!elevesResult.success) {
    Logger.log("❌ Erreur chargement");
    return;
  }
  
  let { students, classesMap } = elevesResult;
  if (!classesMap && students) {
    classesMap = {};
    students.forEach(eleve => {
      if (eleve && eleve.CLASSE) {
        if (!classesMap[eleve.CLASSE]) classesMap[eleve.CLASSE] = [];
        classesMap[eleve.CLASSE].push(eleve);
      }
    });
  }
  
  classifierEleves(students, ["com1", "tra4", "part4"]);
  
  // Analyser la situation
  const stats = analyserPariteEtEffectifs(classesMap);
  afficherStatsCompletes(stats);
  
  // Calculer et afficher les cibles
  const cibles = calculerCiblesIdeales(stats);
  Logger.log("\n=== CIBLES IDÉALES ===");
  Logger.log(`Pour équilibrer parfaitement avec ${(stats.global.ratioF * 100).toFixed(0)}% de filles:`);
  Logger.log(`  - Effectif par classe: ${cibles.effectifMoyen.toFixed(1)} (±${cibles.toleranceEffectif})`);
  Logger.log(`  - Filles par classe: ${cibles.nbFIdeal.toFixed(1)}`);
  Logger.log(`  - Garçons par classe: ${cibles.nbMIdeal.toFixed(1)}`);
  
  // Identifier et afficher les écarts
  const ecarts = identifierEcartsCibles(stats, cibles);
  afficherEcartsCibles(ecarts);
  
  // Proposer des solutions
  Logger.log("\n=== RECOMMANDATIONS ===");
  if (ecarts[0].scoreEcart > 5) {
    Logger.log("⚠️ Des déséquilibres importants ont été détectés.");
    Logger.log("Pour optimiser la répartition, lancez: optimiserPariteMF()");
    
    // Estimer le nombre de swaps nécessaires
    const nbSwapsEstime = Math.min(
      Math.ceil(ecarts.filter(e => e.scoreEcart > 2).length / 2),
      10
    );
    Logger.log(`Nombre de swaps estimé: ${nbSwapsEstime}`);
  } else {
    Logger.log("✅ La répartition est déjà relativement équilibrée.");
    Logger.log("Des ajustements mineurs pourraient encore améliorer la situation.");
  }
  
  return { stats, cibles, ecarts };
}

// =================================================
// 13. FONCTION TEST POUR UNE CLASSE SPÉCIFIQUE
// =================================================
function analyserClasseSpecifique(nomClasse) {
  Logger.log(`=== ANALYSE CLASSE ${nomClasse} ===`);
  
  const config = getConfig();
  const elevesResult = chargerElevesEtClasses(config, "MOBILITE");
  
  if (!elevesResult.success) {
    Logger.log("❌ Erreur chargement");
    return;
  }
  
  let { students, classesMap } = elevesResult;
  if (!classesMap && students) {
    classesMap = {};
    students.forEach(eleve => {
      if (eleve && eleve.CLASSE) {
        if (!classesMap[eleve.CLASSE]) classesMap[eleve.CLASSE] = [];
        classesMap[eleve.CLASSE].push(eleve);
      }
    });
  }
  
  const eleves = classesMap[nomClasse];
  if (!eleves) {
    Logger.log(`❌ Classe ${nomClasse} non trouvée`);
    return;
  }
  
  classifierEleves(eleves, ["com1", "tra4", "part4"]);
  
  // Analyser la classe
  const nbF = eleves.filter(e => e.SEXE === 'F').length;
  const nbM = eleves.filter(e => e.SEXE === 'M').length;
  const total = eleves.length;
  
  Logger.log(`\nEffectif: ${total} élèves`);
  Logger.log(`Parité: ${nbF}F/${nbM}M (${(nbF/total*100).toFixed(0)}%F)`);
  
  // Analyser les élèves échangeables
  const echangeablesF = eleves.filter(e => 
    e.SEXE === 'F' && 
    (e.mobilite || 'LIBRE') !== 'FIXE' && 
    (e.mobilite || 'LIBRE') !== 'SPEC'
  );
  const echangeablesM = eleves.filter(e => 
    e.SEXE === 'M' && 
    (e.mobilite || 'LIBRE') !== 'FIXE' && 
    (e.mobilite || 'LIBRE') !== 'SPEC'
  );
  
  Logger.log(`\nÉlèves échangeables:`);
  Logger.log(`  Filles: ${echangeablesF.length}/${nbF}`);
  Logger.log(`  Garçons: ${echangeablesM.length}/${nbM}`);
  
  // Détailler les contraintes
  const fixes = eleves.filter(e => e.mobilite === 'FIXE').length;
  const specs = eleves.filter(e => e.mobilite === 'SPEC').length;
  const permuts = eleves.filter(e => e.mobilite === 'PERMUT').length;
  const condis = eleves.filter(e => e.mobilite === 'CONDI').length;
  
  Logger.log(`\nContraintes mobilité:`);
  Logger.log(`  FIXE: ${fixes}`);
  Logger.log(`  SPEC: ${specs}`);
  Logger.log(`  PERMUT: ${permuts}`);
  Logger.log(`  CONDI: ${condis}`);
  Logger.log(`  LIBRE: ${total - fixes - specs - permuts - condis}`);
  
  // Analyser les scores
  const tetes = eleves.filter(e => e.estTeteDeClasse).length;
  const niveau1 = eleves.filter(e => e.estNiveau1).length;
  
  Logger.log(`\nNiveaux:`);
  Logger.log(`  Têtes de classe: ${tetes}`);
  Logger.log(`  Niveau 1: ${niveau1}`);
  
  // Calculer l'écart à la cible
  const stats = analyserPariteEtEffectifs(classesMap);
  const cibles = calculerCiblesIdeales(stats);
  const ecartF = nbF - cibles.nbFIdeal;
  const ecartEffectif = total - cibles.effectifMoyen;
  
  Logger.log(`\nÉcarts aux cibles:`);
  Logger.log(`  Effectif: ${ecartEffectif > 0 ? '+' : ''}${ecartEffectif.toFixed(1)} (cible: ${cibles.effectifMoyen.toFixed(1)})`);
  Logger.log(`  Filles: ${ecartF > 0 ? '+' : ''}${ecartF.toFixed(1)} (cible: ${cibles.nbFIdeal.toFixed(1)})`);
  
  if (Math.abs(ecartF) > 2) {
    Logger.log(`\n⚠️ Cette classe a un déséquilibre important de parité.`);
    Logger.log(`Recommandation: ${ecartF > 0 ? 'échanger des F contre des M' : 'échanger des M contre des F'}`);
  }
  
  return {
    classe: nomClasse,
    effectif: total,
    nbF, nbM,
    echangeables: { F: echangeablesF.length, M: echangeablesM.length },
    ecarts: { effectif: ecartEffectif, filles: ecartF }
  };
}

// =================================================
// 17. MENU AMÉLIORÉ V2
// =================================================
function ajouterMenuPariteV3() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 Optimisations V14')
    .addItem('⚖️ Optimiser Parité M/F (V2)', 'optimiserPariteMF_V2')
    .addItem('📊 Analyser Parité (sans modifier)', 'analyserPariteSansModifier')
    .addItem('📈 Exporter Stats vers onglet', 'menuExporterStats')
    .addItem('🔍 Analyser une classe...', 'menuAnalyserClasse')
    .addSeparator()
    .addItem('⚙️ Configurer paramètres parité...', 'menuConfigurerParite')
    .addSeparator()
    .addItem('🔄 Optimisation Standard V14', 'menuOptimizationDialog')
    .addToUi();
}

function menuExporterStats() {
  const classesMap = rechargerClassesMap();
  const stats = analyserPariteEtEffectifs(classesMap);
  exporterStatsPariteVersOnglet(stats);
  SpreadsheetApp.getUi().alert('Export terminé', 'Les statistiques ont été exportées vers l\'onglet STATS_PARITE', SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuConfigurerParite() {
  const ui = SpreadsheetApp.getUi();
  const config = getPariteConfig();
  
  const html = `
    <div style="padding: 10px;">
      <h3>Configuration Parité</h3>
      <p>Seuil similarité: ${config.SEUIL_SIMILARITE}</p>
      <p>Tolérance effectif: ±${config.TOLERANCE_EFFECTIF}</p>
      <p>Swaps asymétriques: ${config.AUTORISER_SWAPS_ASYMETRIQUES ? 'Oui' : 'Non'}</p>
      <hr>
      <p><small>Pour modifier, éditez PARITE_CONFIG dans le code ou Config.gs</small></p>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(300)
    .setHeight(200);
  
  ui.showModalDialog(htmlOutput, 'Configuration Parité');
}