/**
 * ==================================================================
 *     FICHIER : Nirvana_V2_Amelioree_Integre.gs
 *     Version améliorée avec intégration MultiSwap + MultiSwap Parité
 *     RESPECTANT STRICTEMENT L'ONGLET _STRUCTURE
 * ==================================================================
 * Version: 2.2
 * Date: 11 Juin 2025
 * 
 * Description:
 *   Version améliorée de Nirvana V2 qui vise une répartition équilibrée
 *   des scores 1-2-3-4 dans chaque classe, tout en respectant STRICTEMENT
 *   les contraintes de l'onglet _STRUCTURE (options, dissociations, etc.)
 *   
 *   NOUVEAUTÉ: Intégration automatique de MultiSwap et MultiSwap Parité
 *   après l'optimisation principale pour affiner l'équilibrage.
 * ==================================================================
 */

'use strict';

// ==================================================================
// SECTION 1: ORCHESTRATEUR PRINCIPAL
// ==================================================================

/**
 * Point d'entrée UI amélioré
 */
function lancerOptimisationNirvanaV2_UI(criteresUI_optionnel) {
  const ui = SpreadsheetApp.getUi();
  const heureDebutTotal = new Date();
  let lock = null;

  try {
    Logger.log(`\n##########################################################`);
    Logger.log(` LANCEMENT NIRVANA V2 AMÉLIORÉ + MULTISWAP - ${heureDebutTotal.toLocaleString('fr-FR')}`);
    Logger.log(` Objectif: Équilibrage des scores 1-2-3-4 + cycles optimisés`);
    Logger.log(` Contraintes: Respect strict de _STRUCTURE`);
    Logger.log(`##########################################################`);

    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      Logger.log("V2 Amélioré: Verrouillage impossible.");
      ui.alert("Optimisation en cours", "Un autre processus est déjà actif.", ui.ButtonSet.OK);
      return { success: false, errorCode: "LOCKED" };
    }

    SpreadsheetApp.getActiveSpreadsheet().toast("Nirvana V2 Amélioré + MultiSwap: Analyse en cours...", "Statut", 10);

    // Appel du moteur amélioré avec MultiSwap intégré
    const resultat = V2_Ameliore_OptimisationEngine(criteresUI_optionnel);

    // Message final
    if (resultat.success) {
      const message = `✅ Optimisation Nirvana V2 + MultiSwap RÉUSSIE !\n\n` +
                     `Swaps principaux: ${resultat.nbSwapsAppliques || 0}\n` +
                     `Cycles généraux: ${resultat.cyclesGeneraux || 0}\n` +
                     `Cycles parité: ${resultat.cyclesParite || 0}\n` +
                     `Score d'équilibre final: ${resultat.scoreEquilibre?.toFixed(2) || 'N/A'}/100\n` +
                     `Durée: ${(resultat.tempsMs / 1000).toFixed(1)} secondes.\n\n` +
                     `Consultez le bilan détaillé dans '${resultat.nomFeuilleBilan}'`;
      
      ui.alert('Optimisation Terminée', message, ui.ButtonSet.OK);
    } else {
      ui.alert('Erreur', `❌ Échec: ${resultat.error}`, ui.ButtonSet.OK);
    }

    return resultat;

  } catch (e) {
    Logger.log(`ERREUR FATALE: ${e.message}\n${e.stack}`);
    ui.alert("Erreur Critique", e.message, ui.ButtonSet.OK);
    return { success: false, error: e.message };
  } finally {
    if (lock) lock.releaseLock();
    Logger.log(`FIN WRAPPER | Durée: ${(new Date() - heureDebutTotal) / 1000}s`);
  }
}

/**
 * Moteur principal amélioré avec intégration MultiSwap
 */
function V2_Ameliore_OptimisationEngine(criteresUI) {
  const heureDebut = new Date();
  const config = getConfig();
  
  // Configuration spécifique V2 Amélioré
  config.V2_POIDS_EQUILIBRE = {
    COM: 0.3,      // Poids pour l'équilibre des scores en COM
    TRA: 0.25,      // Poids pour l'équilibre des scores en TRA  
    PART: 0.25,     // Poids pour l'équilibre des scores en PART
    PARITE: 0.20    // Poids pour la parité F/M
  };
  
  config.V2_TOLERANCE_EQUILIBRE = 0.02; // Tolérance de 2% par rapport à la cible
  config.V2_MAX_ITERATIONS_GLOBAL = 1000;
  config.V2_MIN_AMELIORATION = 0.0001; // Amélioration minimale pour continuer

  try {
    Logger.log("V2 Amélioré: Préparation des données...");
    const dataContext = V2_Ameliore_PreparerDonnees(config);
    
    Logger.log("V2 Amélioré: Calcul de l'état initial...");
    const etatInitial = V2_Ameliore_CalculerEtatGlobal(dataContext, config);
    V2_Ameliore_LogEtat("INITIAL", etatInitial, config);
    
    Logger.log("V2 Amélioré: Optimisation principale en cours...");
    const journalSwaps = V2_Ameliore_OptimiserGlobal(dataContext, config);
    
    Logger.log("V2 Amélioré: Calcul de l'état après optimisation principale...");
    const etatApresSwaps = V2_Ameliore_CalculerEtatGlobal(dataContext, config);
    V2_Ameliore_LogEtat("APRÈS SWAPS PRINCIPAUX", etatApresSwaps, config);
    
    // ================================================================
    // NOUVEAUTÉ: INTÉGRATION MULTISWAP
    // ================================================================
    
    let cyclesGeneraux = 0;
    let cyclesParite = 0;
    let swapsMultiGeneraux = [];
    let swapsMultiParite = [];
    
    // 1. MultiSwap général (équilibrage critères avec cycles de 3)
    Logger.log("\n" + "=".repeat(60));
    Logger.log("V2 Amélioré: Lancement MultiSwap équilibrage général...");
    Logger.log("=".repeat(60));
    
    try {
      if (typeof V2_Ameliore_MultiSwap_AvecRetourSwaps === 'function') {
        const resultMulti = V2_Ameliore_MultiSwap_AvecRetourSwaps(dataContext, config);
        cyclesGeneraux = resultMulti.nbCycles;
        swapsMultiGeneraux = resultMulti.swapsDetailles || [];
        Logger.log(`✅ MultiSwap général terminé: ${cyclesGeneraux} cycles (${swapsMultiGeneraux.length} échanges)`);
      } else {
        Logger.log("⚠️ Fonction V2_Ameliore_MultiSwap_AvecRetourSwaps non disponible");
      }
    } catch (e) {
      Logger.log(`❌ Erreur MultiSwap général: ${e.message}`);
    }
    
    // 2. MultiSwap parité (correction genre avec cycles de 4)  
    Logger.log("\n" + "=".repeat(60));
    Logger.log("V2 Amélioré: Lancement MultiSwap correction parité...");
    Logger.log("=".repeat(60));
    
    try {
      if (typeof V2_Ameliore_MultiSwap4_Parite_AvecRetourSwaps === 'function') {
        const resultParite = V2_Ameliore_MultiSwap4_Parite_AvecRetourSwaps(dataContext, config);
        cyclesParite = resultParite.nbCycles;
        swapsMultiParite = resultParite.swapsDetailles || [];
        Logger.log(`✅ MultiSwap parité terminé: ${cyclesParite} cycles (${swapsMultiParite.length} échanges)`);
      } else {
        Logger.log("⚠️ Fonction V2_Ameliore_MultiSwap4_Parite_AvecRetourSwaps non disponible");
      }
    } catch (e) {
      Logger.log(`❌ Erreur MultiSwap parité: ${e.message}`);
    }
    
    Logger.log("\n" + "=".repeat(60));
    Logger.log(`BILAN MULTISWAP: ${cyclesGeneraux} cycles généraux + ${cyclesParite} cycles parité`);
    Logger.log(`TOTAL ÉCHANGES: ${swapsMultiGeneraux.length + swapsMultiParite.length} échanges MultiSwap`);
    Logger.log("=".repeat(60));
    
    // ================================================================
    // CONCATÉNATION DES SWAPS MULTISWAP AU JOURNAL PRINCIPAL
    // ================================================================
    
    // Ajouter les swaps MultiSwap au journal principal pour application
    if (swapsMultiGeneraux.length > 0) {
      Logger.log(`Ajout de ${swapsMultiGeneraux.length} échanges MultiSwap général au journal`);
      journalSwaps.push(...swapsMultiGeneraux);
    }
    
    if (swapsMultiParite.length > 0) {
      Logger.log(`Ajout de ${swapsMultiParite.length} échanges MultiSwap parité au journal`);
      journalSwaps.push(...swapsMultiParite);
    }
    
    // ================================================================
    // FIN INTÉGRATION MULTISWAP
    // ================================================================
    
/*=================================================================
  CORRECTION PARITÉ FINALE  (Patch_Nirvana_V2_Parite_PROD_READY.gs)
=================================================================*/
Logger.log("\n" + "=".repeat(60));
Logger.log("V2 Amélioré : Lancement Correction Parité Finale Débloquée …");
Logger.log("=".repeat(60));

let opsCorrectionPariteFinale = [];
try {
  if (typeof integrerCorrectionPariteDansNirvanaV2 === 'function') {

    // 1) Appel – on passe le dataContext courant et la config globale
    opsCorrectionPariteFinale = integrerCorrectionPariteDansNirvanaV2(
      dataContext,
      config                 // DEBUG_MODE_PARITE hérité ici
    );

    Logger.log(
      `✅ Correction Parité Finale : ${opsCorrectionPariteFinale.length} opération(s) générée(s).`
    );

    // 2) On ajoute ces opérations au journal global
    if (opsCorrectionPariteFinale.length) {
      journalSwaps.push(...opsCorrectionPariteFinale);
      Logger.log(`Ajout de ${opsCorrectionPariteFinale.length} opération(s) de parité au journalSwaps.`);
    }

  } else {
    Logger.log("⚠️  Fonction integrerCorrectionPariteDansNirvanaV2 absente – script parité non chargé ?");
  }
} catch (err) {
  Logger.log(`❌ Erreur Correction Parité Finale : ${err.message}`);
  Logger.log(err.stack);
}
/*==========================  FIN PARITÉ  =========================*/



    Logger.log("V2 Amélioré: Calcul de l'état final...");
    const etatFinal = V2_Ameliore_CalculerEtatGlobal(dataContext, config);
    V2_Ameliore_LogEtat("FINAL (après MultiSwap)", etatFinal, config);
    
    Logger.log("V2 Amélioré: Application de TOUS les swaps (principaux + MultiSwap)...");
    const nbSwapsAppliques = V2_Ameliore_AppliquerSwaps(journalSwaps, dataContext, config);
    
    Logger.log("V2 Amélioré: Génération du bilan...");
    const nomFeuilleBilan = V2_Ameliore_GenererBilan(
      dataContext, 
      etatInitial, 
      etatFinal, 
      journalSwaps, 
      config,
      { cyclesGeneraux, cyclesParite } // Infos MultiSwap
    );
    
    // CORRECTION: Calcul des statistiques finales APRÈS application de tous les swaps
    Logger.log("V2 Amélioré: Calcul des statistiques finales (APRÈS tous les swaps)...");
    try {
      V2_calculerStatistiquesFinales();
      Logger.log("✅ Statistiques style Parité calculées pour Nirvana V2.");
    } catch (e) {
      Logger.log("❌ Erreur calcul stats Nirvana V2: " + e.message);
    }
    
    const tempsMs = new Date() - heureDebut;
    Logger.log(`\n${"=".repeat(60)}`);
    Logger.log(`V2 Amélioré + MultiSwap terminé en ${(tempsMs / 1000).toFixed(1)}s`);
    Logger.log(`Score: ${etatInitial.scoreGlobal.toFixed(2)} → ${etatFinal.scoreGlobal.toFixed(2)} (+${(etatFinal.scoreGlobal - etatInitial.scoreGlobal).toFixed(2)})`);
    Logger.log(`Total échanges appliqués: ${nbSwapsAppliques} (swaps principaux + cycles MultiSwap)`);
    Logger.log(`${"=".repeat(60)}`);
    
    return {
      success: true,
      nbSwapsAppliques: nbSwapsAppliques,
      cyclesGeneraux: cyclesGeneraux,
      cyclesParite: cyclesParite,
      scoreEquilibre: etatFinal.scoreGlobal,
      tempsMs: tempsMs,
      nomFeuilleBilan: nomFeuilleBilan
    };
    
  } catch (e) {
    Logger.log(`ERREUR Moteur V2 Amélioré: ${e.message}\n${e.stack}`);
    return { success: false, error: e.message };
  }
}

// ==================================================================
// SECTION 2: PRÉPARATION DES DONNÉES
// ==================================================================

function V2_Ameliore_PreparerDonnees(config, criteresUI) {
  if (!__nirvanaDataService || typeof __nirvanaDataService.prepareData !== 'function') {
    throw new Error('NirvanaDataBackend indisponible pour préparer les données');
  }

  return __nirvanaDataService.prepareData({ config, criteresUI });
}

// ==================================================================
// SECTION 3: CALCUL DE L'ÉTAT ET SCORE D'ÉQUILIBRE
// ==================================================================

function V2_Ameliore_CalculerEtatGlobal(dataContext, config) {
  const CRITERES = ['COM', 'TRA', 'PART'];           // pour éviter les littéraux partout
  const poids     = config.V2_POIDS_EQUILIBRE || {}; // Fallback si la clé n'existe pas

  const etat = {
    classes     : {},
    scoreGlobal : 0,
    details     : {}        // (reste libre pour vos futurs ajouts)
  };

  let scoreTotal = 0;

  // ──────────────────────────────
  // Parcours des classes
  // ──────────────────────────────
  Object.keys(dataContext.classesState).forEach(classe => {
    const eleves   = dataContext.classesState[classe] || [];
    const effectif = eleves.length;

    // Sécurité : on ignore carrément une classe vide
    if (effectif === 0) {
      Logger.log(`Info : ${classe} est vide, on l'ignore.`);
      return;
    }

    // 1. Distributions par critère ──────────────────────────────────────────────
    const distributions = {};
    CRITERES.forEach(crit => {
      distributions[crit] = V2_Ameliore_CalculerDistribution(eleves, crit);
    });

    // 2. Parité  ────────────────────────────────────────────────────────────────
    const parite = V2_Ameliore_CalculerParite(eleves);

    // 3. Stats (LV2 & Options) — pour information uniquement
    const lv2Stats     = V2_Ameliore_CalculerStatsLV2(eleves);
    const optionsStats = V2_Ameliore_CalculerStatsOptions(eleves);

    // 4. Calcul du score de classe  ─────────────────────────────────────────────
    let scoreClasse   = 0;
    const detailsClasse = {};

    // a) Score distribution (COM / TRA / PART)
    CRITERES.forEach(crit => {
      const ciblesCrit = dataContext.ciblesParClasse?.[classe]?.[crit];

      if (!ciblesCrit) {
        Logger.log(`⚠️  Pas de cibles pour ${classe} – ${crit}. Score du critère = 0`);
        detailsClasse[crit] = 0;
        return;                           // critère suivant
      }

      const scoreCrit = V2_Ameliore_CalculerScoreDistribution(
        distributions[crit],
        ciblesCrit,
        effectif
      );

      scoreClasse          += (poids[crit] || 0) * scoreCrit;
      detailsClasse[crit]   = scoreCrit;
    });

    // b) Score parité
    const scoreParite = V2_Ameliore_CalculerScoreParite(parite);
    scoreClasse      += (poids.PARITE || 0) * scoreParite;
    detailsClasse.parite = scoreParite;

    // 5. Cache interne (si structure présente)
    if (dataContext.classeCaches?.[classe]) {
      dataContext.classeCaches[classe].dist   = Object.assign({}, distributions);
      dataContext.classeCaches[classe].parite = { ...parite };
      dataContext.classeCaches[classe].score  = scoreClasse;
    }

    // 6. Récap. dans l'objet état
    etat.classes[classe] = {
      effectif,
      distributions,
      cibles       : dataContext.ciblesParClasse?.[classe],
      parite,
      lv2Stats,
      optionsStats,
      score        : scoreClasse,
      details      : detailsClasse
    };

    scoreTotal += scoreClasse;
  });

  // ──────────────────────────────
  // Score global (moyenne)
  // ──────────────────────────────
  const nbClassesReelles = Object.keys(etat.classes).length || 1;  // évite division par 0
  etat.scoreGlobal       = scoreTotal / nbClassesReelles;
  dataContext.scoreGlobal = etat.scoreGlobal;                      // pour compat.

  return etat;
}

function V2_Ameliore_CalculerDistribution(eleves, critere) {
  const dist = {'1': 0, '2': 0, '3': 0, '4': 0};
  eleves.forEach(e => {
    const score = e[critere];
    if (score >= 1 && score <= 4) {
      dist[String(Math.round(score))]++;
    }
  });
  return dist;
}

function V2_Ameliore_CalculerScoreDistribution(distributionActuelle, cibles, effectif) {
  if (effectif === 0) return 100;
  
  let ecartTotal = 0;
  let countTotal = 0;
  
  ['1', '2', '3', '4'].forEach(score => {
    const actuel = distributionActuelle[score] || 0;
    const cible = cibles[score] || 0;
    
    // Écart absolu normalisé
    if (cible > 0) {
      ecartTotal += Math.abs(actuel - cible) / cible;
      countTotal++;
    } else if (actuel > 0) {
      // Pénalité si on a des élèves alors qu'on ne devrait pas
      ecartTotal += 1;
      countTotal++;
    }
  });
  
  if (countTotal === 0) return 100;
  
  // Score = 100 - (écart moyen * 100)
  const ecartMoyen = ecartTotal / countTotal;
  return Math.max(0, 100 * (1 - ecartMoyen));
}

function updateCachesApresSwap(cls1, cls2, dataContext, config) {
  const cache = dataContext.classeCaches;
  const poids = config.V2_POIDS_EQUILIBRE;
  
  [cls1, cls2].forEach(cl => {
    const eleves = dataContext.classesState[cl];
    const effectif = eleves.length;
    
    // 1. Recalculer distributions
    ['COM', 'TRA', 'PART'].forEach(critere => {
      const dist = {1:0, 2:0, 3:0, 4:0};
      eleves.forEach(e => {
        const score = Math.round(e[`niveau${critere}`] || 0);
        if (score >= 1 && score <= 4) dist[score]++;
      });
      cache[cl].dist[critere] = { ...dist }; // Copie pour éviter les références
    });
    
    // 2. Recalculer parité
    let f = 0, m = 0;
    eleves.forEach(e => {
      if (e.SEXE === 'F') f++;
      else if (e.SEXE === 'M') m++;
    });
    cache[cl].parite = { F: f, M: m, total: f + m };
    
    // 3. Recalculer score
    let scoreClasse = 0;
    ['COM', 'TRA', 'PART'].forEach(critere => {
      const scoreCritere = V2_Ameliore_CalculerScoreDistribution(
        cache[cl].dist[critere],
        dataContext.ciblesParClasse[cl][critere],
        effectif
      );
      scoreClasse += poids[critere] * scoreCritere;
    });
    
    const scoreParite = V2_Ameliore_CalculerScoreParite(cache[cl].parite);
    scoreClasse += poids.PARITE * scoreParite;
    cache[cl].score = scoreClasse;
  });
  
  // 4. Score global
  const total = Object.values(cache).reduce((s, c) => s + c.score, 0);
  dataContext.scoreGlobal = total / dataContext.nbClasses;
}

// ==================================================================
// SECTION 4: OPTIMISATION GLOBALE
// ==================================================================

function V2_Ameliore_OptimiserGlobal(dataContext, config) {
  const journalSwaps = [];
  let iteration = 0;
  let scoreActuel = V2_Ameliore_CalculerEtatGlobal(dataContext, config).scoreGlobal;
  let continuer = true;
  let derniereAmelioration = 0;
  
  Logger.log(`Score initial: ${scoreActuel.toFixed(2)}/100`);
  
  while (continuer && iteration < config.V2_MAX_ITERATIONS_GLOBAL) {
    iteration++;
    let meilleurSwap = null;
    let meilleureAmelioration = 0;
    
    // Stratégie: on commence par chercher les swaps entre classes ayant les plus gros écarts
    const classesParScore = V2_Ameliore_TrierClassesParScore(dataContext, config);
    
    // On teste en priorité les swaps entre classes mal équilibrées
    const nbClassesATester = Math.min(classesParScore.length, 10); // Limiter pour performance
    
    for (let i = 0; i < nbClassesATester; i++) {
      for (let j = i + 1; j < nbClassesATester; j++) {
        const classe1 = classesParScore[i].classe;
        const classe2 = classesParScore[j].classe;
        
        // Générer les swaps possibles entre ces deux classes
        const swapsPossibles = V2_Ameliore_GenererSwapsPossibles(
          classe1, classe2, dataContext, config
        );
        
        // Tester un échantillon aléatoire si trop de swaps possibles
        const swapsATester = swapsPossibles.length > 50 ? 
          V2_Ameliore_EchantillonnerSwaps(swapsPossibles, 50) : 
          swapsPossibles;
        
        for (const swap of swapsATester) {
          // ► tester
          V2_Ameliore_SimulerSwap(swap, dataContext);
          updateCachesApresSwap(swap.classe1, swap.classe2, dataContext, config);
          const amelioration = dataContext.scoreGlobal - scoreActuel;

          // ► toujours annuler juste après le test
          V2_Ameliore_AnnulerSwap(swap, dataContext);
          updateCachesApresSwap(swap.classe1, swap.classe2, dataContext, config);

          if (amelioration > meilleureAmelioration) {
            meilleureAmelioration = amelioration;
            meilleurSwap = swap;
          }
        }
      }
    }
    
    // Appliquer le meilleur swap trouvé
    if (meilleurSwap && meilleureAmelioration > config.V2_MIN_AMELIORATION) {
      V2_Ameliore_AppliquerSwapDansContext(meilleurSwap, dataContext);
      updateCachesApresSwap(meilleurSwap.classe1, meilleurSwap.classe2, dataContext, config);
      journalSwaps.push(meilleurSwap);
      scoreActuel = dataContext.scoreGlobal;
      derniereAmelioration = meilleureAmelioration;
      
      if (iteration % 10 === 0) {
        Logger.log(`Iteration ${iteration}: Score = ${scoreActuel.toFixed(2)}, Amélioration = +${meilleureAmelioration.toFixed(4)}`);
      }
    } else {
      continuer = false;
      Logger.log(`Optimisation terminée à l'itération ${iteration}. Score final: ${scoreActuel.toFixed(2)}`);
    }
    
    // Arrêt si on stagne
    if (iteration > 50 && derniereAmelioration < 0.00001) {
      continuer = false;
      Logger.log(`Arrêt pour stagnation à l'itération ${iteration}`);
    }
  }
  
  return journalSwaps;
}

function V2_Ameliore_AnalyserMotifSwap(swap, dataContext) {
  const motifs = [];
  ['COM', 'TRA', 'PART'].forEach(critere => {
    const score1 = Math.round(swap.eleve1[`niveau${critere}`] || 0);
    const score2 = Math.round(swap.eleve2[`niveau${critere}`] || 0);
    if (score1 !== score2) {
      motifs.push(`${critere}: ${score1}⇄${score2}`);
    }
  });
  return motifs.join(', ') || 'Équilibrage';
}

function V2_Ameliore_GenererSwapsPossibles(classe1, classe2, dataContext, config) {
  const swaps   = [];
  const eleves1 = dataContext.classesState[classe1];
  const eleves2 = dataContext.classesState[classe2];

  // État actuel
  const etat     = V2_Ameliore_CalculerEtatGlobal(dataContext, config);
  const dist1    = etat.classes[classe1].distributions;
  const dist2    = etat.classes[classe2].distributions;
  const cibles1  = dataContext.ciblesParClasse[classe1];
  const cibles2  = dataContext.ciblesParClasse[classe2];
  const parite1  = etat.classes[classe1].parite;
  const parite2  = etat.classes[classe2].parite;

  for (const e1 of eleves1) {
    if (e1.mobilite === 'FIXE' || e1.mobilite === 'SPEC') continue;

    for (const e2 of eleves2) {
      if (e2.mobilite === 'FIXE' || e2.mobilite === 'SPEC') continue;

      if (V2_Ameliore_SwapRespectContraintes(e1, e2, classe1, classe2, dataContext) &&
          V2_Ameliore_SwapPotentiellementBenefique(
                e1, e2,
                dist1, dist2,
                cibles1, cibles2,
                parite1, parite2)) {

        swaps.push({ eleve1: e1, eleve2: e2, classe1, classe2 });
      }
    }
  }
  return swaps;
}

function V2_Ameliore_SwapRespectContraintes(e1, e2, classe1, classe2, dataContext) {
  // PRIORITÉ 1: Utiliser la fonction existante respecteContraintes qui vérifie TOUT
  if (typeof respecteContraintes === 'function') {
    const result = respecteContraintes(
      e1, e2,
      dataContext.elevesValides,
      dataContext.structureData,
      null,
      dataContext.optionPools,
      dataContext.dissocMap
    );
    
    if (!result) {
      // La fonction a déjà loggé la raison du rejet
      return false;
    }
    return true;
  }
  
  // Fallback si respecteContraintes n'est pas disponible
  Logger.log("WARN: fonction respecteContraintes non disponible, vérifications basiques");
  
  // Vérifier les OPTIONS (CRITIQUE - doit respecter _STRUCTURE)
  if (e1.optionKey) {
    const pool1 = dataContext.optionPools[e1.optionKey];
    if (!pool1 || !pool1.includes(classe2.toUpperCase())) {
      Logger.log(`Rejet: ${e1.ID_ELEVE} a l'option ${e1.optionKey} qui n'est pas autorisée en ${classe2}`);
      return false;
    }
  }
  
  if (e2.optionKey) {
    const pool2 = dataContext.optionPools[e2.optionKey];
    if (!pool2 || !pool2.includes(classe1.toUpperCase())) {
      Logger.log(`Rejet: ${e2.ID_ELEVE} a l'option ${e2.optionKey} qui n'est pas autorisée en ${classe1}`);
      return false;
    }
  }
  
  // Vérifier les DISSOCIATIONS
  if (e1.DISSO) {
    const dissocSet2 = dataContext.dissocMap[classe2];
    if (dissocSet2 && dissocSet2.has(e1.DISSO)) {
      Logger.log(`Rejet: ${e1.ID_ELEVE} a le code DISSO ${e1.DISSO} incompatible avec ${classe2}`);
      return false;
    }
  }
  
  if (e2.DISSO) {
    const dissocSet1 = dataContext.dissocMap[classe1];
    if (dissocSet1 && dissocSet1.has(e2.DISSO)) {
      Logger.log(`Rejet: ${e2.ID_ELEVE} a le code DISSO ${e2.DISSO} incompatible avec ${classe1}`);
      return false;
    }
  }
  
  return true;
}

/**
 * Retourne true si l'échange rapproche *au moins une* des deux classes
 *   – de sa cible de distribution (COM / TRA / PART)              OU
 *   – de l'équilibre F/M                                         .
 *
 * @param  {Object}  e1, e2          – élèves candidats
 * @param  {Object}  dist1, dist2    – distributions actuelles de la classe1 / classe2
 * @param  {Object}  cibles1, cibles2– cibles par classe (scores)
 * @param  {Object}  parite1, parite2– {F, M, total} pour chaque classe
 */
function V2_Ameliore_SwapPotentiellementBenefique(
  e1, e2,
  dist1, dist2,
  cibles1, cibles2,
  parite1, parite2
) {
  /* ---------- 1.  COM / TRA / PART ----------------------------- */
  for (const critere of ['COM', 'TRA', 'PART']) {
    const score1 = String(Math.round(e1[`niveau${critere}`] || 0));
    const score2 = String(Math.round(e2[`niveau${critere}`] || 0));
    if (score1 === score2) continue;                       // pas d'impact

    // Classe 1 excédentaire & classe 2 déficitaire
    const excès1      = (dist1[critere][score1] || 0) > (cibles1[critere][score1] || 0);
    const besoin2     = (dist2[critere][score1] || 0) < (cibles2[critere][score1] || 0);

    // Symétrique : classe 2 excédentaire & classe 1 déficitaire
    const excès2      = (dist2[critere][score2] || 0) > (cibles2[critere][score2] || 0);
    const besoin1     = (dist1[critere][score2] || 0) < (cibles1[critere][score2] || 0);

    if ((excès1 && besoin2) || (excès2 && besoin1)) return true;
  }

  /* ---------- 2.  PARITÉ F / M  -------------------------------- */
  const deltaParite = p => Math.abs((p.F / p.total) - 0.5);

  // Avant échange
  const avant1 = deltaParite(parite1);
  const avant2 = deltaParite(parite2);

  // Après échange
  const deltaF1 = (e1.SEXE === 'F' ? -1 : 1) * (e1.SEXE !== e2.SEXE ? 1 : 0);
  const deltaF2 = -deltaF1; // effet miroir

  const apres1 = deltaParite({
    F: parite1.F + deltaF1,
    M: parite1.M - deltaF1,
    total: parite1.total
  });

  const apres2 = deltaParite({
    F: parite2.F + deltaF2,
    M: parite2.M - deltaF2,
    total: parite2.total
  });

  return (apres1 < avant1) || (apres2 < avant2);
}

// ==================================================================
// SECTION 5: UTILITAIRES ET HELPERS
// ==================================================================

function V2_Ameliore_SimulerSwap(swap, dataContext) {
  const idx1 = dataContext.classesState[swap.classe1].indexOf(swap.eleve1);
  const idx2 = dataContext.classesState[swap.classe2].indexOf(swap.eleve2);
  
  // Échanger
  dataContext.classesState[swap.classe1][idx1] = swap.eleve2;
  dataContext.classesState[swap.classe2][idx2] = swap.eleve1;
  
  // Mettre à jour les classes dans les objets élèves
  swap.eleve1.CLASSE = swap.classe2;
  swap.eleve2.CLASSE = swap.classe1;
}

function V2_Ameliore_AnnulerSwap(swap, dataContext) {
  const idx1 = dataContext.classesState[swap.classe2].indexOf(swap.eleve1);
  const idx2 = dataContext.classesState[swap.classe1].indexOf(swap.eleve2);
  
  // Remettre en place
  dataContext.classesState[swap.classe2][idx1] = swap.eleve2;
  dataContext.classesState[swap.classe1][idx2] = swap.eleve1;
  
  // Restaurer les classes
  swap.eleve1.CLASSE = swap.classe1;
  swap.eleve2.CLASSE = swap.classe2;
}

function V2_Ameliore_AppliquerSwapDansContext(swap, dataContext) {
  V2_Ameliore_SimulerSwap(swap, dataContext);
  
  // Mettre à jour la dissocMap
  if (typeof updateDissocMapForSwap === 'function') {
    updateDissocMapForSwap(dataContext.dissocMap, swap.eleve1, swap.classe1, swap.classe2);
    updateDissocMapForSwap(dataContext.dissocMap, swap.eleve2, swap.classe2, swap.classe1);
  }
  
  // Enregistrer les détails du swap
  swap.motif = V2_Ameliore_AnalyserMotifSwap(swap, dataContext);
}

// ==================================================================
// 5. AJOUTER CES FONCTIONS SI ELLES N'EXISTENT PAS
// ==================================================================

function V2_Ameliore_CalculerParite(eleves) {
  let nbF = 0, nbM = 0;
  eleves.forEach(e => {
    if (e.SEXE === 'F') nbF++;
    else if (e.SEXE === 'M') nbM++;
  });
  return { F: nbF, M: nbM, total: nbF + nbM };
}

function V2_Ameliore_CalculerScoreParite(parite) {
  if (parite.total === 0) return 100;               // sécurité

  const ecart = Math.abs(parite.F / parite.total - 0.5);

  /* 400 → écart de 10 %  ⇒  score 60 (au lieu de 80)  
     600 → écart de 10 %  ⇒  score 40
     800 → écart de 10 %  ⇒  score 20                          */
  const FACTEUR = 600;                              // <-- ajustez ici

  return Math.max(0, 100 - ecart * FACTEUR);
}

function V2_Ameliore_CalculerStatsLV2(eleves) {
  const stats = {};
  eleves.forEach(e => {
    const lv2 = e.LV2 || 'AUTRE';
    stats[lv2] = (stats[lv2] || 0) + 1;
  });
  return stats;
}

function V2_Ameliore_CalculerStatsOptions(eleves) {
  const stats = {};
  eleves.forEach(e => {
    const opt = e.OPT || 'SANS';
    stats[opt] = (stats[opt] || 0) + 1;
  });
  return stats;
}

function V2_Ameliore_CalculerDistributionGlobale(eleves, criteres) {
  const dist = {};
  criteres.forEach(crit => {
    dist[crit] = {'1': 0, '2': 0, '3': 0, '4': 0};
    eleves.forEach(e => {
      const score = e[`niveau${crit}`];
      if (score >= 1 && score <= 4) {
        dist[crit][String(Math.round(score))]++;
      }
    });
  });
  return dist;
}

function V2_Ameliore_TrierClassesParScore(dataContext, config) {
  const etat = V2_Ameliore_CalculerEtatGlobal(dataContext, config);
  const classesAvecScore = [];
  
  Object.keys(etat.classes).forEach(classe => {
    classesAvecScore.push({
      classe: classe,
      score: etat.classes[classe].score
    });
  });
  
  // Trier par score croissant (les moins bien équilibrées en premier)
  return classesAvecScore.sort((a, b) => a.score - b.score);
}

function V2_Ameliore_EchantillonnerSwaps(swaps, nbMax) {
  if (swaps.length <= nbMax) return swaps;
  
  const echantillon = [];
  const indices = new Set();
  
  while (echantillon.length < nbMax) {
    const idx = Math.floor(Math.random() * swaps.length);
    if (!indices.has(idx)) {
      indices.add(idx);
      echantillon.push(swaps[idx]);
    }
  }
  
  return echantillon;
}

// ────────────────────────────────────────────────────────────────
// 1. DÉDOUBLONNAGE ROBUSTE POUR MULTISWAP
// ────────────────────────────────────────────────────────────────
function V2_Ameliore_DedoublonnerSwaps(journal) {
  const dejaVu = new Set();           // IDs déjà traités
  const resultat = [];
  
  // On parcourt la liste à l'envers pour garder la DERNIÈRE occurrence
  for (let i = journal.length - 1; i >= 0; i--) {
    const s = journal[i];
    
    // Extraction sécurisée des IDs
    const id1 = s.eleve1?.ID_ELEVE || s.eleve1ID || `UNKNOWN_${i}_1`;
    const id2 = s.eleve2?.ID_ELEVE || s.eleve2ID || `UNKNOWN_${i}_2`;
    
    // Vérifier si déjà traité
    if (dejaVu.has(id1) || dejaVu.has(id2)) {
      Logger.log(`Dédoublonnage: Skip swap ${i} (${id1} ↔ ${id2}) - Déjà traité`);
      continue;
    }
    
    // Ajouter au résultat (en tête pour garder l'ordre)
    resultat.unshift(s);
    dejaVu.add(id1);
    dejaVu.add(id2);
  }
  
  Logger.log(`Dédoublonnage: ${journal.length} → ${resultat.length} swaps (${journal.length - resultat.length} supprimés)`);
  return resultat;
}

// ==================================================================
// SECTION 6: APPLICATION DES SWAPS ET GÉNÉRATION DU BILAN
// ==================================================================

function V2_Ameliore_AppliquerSwaps(journalSwaps, dataContext, config) {
  if (!journalSwaps || journalSwaps.length === 0) {
    Logger.log("V2 Amélioré: Aucun swap à appliquer");
    return 0;
  }

  // ►► PATCH 1 : dé-doublonner ◄◄
  const swapsFinaux = V2_Ameliore_DedoublonnerSwaps(journalSwaps);
  Logger.log(`V2 Amélioré: Application de ${swapsFinaux.length} swaps (après dé-doublonnage)…`);

  // Préparer les swaps pour la fonction existante
  swapsFinaux.forEach(swap => {
    swap.eleve1ID   = swap.eleve1.ID_ELEVE;
    swap.eleve2ID   = swap.eleve2.ID_ELEVE;
    swap.eleve1Nom  = swap.eleve1.NOM;
    swap.eleve2Nom  = swap.eleve2.NOM;
    swap.oldClasseE1 = swap.classe1;
    swap.oldClasseE2 = swap.classe2;
    swap.newClasseE1 = swap.classe2;
    swap.newClasseE2 = swap.classe1;
  });

  // Utiliser la fonction existante V2
  if (typeof V2_executerSwapsDansOngletsRobust === 'function') {
    return V2_executerSwapsDansOngletsRobust(swapsFinaux, dataContext, config);
  }

  // Sinon utiliser la version V14
  if (typeof executerSwapsDansOnglets === 'function') {
    try {
      executerSwapsDansOnglets(swapsFinaux);
      return swapsFinaux.length;
    } catch (e) {
      Logger.log(`Erreur application swaps V14: ${e.message}`);
      return 0;
    }
  }

  Logger.log("ERREUR: Aucune fonction d'application de swaps disponible");
  return 0;
}

function V2_Ameliore_GenererBilan(dataContext, etatInitial, etatFinal, journalSwaps, config, multiswapInfo = {}) {
  const nomFeuille = config.V2_SHEET_NAME_BILAN || "_SWAPS_V2";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(nomFeuille);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(nomFeuille);
  }
  
  // Titre principal
  const titres = [
    [`=== BILAN NIRVANA V2 AMÉLIORÉ + MULTISWAP ===`, '', '', '', '', '', '', '', '', ''],
    [`Score: ${etatInitial.scoreGlobal.toFixed(2)} → ${etatFinal.scoreGlobal.toFixed(2)} (+${(etatFinal.scoreGlobal - etatInitial.scoreGlobal).toFixed(2)})`, '', '', '', '', '', '', '', '', ''],
    [`Swaps principaux: ${journalSwaps.length} | Cycles généraux: ${multiswapInfo.cyclesGeneraux || 0} | Cycles parité: ${multiswapInfo.cyclesParite || 0}`, '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '']
  ];
  
  // En-tête des swaps
  const headers = [['#', 'Élève 1', 'Classe Origine', '→', 'Classe Destination', 'Élève 2', 'Classe Origine', '→', 'Classe Destination', 'Motif']];
  
  // Liste des swaps avec formatage sécurisé
  const data = [];
  journalSwaps.forEach((swap, idx) => {
    try {
      // Gestion sécurisée des propriétés manquantes
      const eleve1Nom = (swap.eleve1?.NOM || swap.eleve1Nom || 'Inconnu') + 
                       (swap.eleve1?.PRENOM ? ' ' + swap.eleve1.PRENOM : '');
      const eleve2Nom = (swap.eleve2?.NOM || swap.eleve2Nom || 'Inconnu') + 
                       (swap.eleve2?.PRENOM ? ' ' + swap.eleve2.PRENOM : '');
      
      const ligne = [
        idx + 1,
        eleve1Nom,
        swap.classe1 || swap.oldClasseE1 || 'N/A',
        '→',
        swap.classe2 || swap.newClasseE1 || 'N/A',
        eleve2Nom,
        swap.classe2 || swap.oldClasseE2 || 'N/A',
        '→',
        swap.classe1 || swap.newClasseE2 || 'N/A',
        swap.motif || 'Équilibrage'
      ];
      
      // Vérifier que la ligne a exactement 10 colonnes
      while (ligne.length < 10) ligne.push('');
      if (ligne.length > 10) ligne.splice(10);
      
      data.push(ligne);
      
    } catch (e) {
      Logger.log(`Erreur formatage swap ${idx}: ${e.message}`);
      // Ligne de fallback
      data.push([
        idx + 1, 
        'Erreur', 
        swap.classe1 || 'N/A', 
        '→', 
        swap.classe2 || 'N/A', 
        'Erreur', 
        swap.classe2 || 'N/A', 
        '→', 
        swap.classe1 || 'N/A', 
        swap.motif || 'Erreur'
      ]);
    }
  });
  
  // Si aucun swap, ajouter une ligne indicative
  if (data.length === 0) {
    data.push(['', 'Aucun échange effectué', '', '', '', '', '', '', '', '']);
  }
  
  // Construire le tableau final
  const allData = [...titres, ...headers, ...data];
  
  try {
    if (allData.length > 0) {
      // Vérifier la cohérence des données
      const maxCols = Math.max(...allData.map(row => row.length));
      Logger.log(`Génération bilan: ${allData.length} lignes, ${maxCols} colonnes max`);
      
      // Normaliser toutes les lignes à 10 colonnes
      const normalizedData = allData.map(row => {
        const normalizedRow = [...row];
        while (normalizedRow.length < 10) normalizedRow.push('');
        return normalizedRow.slice(0, 10);
      });
      
      sheet.getRange(1, 1, normalizedData.length, 10).setValues(normalizedData);
      
      // Mise en forme
      if (normalizedData.length >= 4) {
        sheet.getRange(1, 1, 3, 10).setFontWeight('bold').setBackground('#e6f3ff');
      }
      if (normalizedData.length >= 5) {
        sheet.getRange(5, 1, 1, 10).setFontWeight('bold').setBackground('#f0f0f0');
      }
      sheet.autoResizeColumns(1, 10);
    }
  } catch (e) {
    Logger.log(`Erreur écriture bilan: ${e.message}`);
    // Écriture de fallback simple
    sheet.getRange(1, 1).setValue(`=== BILAN NIRVANA V2 (${journalSwaps.length} échanges) ===`);
    sheet.getRange(2, 1).setValue(`Score: ${etatInitial.scoreGlobal.toFixed(2)} → ${etatFinal.scoreGlobal.toFixed(2)}`);
  }
  
  // Message de confirmation
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${journalSwaps.length} échanges effectués (swaps + cycles) - Voir onglet "${nomFeuille}"`,
    'Optimisation terminée',
    5
  );
  
  return nomFeuille;
}

function V2_Ameliore_LogEtat(phase, etat, config) {
  Logger.log(`\n=== ÉTAT ${phase} ===`);
  Logger.log(`Score global: ${etat.scoreGlobal.toFixed(2)}/100`);
  
  // Top 3 meilleures et pires classes
  const classesTriees = Object.keys(etat.classes)
    .map(c => ({classe: c, score: etat.classes[c].score}))
    .sort((a, b) => b.score - a.score);
  
  Logger.log("Top 3 classes:");
  classesTriees.slice(0, 3).forEach(c => {
    Logger.log(`  ${c.classe}: ${c.score.toFixed(2)}`);
  });
  
  Logger.log("Bottom 3 classes:");
  classesTriees.slice(-3).forEach(c => {
    Logger.log(`  ${c.classe}: ${c.score.toFixed(2)}`);
  });
  
  Logger.log(`=== FIN ÉTAT ${phase} ===\n`);
}

// ==================================================================
// FONCTION D'AIGUILLAGE POUR COMPATIBILITÉ
// ==================================================================

function lancerOptimisationVarianteB_Wrapper(criteresUI) {
  Logger.log("Redirection vers Nirvana V2 Amélioré + MultiSwap...");
  return lancerOptimisationNirvanaV2_UI(criteresUI);
}

// ==================================================================
// C'EST TOUT ! TEST RAPIDE :
// ==================================================================

function testDeltaScore() {
  const cfg = getConfig();
  const ctx = V2_Ameliore_PreparerDonnees(cfg);

  // 1️⃣ Calculer obligatoirement le score AVANT de lire/mettre en cache
  const etatActuel   = V2_Ameliore_CalculerEtatGlobal(ctx, cfg);
  const nouveauScore = etatActuel.scoreGlobal;

  // 2️⃣ ­Si le cache n'existe pas encore on l'initialise ici
  if (ctx.scoreGlobal == null) {
    ctx.scoreGlobal = nouveauScore;
    Logger.log(`🏁 1er passage : score initial enregistré : ${ctx.scoreGlobal.toFixed(2)}`);
    return;
  }

  // 3️⃣ Calcul du delta par rapport à la valeur précédente
  const delta = nouveauScore - ctx.scoreGlobal;
  Logger.log(
    `Ancien : ${ctx.scoreGlobal.toFixed(2)}  —  Nouveau : ${nouveauScore.toFixed(2)}  →  Δ : ${delta.toFixed(2)}`
  );

  // 4️⃣ Mettre à jour le cache pour la prochaine exécution
  ctx.scoreGlobal = nouveauScore;
}

/**
 * ==================================================================
 * AJOUTS POUR NIRVANA_V2_AMELIOREE.GS - CALCUL DES STATS COMME PARITÉ
 * ==================================================================
 * À ajouter à la fin du fichier Nirvana_V2_Amelioree.gs
 * À appeler à la fin de V2_Ameliore_OptimisationEngine
 */

// Constantes de style reprises du module Parité (style "Maquette")
const V2_STATS_STYLE = {
  HEADER_BG: "#C6E0B4",
  SEXE_F_COLOR: "#F28EA8", // Rose Maquette
  SEXE_M_COLOR: "#4F81BD", // Bleu Maquette
  LV2_ESP_COLOR: "#E59838", // Orange pour ESP
  LV2_AUTRE_COLOR: "#A3E4D7", // Bleu-vert clair pour autres
  SCORE_COLORS: {
    1: "#FF0000", // Rouge
    2: "#FFD966", // Jaune
    3: "#3CB371", // Vert moyen
    4: "#006400"  // Vert foncé
  },
  SCORE_FONT_COLORS: {
    1: "#FFFFFF", // Blanc
    2: "#000000", // Noir
    3: "#FFFFFF", // Blanc
    4: "#FFFFFF"  // Blanc
  }
};

/**
 * Fonction principale de calcul des statistiques Nirvana V2
 * À appeler à la fin de V2_Ameliore_OptimisationEngine
 */
function V2_calculerStatistiquesFinales() {
  Logger.log("=== DÉBUT CALCUL STATISTIQUES NIRVANA V2 (style Parité) ===");
  
  try {
    const testSheets = V2_getTestSheets();
    if (testSheets.length === 0) {
      Logger.log("Aucun onglet TEST trouvé pour les stats Nirvana V2");
      return;
    }
    
    Logger.log(`Calcul des stats Nirvana V2 pour ${testSheets.length} onglets TEST`);
    
    // Calculer et écrire les stats pour chaque onglet TEST
    testSheets.forEach(sheet => {
      V2_calculerEtEcrireStatsSheet(sheet);
    });
    
    Logger.log("=== FIN CALCUL STATISTIQUES NIRVANA V2 ===");
    
  } catch (e) {
    Logger.log(`Erreur V2_calculerStatistiquesFinales: ${e.message}`);
    Logger.log(e.stack);
  }
}

/**
 * Récupère les onglets TEST pour Nirvana V2
 */
function V2_getTestSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  const testSuffix = config?.TEST_SUFFIX || "TEST";
  const testSuffixRegex = new RegExp(testSuffix + '$', 'i');
  
  return ss.getSheets().filter(sheet => {
    const name = sheet.getName();
    return testSuffixRegex.test(name);
  });
}

/**
 * Calcule et écrit les statistiques pour une feuille (style Parité)
 */
function V2_calculerEtEcrireStatsSheet(sheet) {
  try {
    const sheetName = sheet.getName();
    Logger.log(`Calcul stats Nirvana V2 pour ${sheetName}`);
    
    // 1. Identifier les colonnes
    const colMap = V2_identifierColonnes(sheet);
    if (!colMap.valide) {
      Logger.log(`Colonnes requises manquantes dans ${sheetName}`);
      return;
    }
    
    // 2. Compter les élèves
    const lastRow = sheet.getLastRow();
    let lastDataRow = 1;
    if (lastRow > 1) {
      // Trouver la dernière ligne avec des données
      const idsColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      for (let i = idsColumn.length - 1; i >= 0; i--) {
        if (String(idsColumn[i]).trim()) {
          lastDataRow = i + 2;
          break;
        }
      }
    }
    const nRows = lastDataRow > 1 ? lastDataRow - 1 : 0;
    
    if (nRows === 0) {
      Logger.log(`Aucun élève dans ${sheetName}`);
      return;
    }
    
    // 3. Calculer les statistiques
    const stats = V2_calculateSheetStats(sheet, nRows, colMap);
    if (!stats) {
      Logger.log(`Échec calcul stats pour ${sheetName}`);
      return;
    }
    
    // 4. Écrire les statistiques
    const statsRow = lastDataRow + 2; // Ligne séparée + 1
    V2_writeSheetStats(sheet, statsRow, colMap, stats);
    
    Logger.log(`Stats Nirvana V2 écrites pour ${sheetName} à la ligne ${statsRow}`);
    
  } catch (e) {
    Logger.log(`Erreur stats Nirvana V2 pour ${sheet.getName()}: ${e.message}`);
  }
}

/**
 * Identifie les colonnes importantes dans la feuille pour Nirvana V2
 */
function V2_identifierColonnes(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h).trim().toUpperCase());
  
  const colMap = {
    ID_ELEVE: -1,
    NOM_PRENOM: -1,
    SEXE: -1,
    LV2: -1,
    OPT: -1,
    COM: -1,
    TRA: -1,
    PART: -1,
    ABS: -1,
    valide: false
  };
  
  // Recherche des colonnes avec fallbacks spécifiques à la structure utilisée par V2
  headers.forEach((header, index) => {
    // Utiliser la même logique que chargerElevesEtClasses dans V14
    if (header.includes("ID_ELEVE") || header === "ID") colMap.ID_ELEVE = index + 1;
    else if (header.includes("NOM") && (header.includes("PRENOM") || header.includes("&"))) colMap.NOM_PRENOM = index + 1;
    else if (header === "SEXE") colMap.SEXE = index + 1;
    else if (header === "LV2") colMap.LV2 = index + 1;
    else if (header === "OPT") colMap.OPT = index + 1;
    else if (header === "COM") colMap.COM = index + 1;
    else if (header === "TRA") colMap.TRA = index + 1;
    else if (header === "PART") colMap.PART = index + 1;
    else if (header === "ABS") colMap.ABS = index + 1;
  });
  
  // Vérifier que les colonnes essentielles existent
  colMap.valide = (colMap.ID_ELEVE > 0 && colMap.SEXE > 0);
  
  return colMap;
}

/**
 * Calcule les statistiques détaillées pour une feuille (Nirvana V2)
 */
function V2_calculateSheetStats(sheet, numDataRows, colMap) {
  const data = sheet.getRange(2, 1, numDataRows, sheet.getMaxColumns()).getValues();
  const validRows = data.filter(r => String(r[0]).trim() !== ""); // Filtrer sur ID_ELEVE
  
  if (validRows.length === 0) return null;
  
  // Fonction helper pour extraire les données d'une colonne
  const getColData = (colKey) => colMap[colKey] > 0 ? validRows.map(r => r[colMap[colKey] - 1]) : [];
  
  // Fonction helper pour compter les valeurs
  const countValues = (colData, value) => colData.filter(cell => 
    String(cell).trim().toUpperCase() === String(value).toUpperCase()).length;
  
  // Fonction helper pour les valeurs non vides
  const countNonEmpty = (colData) => colData.filter(cell => String(cell).trim() !== "").length;
  
  // Fonction helper pour convertir en nombres
  const toNumericArray = (colData) => colData.map(cell => {
    const num = Number(String(cell).replace(',', '.'));
    return isNaN(num) ? 0 : num;
  });
  
  // Fonction helper pour calculer la moyenne
  const calculateAverage = (numArray) => {
    const filtered = numArray.filter(n => n > 0);
    return filtered.length > 0 ? filtered.reduce((s, v) => s + v, 0) / filtered.length : 0;
  };
  
  // Données par colonne
  const sexeData = getColData('SEXE');
  const lv2Data = getColData('LV2');
  const optData = getColData('OPT');
  
  const stats = {
    genreCounts: [countValues(sexeData, 'F'), countValues(sexeData, 'M')],
    lv2Counts: [
      countValues(lv2Data, 'ESP'),
      lv2Data.length - countValues(lv2Data, 'ESP') - countValues(lv2Data, '') // Autres non-vides
    ],
    optionsCounts: [countNonEmpty(optData)],
    criteresScores: {},
    criteresMoyennes: []
  };
  
  // Calcul des scores par critère (identique à V14)
  ['COM', 'TRA', 'PART', 'ABS'].forEach(critKey => {
    const critDataNum = toNumericArray(getColData(critKey));
    stats.criteresScores[critKey] = {};
    
    // Compter chaque score 1-4
    for (let score = 1; score <= 4; score++) {
      stats.criteresScores[critKey][score] = critDataNum.filter(val => val === score).length;
    }
    
    // Calculer la moyenne
    stats.criteresMoyennes.push(calculateAverage(critDataNum));
  });
  
  // S'assurer qu'on a 4 moyennes
  while (stats.criteresMoyennes.length < 4) {
    stats.criteresMoyennes.push(0);
  }
  
  return stats;
}

/**
 * Écrit les statistiques dans la feuille avec le formatage Maquette (Nirvana V2)
 */
function V2_writeSheetStats(sheet, row, colMap, stats) {
  const maxCol = Math.max(
    colMap.NOM_PRENOM || 1,
    colMap.SEXE || 1,
    colMap.LV2 || 1,
    colMap.OPT || 1,
    colMap.COM || 1,
    colMap.TRA || 1,
    colMap.PART || 1,
    colMap.ABS || 1
  );
  
  if (sheet.getMaxRows() < row + 7) return; // Pas assez de place
  
  // Nettoyer la zone des statistiques
  sheet.getRange(row, 1, 7, maxCol).clearContent().clearFormat();
  
  // Fonction helper pour définir valeur et style
  const set = (r, c, v, style = {}) => {
    if (c > 0 && c <= sheet.getMaxColumns() && r <= sheet.getMaxRows()) {
      const cell = sheet.getRange(r, c);
      cell.setValue(v);
      if (style.bold) cell.setFontWeight('bold');
      if (style.align) cell.setHorizontalAlignment(style.align);
      if (style.bg) cell.setBackground(style.bg);
      if (style.fg) cell.setFontColor(style.fg);
      if (style.fmt) cell.setNumberFormat(style.fmt);
      if (style.italic) cell.setFontStyle('italic');
    }
  };
  
  if (!stats) {
    set(row, 1, 'Pas de données', { italic: true });
    return;
  }
  
  // Statistiques de genre/langues sur le côté (colonnes E et F)
  if (colMap.SEXE > 0) {
    // Filles
    set(row, 5, stats.genreCounts[0], { 
      align: 'center', 
      bg: V2_STATS_STYLE.SEXE_F_COLOR, 
      bold: true 
    });
    
    // Garçons
    set(row + 1, 5, stats.genreCounts[1], { 
      align: 'center', 
      bg: V2_STATS_STYLE.SEXE_M_COLOR, 
      fg: 'white', 
      bold: true 
    });
  }
  
  if (colMap.LV2 > 0) {
    // ESP
    set(row, 6, stats.lv2Counts[0], { 
      align: 'center', 
      bg: V2_STATS_STYLE.LV2_ESP_COLOR, 
      bold: true 
    });
    
    // Autres LV2
    set(row + 1, 6, stats.lv2Counts[1], { 
      align: 'center', 
      bg: V2_STATS_STYLE.LV2_AUTRE_COLOR, 
      bold: true 
    });
  }
  
  if (colMap.OPT > 0) {
    // Options total
    set(row, 7, stats.optionsCounts[0], { 
      align: 'center', 
      bold: true 
    });
  }
  
  // Scores par critère dans leurs colonnes respectives
  ['COM', 'TRA', 'PART', 'ABS'].forEach((critKey, index) => {
    const columnIndex = colMap[critKey];
    if (columnIndex > 0) {
      // Score 4 (vert foncé, police blanche)
      set(row, columnIndex, stats.criteresScores[critKey][4], { 
        align: 'center', 
        bg: V2_STATS_STYLE.SCORE_COLORS[4], 
        bold: true, 
        fg: V2_STATS_STYLE.SCORE_FONT_COLORS[4] 
      });
      
      // Score 3 (vert clair)
      set(row + 1, columnIndex, stats.criteresScores[critKey][3], { 
        align: 'center', 
        bg: V2_STATS_STYLE.SCORE_COLORS[3], 
        bold: true,
        fg: V2_STATS_STYLE.SCORE_FONT_COLORS[3]
      });
      
      // Score 2 (jaune)
      set(row + 2, columnIndex, stats.criteresScores[critKey][2], { 
        align: 'center', 
        bg: V2_STATS_STYLE.SCORE_COLORS[2], 
        bold: true,
        fg: V2_STATS_STYLE.SCORE_FONT_COLORS[2]
      });
      
      // Score 1 (rouge)
      set(row + 3, columnIndex, stats.criteresScores[critKey][1], { 
        align: 'center', 
        bg: V2_STATS_STYLE.SCORE_COLORS[1], 
        bold: true,
        fg: V2_STATS_STYLE.SCORE_FONT_COLORS[1]
      });
      
      // Moyenne
      set(row + 4, columnIndex, stats.criteresMoyennes[index], { 
        align: 'center', 
        bold: true, 
        fmt: '#,##0.00' 
      });
    }
  });
}

/**
 * ==================================================================
 * FONCTION DE TEST INTÉGRÉE
 * ==================================================================
 */
function testNirvanaV2MultiSwapIntegre() {
  Logger.log("=== TEST NIRVANA V2 + MULTISWAP INTÉGRÉ ===");
  
  try {
    const resultat = lancerOptimisationNirvanaV2_UI();
    
    if (resultat.success) {
      Logger.log("✅ Test réussi !");
      Logger.log(`Swaps: ${resultat.nbSwapsAppliques}`);
      Logger.log(`Cycles généraux: ${resultat.cyclesGeneraux}`);
      Logger.log(`Cycles parité: ${resultat.cyclesParite}`);
      Logger.log(`Score final: ${resultat.scoreEquilibre}`);
    } else {
      Logger.log("❌ Test échoué: " + resultat.error);
    }
    
  } catch (e) {
    Logger.log("❌ Erreur test: " + e.message);
  }
}

/**
 * ==================================================================
 * FONCTIONS WRAPPER POUR MULTISWAP AVEC RETOUR DÉTAILLÉ
 * ==================================================================
 */

/**
 * Wrapper pour MultiSwap général qui retourne les détails des échanges
 */
function V2_Ameliore_MultiSwap_AvecRetourSwaps(dataCtx, cfg) {
  const swapsDetailles = [];
  let nbCycles = 0;
  
  // Sauvegarder l'état original pour comparaison
  const classesStateOriginal = {};
  Object.keys(dataCtx.classesState).forEach(classe => {
    classesStateOriginal[classe] = [...dataCtx.classesState[classe]];
  });
  
  // Appeler la fonction MultiSwap originale si elle existe
  if (typeof V2_Ameliore_MultiSwap === 'function') {
    nbCycles = V2_Ameliore_MultiSwap(dataCtx, cfg);
    
    // Détecter les changements et les convertir en format swap
    if (nbCycles > 0) {
      swapsDetailles.push(..._detecterChangementsEtConvertir(
        classesStateOriginal, 
        dataCtx.classesState, 
        'MultiSwap-Général'
      ));
    }
  }
  
  return {
    nbCycles: nbCycles,
    swapsDetailles: swapsDetailles
  };
}

/**
 * Wrapper pour MultiSwap parité qui retourne les détails des échanges
 */
function V2_Ameliore_MultiSwap4_Parite_AvecRetourSwaps(dataCtx, cfg) {
  const swapsDetailles = [];
  let nbCycles = 0;
  
  // Sauvegarder l'état original pour comparaison
  const classesStateOriginal = {};
  Object.keys(dataCtx.classesState).forEach(classe => {
    classesStateOriginal[classe] = [...dataCtx.classesState[classe]];
  });
  
  // Appeler la fonction MultiSwap parité originale si elle existe
  if (typeof V2_Ameliore_MultiSwap4_Parite === 'function') {
    nbCycles = V2_Ameliore_MultiSwap4_Parite(dataCtx, cfg);
    
    // Détecter les changements et les convertir en format swap
    if (nbCycles > 0) {
      swapsDetailles.push(..._detecterChangementsEtConvertir(
        classesStateOriginal, 
        dataCtx.classesState, 
        'MultiSwap-Parité'
      ));
    }
  }
  
  return {
    nbCycles: nbCycles,
    swapsDetailles: swapsDetailles
  };
}

/**
 * Détecte les changements entre deux états et les convertit en format swap
 */
function _detecterChangementsEtConvertir(avant, apres, motifBase) {
  const swaps = [];
  const dejaTraites = new Set();
  
  // Parcourir toutes les classes pour détecter les changements
  Object.keys(avant).forEach(classe => {
    const elevesAvant = avant[classe] || [];
    const elevesApres = apres[classe] || [];
    
    // Trouver les élèves qui ont quitté cette classe
    elevesAvant.forEach(eleveAvant => {
      if (dejaTraites.has(eleveAvant.ID_ELEVE)) return;
      
      // Chercher où cet élève se trouve maintenant
      const eleveApres = elevesApres.find(e => e.ID_ELEVE === eleveAvant.ID_ELEVE);
      
      if (!eleveApres) {
        // L'élève a quitté cette classe, chercher sa nouvelle classe
        Object.keys(apres).forEach(nouvelleClasse => {
          if (nouvelleClasse === classe) return;
          
          const eleveEnNouvPos = apres[nouvelleClasse].find(e => e.ID_ELEVE === eleveAvant.ID_ELEVE);
          if (eleveEnNouvPos) {
            // Trouver l'élève qui a pris sa place (swap)
            const elevesNouvelleClasseAvant = avant[nouvelleClasse] || [];
            const elevesQuitteNouvelleClasse = elevesNouvelleClasseAvant.filter(e => 
              !apres[nouvelleClasse].find(ae => ae.ID_ELEVE === e.ID_ELEVE)
            );
            
            // Chercher lequel de ces élèves est maintenant dans la classe originale
            elevesQuitteNouvelleClasse.forEach(candidat => {
              const candidatDansClasseOriginale = elevesApres.find(e => e.ID_ELEVE === candidat.ID_ELEVE);
              if (candidatDansClasseOriginale && !dejaTraites.has(candidat.ID_ELEVE)) {
                // Trouvé un swap !
                swaps.push({
                  eleve1: eleveAvant,
                  eleve2: candidat,
                  eleve1ID: eleveAvant.ID_ELEVE,
                  eleve2ID: candidat.ID_ELEVE,
                  eleve1Nom: eleveAvant.NOM || eleveAvant.ID_ELEVE,
                  eleve2Nom: candidat.NOM || candidat.ID_ELEVE,
                  classe1: classe,
                  classe2: nouvelleClasse,
                  oldClasseE1: classe,
                  oldClasseE2: nouvelleClasse,
                  newClasseE1: nouvelleClasse,
                  newClasseE2: classe,
                  motif: motifBase
                });
                
                dejaTraites.add(eleveAvant.ID_ELEVE);
                dejaTraites.add(candidat.ID_ELEVE);
              }
            });
          }
        });
      }
    });
  });
  
  Logger.log(`_detecterChangementsEtConvertir: ${swaps.length} swaps détectés pour ${motifBase}`);
  return swaps;
}

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
          niveauCOM: colIndexes.COM !== undefined ? parseFloat(row[colIndexes.COM]) || 0 : 0,
          niveauTRA: colIndexes.TRA !== undefined ? parseFloat(row[colIndexes.TRA]) || 0 : 0,
          niveauPART: colIndexes.PART !== undefined ? parseFloat(row[colIndexes.PART]) || 0 : 0,
          niveauABS: colIndexes.ABS !== undefined ? parseFloat(row[colIndexes.ABS]) || 0 : 0,
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
function V2_Ameliore_PreparerDonnees_AvecSEXE(config, criteresUI) {
  if (!__nirvanaDataService || typeof __nirvanaDataService.prepareData !== 'function') {
    throw new Error('NirvanaDataBackend indisponible pour préparer les données (SEXE)');
  }

  return __nirvanaDataService.prepareData({ config, criteresUI });
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
function V2_Ameliore_PreparerDonnees(config, criteresUI) {
  return V2_Ameliore_PreparerDonnees_AvecSEXE(config, criteresUI);
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

/**
 * ==================================================================
 *    FICHIER  : Fix_Detection_Sexe_Parite.gs
 *    VERSION  : 1.0
 *    OBJET    : Corriger la détection de la colonne SEXE
 * ==================================================================
 */

'use strict';

/**
 * Fonction de diagnostic pour vérifier les colonnes SEXE
 */
function diagnosticColonneSexe() {
  Logger.log("=== DIAGNOSTIC COLONNE SEXE ===");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  const testSuffix = config?.TEST_SUFFIX || "TEST";
  
  // Parcourir tous les onglets TEST
  const sheets = ss.getSheets().filter(s => s.getName().endsWith(testSuffix));
  
  sheets.forEach(sheet => {
    Logger.log(`\n--- Analyse de ${sheet.getName()} ---`);
    
    // Lire les en-têtes
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log(`En-têtes: ${headers.join(' | ')}`);
    
    // Chercher la colonne SEXE ou ses variantes
    const sexeVariants = ['SEXE', 'Sexe', 'SEX', 'GENRE', 'Genre', 'H/F', 'G/F', 'M/F', 'F/M', 'F/G'];
    let sexeCol = -1;
    
    headers.forEach((header, idx) => {
      const h = String(header).trim();
      if (sexeVariants.some(v => h.toUpperCase().includes(v.toUpperCase()))) {
        sexeCol = idx + 1;
        Logger.log(`Colonne SEXE trouvée: position ${sexeCol} (${header})`);
      }
    });
    
    if (sexeCol > 0) {
      // Lire quelques valeurs
      const lastRow = sheet.getLastRow();
      const sampleSize = Math.min(10, lastRow - 1);
      if (sampleSize > 0) {
        const values = sheet.getRange(2, sexeCol, sampleSize, 1).getValues().flat();
        Logger.log(`Échantillon valeurs: ${values.join(', ')}`);
        
        // Compter les F et M
        const f = values.filter(v => String(v).trim().toUpperCase() === 'F').length;
        const m = values.filter(v => ['M', 'G', 'H'].includes(String(v).trim().toUpperCase())).length;
        Logger.log(`Dans l'échantillon: ${f} F, ${m} M/G/H`);
      }
    } else {
      Logger.log("⚠️ AUCUNE colonne SEXE détectée !");
    }
  });
  
  Logger.log("\n=== FIN DIAGNOSTIC ===");
}

/**
 * Injection forcée du champ SEXE dans le contexte de données
 */
function injecterSexeForcee(dataContext) {
  Logger.log("\n=== INJECTION FORCÉE DU SEXE ===");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SEXE_VARIANTS = ['SEXE', 'Sexe', 'SEX', 'GENRE', 'Genre', 'H/F', 'G/F', 'M/F', 'F/M', 'F/G'];
  let totalInjected = 0;
  let totalEleves = 0;

  Object.keys(dataContext.classesState).forEach(classeName => {
    Logger.log(`\nTraitement de ${classeName}...`);
    
    const sheet = ss.getSheetByName(classeName);
    if (!sheet) {
      Logger.log(`⚠️ Feuille ${classeName} non trouvée`);
      return;
    }

    // Lire toutes les données
    const allData = sheet.getDataRange().getValues();
    if (allData.length < 2) return;

    const headers = allData[0];
    
    // Chercher colonne SEXE
    let sexeIdx = -1;
    headers.forEach((h, idx) => {
      if (sexeIdx === -1 && h) {
        const headerStr = String(h).trim();
        if (SEXE_VARIANTS.some(v => headerStr.toUpperCase().includes(v.toUpperCase()))) {
          sexeIdx = idx;
          Logger.log(`Colonne SEXE trouvée: position ${idx + 1} (${h})`);
        }
      }
    });
    
    // Chercher colonne ID_ELEVE
    const idIdx = headers.findIndex(h => 
      h && String(h).trim().toUpperCase().includes('ID_ELEVE')
    );
    
    if (sexeIdx === -1) {
      Logger.log(`❌ Pas de colonne SEXE dans ${classeName}`);
      return;
    }
    
    if (idIdx === -1) {
      Logger.log(`❌ Pas de colonne ID_ELEVE dans ${classeName}`);
      return;
    }

    // Parcourir les élèves et injecter le sexe
    const students = dataContext.classesState[classeName];
    Logger.log(`${students.length} élèves à traiter`);

    for (let r = 1; r < allData.length; r++) {
      const row = allData[r];
      const id = row[idIdx];
      if (!id) continue;
      
      const sexeRaw = row[sexeIdx];
      const sexe = normaliserSexe(sexeRaw);
      
      // Trouver l'élève correspondant
      const student = students.find(e => e.ID_ELEVE === id);
      if (student) {
        totalEleves++;
        if (!student.SEXE || student.SEXE === '') {
          student.SEXE = sexe;
          if (sexe === 'F' || sexe === 'M') {
            totalInjected++;
            Logger.log(`✓ ${student.NOM || id}: SEXE = ${sexe}`);
          } else {
            Logger.log(`⚠️ ${student.NOM || id}: Sexe non reconnu: "${sexeRaw}"`);
          }
        }
      }
    }
    
    // Compter après injection
    const nbF = students.filter(e => e.SEXE === 'F').length;
    const nbM = students.filter(e => e.SEXE === 'M').length;
    const nbAutre = students.length - nbF - nbM;
    Logger.log(`Résultat ${classeName}: ${nbF}F, ${nbM}M, ${nbAutre} non définis`);
  });

  Logger.log(`\n=== BILAN INJECTION ===`);
  Logger.log(`Total élèves: ${totalEleves}`);
  Logger.log(`Sexe injecté: ${totalInjected}`);
  
  return totalInjected;
}

/**
 * Normaliser les valeurs de sexe
 */
function normaliserSexe(raw) {
  if (!raw) return '';
  
  const val = String(raw).trim().toUpperCase();
  
  // Féminin
  if (val === 'F' || val === 'FILLE' || val === 'FEMME' || val === 'FEM') {
    return 'F';
  }
  
  // Masculin
  if (val === 'M' || val === 'H' || val === 'G' || 
      val === 'GARCON' || val === 'GARÇON' || 
      val === 'HOMME' || val === 'MASC' || val === 'MASCULIN') {
    return 'M';
  }
  
  // Cas spéciaux
  if (val === '1') return 'M'; // Parfois 1=M, 2=F
  if (val === '2') return 'F';
  
  // Si ça commence par F ou M
  if (val.charAt(0) === 'F') return 'F';
  if (val.charAt(0) === 'M' || val.charAt(0) === 'G' || val.charAt(0) === 'H') return 'M';
  
  return '';
}

/**
 * Version corrigée de la fonction d'optimisation parité agressive
 */
function lancerOptimisationPariteAgressive_Corrigee() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log("\n=== LANCEMENT PARITÉ AGRESSIVE AVEC CORRECTION SEXE ===");
    
    // 1. D'abord faire le diagnostic
    diagnosticColonneSexe();
    
    // 2. Préparer les données
    const config = getConfig();
    config.V2_POIDS_EQUILIBRE = {
      COM: 0.05,
      TRA: 0.05,
      PART: 0.05,
      PARITE: 0.85
    };
    
    const dataContext = V2_Ameliore_PreparerDonnees(config);
    
    // 3. Forcer l'injection du sexe
    const nbInjected = injecterSexeForcee(dataContext);
    Logger.log(`\n✅ ${nbInjected} valeurs SEXE injectées`);
    
    // 4. Vérifier l'état après injection
    Logger.log("\n=== VÉRIFICATION APRÈS INJECTION ===");
    Object.keys(dataContext.classesState).forEach(classe => {
      const eleves = dataContext.classesState[classe];
      const nbF = eleves.filter(e => e.SEXE === 'F').length;
      const nbM = eleves.filter(e => e.SEXE === 'M').length;
      Logger.log(`${classe}: ${nbF}F, ${nbM}M sur ${eleves.length} élèves`);
    });
    
    // 5. Maintenant lancer l'optimisation parité
    if (nbInjected > 0) {
      Logger.log("\n=== LANCEMENT OPTIMISATION PARITÉ ===");
      
      // Utiliser la version agressive existante
      const resultat = V2_Ameliore_OptimisationEngine_PariteAgressive();
      
      if (resultat.success) {
        ui.alert('Succès', 
          `Optimisation parité terminée !\n\n` +
          `${nbInjected} valeurs SEXE corrigées\n` +
          `Opérations effectuées: ${resultat.nbSwapsAppliques || 0}`,
          ui.ButtonSet.OK
        );
      }
      
      return resultat;
      
    } else {
      ui.alert('Problème détecté',
        'Aucune valeur SEXE n\'a pu être injectée.\n\n' +
        'Vérifiez que :\n' +
        '1. Une colonne SEXE existe dans chaque onglet\n' +
        '2. Les valeurs sont F, M, G ou H\n' +
        '3. Consultez les logs pour plus de détails',
        ui.ButtonSet.OK
      );
      
      return { success: false, error: "Pas de données SEXE" };
    }
    
  } catch (e) {
    Logger.log(`ERREUR: ${e.message}\n${e.stack}`);
    ui.alert('Erreur', e.message, ui.ButtonSet.OK);
    return { success: false, error: e.message };
  }
}

/**
 * Test complet
 */
function testCorrectionSexeParite() {
  Logger.log("=== TEST CORRECTION SEXE + PARITÉ ===");
  
  // 1. Diagnostic seul
  diagnosticColonneSexe();
  
  // 2. Test avec correction
  const resultat = lancerOptimisationPariteAgressive_Corrigee();
  
  if (resultat.success) {
    Logger.log("✅ Test réussi !");
  } else {
    Logger.log("❌ Test échoué: " + resultat.error);
  }
}