/**
 * ==================================================================
 *  PARITY_INTELLIGENT_PATCH.GS  – Patch autonome « Stratégie Deux Coups »
 * ------------------------------------------------------------------
 *  Module autonome pour la correction de parité dans les classes
 *  avec stratégie agressive en deux temps et gestion des contraintes.
 * ------------------------------------------------------------------
 *  CHANGELOG v2.5.3 (16 juin 2025 - CORRECTION)
 *   – FIX : Correction critique du respect des contraintes d'options lors des mouvements
 *   – NEW : Validation finale de toutes les opérations avant application
 *   – NEW : Recherche exhaustive d'élèves mobiles compatibles (pas seulement le premier trouvé)
 *   – NEW : Logs détaillés pour tracer les rejets d'options
 *   – MAINT : Amélioration de la robustesse des vérifications de contraintes
 * 
 *  CHANGELOG v2.5.2 (16 juin 2025)
 *   – NEW : Modification de psv5_strategieDeuxCoupsGenerique pour permettre des mouvements
 *           vers des classes "moins déséquilibrées" si toutes les classes ont un surplus du même sexe.
 *           Introduit cfg.PSV5_DIFF_DELTA_MIN_POUR_TRANSFERT_MEME_SENS.
 *   – NEW : Intégration des fonctions de lancement agressif et de diagnostic avancé.
 *   – MAINT : Affinements des logs et de la robustesse.
 * ==================================================================
 */

// ================================================================
// 0.  CONFIGURATION GLOBALE DU DEBUG MODE POUR CE FICHIER
// ================================================================
let PSV5_DEBUG_MODE = false; 

function psv5_initialiserDebugMode(appConfig) {
    if (appConfig && typeof appConfig.DEBUG_MODE_PARITY_STRATEGY === 'boolean') {
        PSV5_DEBUG_MODE = appConfig.DEBUG_MODE_PARITY_STRATEGY;
    } else if (appConfig && typeof appConfig.DEBUG_MODE_PARITE === 'boolean') {
        PSV5_DEBUG_MODE = appConfig.DEBUG_MODE_PARITE;
    } else if (appConfig && typeof appConfig.DEBUG_MODE === 'boolean') {
        PSV5_DEBUG_MODE = appConfig.DEBUG_MODE;
    } else {
        PSV5_DEBUG_MODE = false; 
    }
}

// ================================================================
// 0.  UTILITAIRES GÉNÉRIQUES (psv5_deltaParite, psv5_isMobile, etc.)
// ================================================================
function psv5_deltaParite(eleveArray) {
  if (!eleveArray || !Array.isArray(eleveArray)) return 0;
  const nbF = eleveArray.filter(e => e && (e.SEXE || "").toString().toUpperCase() === "F").length;
  return (eleveArray.length - nbF) - nbF;
}

function psv5_isMobile(e, cfg) {
  if (!e) return false;
  const statut = String(e.MOBILITE || "").toUpperCase();
  const mobilitesFixes = (Array.isArray(cfg?.V2_MOBILITES_CONSIDEREES_FIXES) ? cfg.V2_MOBILITES_CONSIDEREES_FIXES : ['FIXE', 'SPEC']);
  return !mobilitesFixes.includes(statut);
}

function psv5_optionOK(eleve, classeCible, ctx) {
  if (!eleve) return false;
  if (!eleve.OPT || eleve.OPT === "" || eleve.OPT === "ESP") return true;
  const optionPoolsInternes = ctx?.optionPools || {};
  const pool = optionPoolsInternes[eleve.OPT];
  
  // DEBUG: Logger pour vérifier les contraintes d'options
  if (PSV5_DEBUG_MODE && pool) {
    Logger.log(`  psv5_optionOK: Élève ${eleve.NOM || eleve.ID_ELEVE} (OPT:${eleve.OPT}) vers ${classeCible}. Pool autorisé: [${pool.join(',')}]. Résultat: ${pool.includes(classeCible)}`);
  }
  
  return !pool || pool.includes(classeCible);
}

function psv5_computeIdealEffectif(total, nbClasses, tol) {
  if (nbClasses === 0) return { ideal: 0, min: 0, max: tol };
  const ideal = Math.floor(total / nbClasses);
  return { ideal, min: ideal, max: ideal + tol };
}

function psv5_findMobile(classeEleves, sexe, cfg) {
  if (!classeEleves || !Array.isArray(classeEleves)) return undefined;
  return classeEleves.find(e => e && (e.SEXE || "").toUpperCase() === sexe && psv5_isMobile(e, cfg));
}

function psv5_findAllMobiles(classeEleves, sexe, cfg) {
  if (!classeEleves || !Array.isArray(classeEleves)) return [];
  return classeEleves.filter(e => e && (e.SEXE || "").toUpperCase() === sexe && psv5_isMobile(e, cfg));
}

function psv5_prepareCfg(userCfg) {
  const cfg = JSON.parse(JSON.stringify(userCfg || {}));
  
  cfg.PSV5_PARITY_TOLERANCE = cfg.PSV5_PARITY_TOLERANCE ?? (cfg.PARITY_TOLERANCE ?? 2);
  cfg.PSV5_SEUIL_SURPLUS_POSITIF_URGENT = cfg.PSV5_SEUIL_SURPLUS_POSITIF_URGENT ?? 5;
  cfg.PSV5_SEUIL_SURPLUS_NEGATIF_URGENT = cfg.PSV5_SEUIL_SURPLUS_NEGATIF_URGENT ?? -5;
  cfg.PSV5_EFFECTIF_MAX_CIBLE_INITIAL = cfg.PSV5_EFFECTIF_MAX_CIBLE_INITIAL || (cfg.EFFECTIF_MAX || 26);
  cfg.PSV5_FLEXIBILITE_EFFECTIF_MAX = cfg.PSV5_FLEXIBILITE_EFFECTIF_MAX ?? 1;
  cfg.PSV5_EFFECTIF_MAX_STRICT = cfg.PSV5_EFFECTIF_MAX_STRICT || (cfg.PSV5_EFFECTIF_MAX_CIBLE_INITIAL + cfg.PSV5_FLEXIBILITE_EFFECTIF_MAX);
  cfg.PSV5_POTENTIEL_CORRECTION_FACTOR = cfg.PSV5_POTENTIEL_CORRECTION_FACTOR || 5.0;
  cfg.PSV5_MAX_ITER_STRATEGIE = cfg.PSV5_MAX_ITER_STRATEGIE || 5;
  // NOUVEAU PARAMETRE : Différence minimale de delta pour transfert entre classes de même surplus
  cfg.PSV5_DIFF_DELTA_MIN_POUR_TRANSFERT_MEME_SENS = cfg.PSV5_DIFF_DELTA_MIN_POUR_TRANSFERT_MEME_SENS ?? 2;


  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const testSuffix = cfg.TEST_SUFFIX || "TEST";
    const sheets = ss.getSheets().filter(s => {
        const name = s.getName();
        return name && typeof name.endsWith === 'function' && name.endsWith(testSuffix);
    });

    if (sheets.length > 0) {
        let totalF = 0, totalM = 0;
        sheets.forEach(sh => {
            const lastRow = sh.getLastRow();
            if (lastRow > 1) {
                const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
                let sexeColNum = 5; 
                const sexeColIdx = headers.findIndex(h => String(h||"").toUpperCase() === "SEXE");
                if (sexeColIdx !== -1) sexeColNum = sexeColIdx + 1;
                else Logger.log(`WARN psv5_prepareCfg: Colonne SEXE non trouvée explicitement dans ${sh.getName()}, utilisation de la colonne ${sexeColNum}`);

                const sexColumnValues = sh.getRange(2, sexeColNum, lastRow - 1, 1).getValues();
                sexColumnValues.forEach(rowValue => {
                    const sexe = String(rowValue[0] || "").toUpperCase();
                    if (sexe === "F") totalF++; else if (sexe === "M") totalM++;
                });
            }
        });
        if ((totalF + totalM) > 0 && (totalF / (totalF + totalM)) < 0.47) {
            cfg.PSV5_PARITY_TOLERANCE = Math.min(cfg.PSV5_PARITY_TOLERANCE, 1);
            if (PSV5_DEBUG_MODE) Logger.log(`psv5_prepareCfg: Ajustement PARITY_TOLERANCE à ${cfg.PSV5_PARITY_TOLERANCE} (déficit global filles).`);
        }
        const effectifs = sheets.map(sh => Math.max(0, sh.getLastRow() - 1));
        const maxEffectifExistant = effectifs.length > 0 ? Math.max(...effectifs) : cfg.PSV5_EFFECTIF_MAX_CIBLE_INITIAL;
        cfg.PSV5_EFFECTIF_MAX_CIBLE_INITIAL = cfg.PSV5_EFFECTIF_MAX_CIBLE_INITIAL || maxEffectifExistant;
        cfg.PSV5_EFFECTIF_MAX_STRICT = cfg.PSV5_EFFECTIF_MAX_STRICT || (cfg.PSV5_EFFECTIF_MAX_CIBLE_INITIAL + cfg.PSV5_FLEXIBILITE_EFFECTIF_MAX);
    } else {
        if (PSV5_DEBUG_MODE) Logger.log("WARN: psv5_prepareCfg - Aucune feuille TEST. Utilisation des valeurs par défaut/existantes pour effectifs.");
    }
  } catch (e) {
      if (PSV5_DEBUG_MODE) Logger.log(`ERREUR dans psv5_prepareCfg: ${e.message}.`);
  }
  if (PSV5_DEBUG_MODE) {
      Logger.log(`psv5_prepareCfg FINAL: TolParité=${cfg.PSV5_PARITY_TOLERANCE}, SeuilsUrgence=P:${cfg.PSV5_SEUIL_SURPLUS_POSITIF_URGENT}/N:${cfg.PSV5_SEUIL_SURPLUS_NEGATIF_URGENT}, DiffDeltaMinMemeSens=${cfg.PSV5_DIFF_DELTA_MIN_POUR_TRANSFERT_MEME_SENS}, EffCibleInit=${cfg.PSV5_EFFECTIF_MAX_CIBLE_INITIAL}, FlexEff=${cfg.PSV5_FLEXIBILITE_EFFECTIF_MAX}, StrictMaxEff=${cfg.PSV5_EFFECTIF_MAX_STRICT}, PotentielFactor=${cfg.PSV5_POTENTIEL_CORRECTION_FACTOR}, MaxIterStrat=${cfg.PSV5_MAX_ITER_STRATEGIE}`);
  }
  return cfg;
}

// ================================================================
// 0.X VALIDATION DES OPÉRATIONS
// ================================================================
function psv5_validerOperations(operations, ctx, cfg) {
  const operationsValides = [];
  let nbRejetees = 0;
  
  for (const op of operations) {
    let valide = true;
    let raisonRejet = "";
    
    if (op.type === "MOVE") {
      // Vérifier que l'élève peut aller dans la classe cible
      if (!psv5_optionOK(op.eleveA, op.classeB, ctx)) {
        valide = false;
        raisonRejet = `Option ${op.eleveA.OPT || 'aucune'} incompatible avec classe ${op.classeB}`;
      }
    } else if (op.type === "SWAP") {
      // Vérifier les deux sens du swap
      if (!psv5_optionOK(op.eleveA, op.classeB, ctx)) {
        valide = false;
        raisonRejet = `Option ${op.eleveA.OPT || 'aucune'} incompatible avec classe ${op.classeB}`;
      } else if (!psv5_optionOK(op.eleveB, op.classeA, ctx)) {
        valide = false;
        raisonRejet = `Option ${op.eleveB.OPT || 'aucune'} incompatible avec classe ${op.classeA}`;
      }
    }
    
    if (valide) {
      operationsValides.push(op);
    } else {
      nbRejetees++;
      if (PSV5_DEBUG_MODE) {
        Logger.log(`⚠️ OPÉRATION REJETÉE À LA VALIDATION: ${op.type} ${op.eleveA.NOM || op.eleveA.ID_ELEVE}. Raison: ${raisonRejet}`);
      }
    }
  }
  
  if (nbRejetees > 0) {
    Logger.log(`VALIDATION FINALE: ${nbRejetees} opérations rejetées sur ${operations.length} pour violation de contraintes d'options.`);
  }
  
  return operationsValides;
}

// ================================================================
// 1.  MOTEUR DE CORRECTION PARITÉ PRINCIPAL
// ================================================================
function psv5_nirvanaV2_CorrectionPariteINTELLIGENTE(ctx, userCfg) {
  let operationsGlobales = [];
  if (!ctx?.classesState || Object.keys(ctx.classesState).length === 0) {
    if (PSV5_DEBUG_MODE) Logger.log("psv5_nirvanaV2_CorrectionPariteINTELLIGENTE: dataContext invalide ou aucune classe. Arrêt.");
    return operationsGlobales;
  }
  
  const cfg = psv5_prepareCfg(userCfg);
  
  const workingClasses = JSON.parse(JSON.stringify(Object.entries(ctx.classesState)
    .map(([nom, eleves]) => ({ 
        nom, 
        eleves: eleves || [], 
        deltaParite: psv5_deltaParite(eleves || []), 
        effectif: (eleves || []).length, 
        scoreUrgence: 0 
    }))));

  workingClasses.forEach(c => { c.deltaParite = psv5_deltaParite(c.eleves); c.effectif = c.eleves.length; });

  const totalElevesInitial = workingClasses.reduce((s, c) => s + c.effectif, 0);
  if (totalElevesInitial === 0) {
    if (PSV5_DEBUG_MODE) Logger.log("psv5_nirvanaV2_CorrectionPariteINTELLIGENTE: Aucun élève. Arrêt.");
    return operationsGlobales;
  }
  
  if (PSV5_DEBUG_MODE) Logger.log("Lancement de psv5_strategieDeuxCoupsGenerique...");
  const opsStrategieAgressive = psv5_strategieDeuxCoupsGenerique(workingClasses, ctx, cfg);
  operationsGlobales.push(...opsStrategieAgressive);

  if (PSV5_DEBUG_MODE) Logger.log("Lancement de psv5_reduceOverflowFinal après stratégie agressive...");
  const opsReduceOverflow = []; 
  psv5_reduceOverflowFinal(workingClasses, cfg.PSV5_EFFECTIF_MAX_STRICT, cfg, ctx, opsReduceOverflow);
  operationsGlobales.push(...opsReduceOverflow);
  
  // Dédoublonnage : on garde la dernière affectation pour chaque élève
  const opsUniques = [];
  const derniereOperationPourEleve = new Map();

  for (let idx = 0; idx < operationsGlobales.length; idx++) {
    const op = operationsGlobales[idx];
    if (op.eleveA?.ID_ELEVE) derniereOperationPourEleve.set(op.eleveA.ID_ELEVE, idx);
    if (op.eleveB?.ID_ELEVE) derniereOperationPourEleve.set(op.eleveB.ID_ELEVE, idx);
  }
  
  const dejaTraitePourOpFinale = new Set();
  for (let i = operationsGlobales.length - 1; i >= 0; i--) {
      const op = operationsGlobales[i];
      const idA = op.eleveA?.ID_ELEVE;
      const idB = op.eleveB?.ID_ELEVE;
      let prendreCetteOp = false;

      if (op.type === "MOVE" && idA && !dejaTraitePourOpFinale.has(idA)) {
          prendreCetteOp = true;
      } else if (op.type === "SWAP" && idA && idB && !dejaTraitePourOpFinale.has(idA) && !dejaTraitePourOpFinale.has(idB)) {
          prendreCetteOp = true;
      }

      if (prendreCetteOp) {
          opsUniques.unshift(op);
          if (idA) dejaTraitePourOpFinale.add(idA);
          if (idB) dejaTraitePourOpFinale.add(idB);
      }
  }

  if (PSV5_DEBUG_MODE) Logger.log(`psv5_nirvanaV2_CorrectionPariteINTELLIGENTE Dédoublonnage: ${operationsGlobales.length} ops brutes → ${opsUniques.length} ops uniques.`);
  
  // Validation finale pour s'assurer qu'aucune opération ne viole les contraintes d'options
  const opsValidees = psv5_validerOperations(opsUniques, ctx, cfg);
  if (opsValidees.length < opsUniques.length) {
    Logger.log(`⚠️ VALIDATION FINALE: ${opsUniques.length - opsValidees.length} opérations rejetées pour violation de contraintes.`);
  }
  
  return opsValidees;
}

// ================================================================
// 1.X STRATÉGIE "DEUX COUPS" GÉNÉRIQUE avec Logique de Cible Modifiée
// ================================================================
function psv5_strategieDeuxCoupsGenerique(workingClasses, ctx, cfg) {
  if (PSV5_DEBUG_MODE && ctx?.optionPools) {
    Logger.log("=== POOLS D'OPTIONS CHARGÉS ===");
    Object.entries(ctx.optionPools).forEach(([opt, classes]) => {
      Logger.log(`  ${opt}: [${classes.join(', ')}]`);
    });
  }
  
  let operationsEffectuees = [];
  let aChangeQuelqueChoseGlobalement = true;
  let iterationPrincipale = 0;
  const MAX_ITERATIONS_PRINCIPALES = cfg.PSV5_MAX_ITER_STRATEGIE || 5;

  const EFFECTIF_MAX_POUR_MOVE_SIMPLE = cfg.PSV5_EFFECTIF_MAX_CIBLE_INITIAL + cfg.PSV5_FLEXIBILITE_EFFECTIF_MAX;
  const EFFECTIF_MAX_POUR_SWAP_FACIL = cfg.PSV5_EFFECTIF_MAX_CIBLE_INITIAL;
  const DIFF_DELTA_MIN_TRANSFERT = cfg.PSV5_DIFF_DELTA_MIN_POUR_TRANSFERT_MEME_SENS || 1;


  while (aChangeQuelqueChoseGlobalement && iterationPrincipale < MAX_ITERATIONS_PRINCIPALES) {
    iterationPrincipale++;
    aChangeQuelqueChoseGlobalement = false;
    if (PSV5_DEBUG_MODE) Logger.log(`--- psv5_strategieDeuxCoupsGenerique - Itération Principale ${iterationPrincipale} ---`);

    workingClasses.forEach(c => {
      c.deltaParite = psv5_deltaParite(c.eleves);
      c.effectif = c.eleves.length;
      const sexeQuiDevraitPartir = c.deltaParite > cfg.PSV5_PARITY_TOLERANCE ? "M" : (c.deltaParite < -cfg.PSV5_PARITY_TOLERANCE ? "F" : null);
      let potentiel = 0;
      if (sexeQuiDevraitPartir) {
        potentiel = c.eleves.filter(e => psv5_isMobile(e, cfg) && e.SEXE === sexeQuiDevraitPartir).length;
      }
      c.scoreUrgence = Math.abs(c.deltaParite) * (1 + potentiel / (cfg.PSV5_POTENTIEL_CORRECTION_FACTOR || 5.0) );
    });

    // --- COUP 1 : Réduire Surplus Positifs (M-F) Urgents ---
    if (PSV5_DEBUG_MODE) Logger.log(`  COUP 1 (It ${iterationPrincipale}): Réduction des surplus Positifs (M > F)`);
    workingClasses.sort((a, b) => b.scoreUrgence - a.scoreUrgence); 

    for (const classeSource of workingClasses) {
      if (classeSource.deltaParite <= cfg.PSV5_PARITY_TOLERANCE) continue;
      if (classeSource.deltaParite <= (cfg.PSV5_SEUIL_SURPLUS_POSITIF_URGENT || 3) && classeSource.scoreUrgence < ((cfg.PSV5_SEUIL_SURPLUS_POSITIF_URGENT || 3) * 1.1) ) continue; 

      const sexeEnSurplus = "M";
      const aucuneClasseAvecDeficitGarcons = workingClasses.every(cl => cl.deltaParite >= -cfg.PSV5_PARITY_TOLERANCE || cl.nom === classeSource.nom);
      
      const ciblesPotentielles = workingClasses
        .filter(c => {
          if (c.nom === classeSource.nom) return false;
          const peutAccueillirPhysiquement = (c.effectif < EFFECTIF_MAX_POUR_MOVE_SIMPLE) || 
                                           (c.effectif <= EFFECTIF_MAX_POUR_SWAP_FACIL && psv5_findMobile(c.eleves, "F", cfg));
          if (!peutAccueillirPhysiquement) return false;

          if (aucuneClasseAvecDeficitGarcons) {
            return c.deltaParite < (classeSource.deltaParite - DIFF_DELTA_MIN_TRANSFERT);
          } else {
            return c.deltaParite < cfg.PSV5_PARITY_TOLERANCE;
          }
        })
        .sort((a, b) => a.deltaParite - b.deltaParite);

      if (PSV5_DEBUG_MODE && ciblesPotentielles.length === 0 && classeSource.deltaParite > cfg.PSV5_PARITY_TOLERANCE) {
        Logger.log(`    COUP 1: Aucune cible trouvée pour ${classeSource.nom} (Δ=${classeSource.deltaParite.toFixed(0)}, ScoreUrg=${classeSource.scoreUrgence.toFixed(1)}). AucuneClasseAvecDeficitGarcons: ${aucuneClasseAvecDeficitGarcons}`);
      }

      for (const classeCible of ciblesPotentielles) {
        if (classeSource.deltaParite <= cfg.PSV5_PARITY_TOLERANCE) break; 
        
        // Essayer TOUS les élèves mobiles du sexe en surplus
        const elevesMobilesSource = psv5_findAllMobiles(classeSource.eleves, sexeEnSurplus, cfg);
        
        if (elevesMobilesSource.length === 0) {
          if (PSV5_DEBUG_MODE) Logger.log(`    Aucun élève ${sexeEnSurplus} mobile dans ${classeSource.nom}`);
          continue;
        }
        
        let eleveMobileSource = null;
        let optionCompatibleTrouvee = false;
        
        // Essayer chaque élève mobile jusqu'à en trouver un compatible
        for (const candidat of elevesMobilesSource) {
          if (psv5_optionOK(candidat, classeCible.nom, ctx)) {
            eleveMobileSource = candidat;
            optionCompatibleTrouvee = true;
            break;
          } else if (PSV5_DEBUG_MODE) {
            Logger.log(`    REJET OPTION: ${candidat.NOM || candidat.ID_ELEVE} (OPT:${candidat.OPT || 'aucune'}) incompatible avec ${classeCible.nom}`);
          }
        }
        
        if (!optionCompatibleTrouvee) {
          if (PSV5_DEBUG_MODE) Logger.log(`    Aucun élève ${sexeEnSurplus} mobile avec option compatible trouvé dans ${classeSource.nom} pour ${classeCible.nom}`);
          continue;
        }

        let opEffectueeCeTour = false;
        const deltaSourceAvant = classeSource.deltaParite;
        const deltaCibleAvant = classeCible.deltaParite;

        if (classeCible.effectif < EFFECTIF_MAX_POUR_MOVE_SIMPLE) {
          classeSource.eleves.splice(classeSource.eleves.indexOf(eleveMobileSource), 1);
          classeCible.eleves.push(eleveMobileSource);
          operationsEffectuees.push({ type: "MOVE", eleveA: eleveMobileSource, classeA: classeSource.nom, classeB: classeCible.nom, motif: `StratC1-M ${classeSource.nom}(${deltaSourceAvant.toFixed(0)})→${classeCible.nom}(${deltaCibleAvant.toFixed(0)})`});
          if (PSV5_DEBUG_MODE) Logger.log(`    MOVE ${sexeEnSurplus}: ${eleveMobileSource.ID_ELEVE} de ${classeSource.nom}(Δ${deltaSourceAvant.toFixed(0)}) vers ${classeCible.nom}(Δ${deltaCibleAvant.toFixed(0)}), EffCible:${classeCible.effectif+1}`);
          opEffectueeCeTour = true;
        } else if (classeCible.effectif <= EFFECTIF_MAX_POUR_SWAP_FACIL) { 
          const sexeOpposeDansCible = "F";
          // Chercher TOUS les candidats possibles pour le swap
          const candidatsCible = psv5_findAllMobiles(classeCible.eleves, sexeOpposeDansCible, cfg);
          let eleveAEchangerDansCible = null;
          
          for (const candidatCible of candidatsCible) {
            // Vérifier que le swap est valide dans les DEUX sens
            if (psv5_optionOK(candidatCible, classeSource.nom, ctx) && 
                psv5_optionOK(eleveMobileSource, classeCible.nom, ctx)) {
              eleveAEchangerDansCible = candidatCible;
              if (PSV5_DEBUG_MODE) {
                Logger.log(`    SWAP validé: ${eleveMobileSource.NOM}(${eleveMobileSource.OPT||'aucune'}) ↔ ${candidatCible.NOM}(${candidatCible.OPT||'aucune'})`);
              }
              break;
            } else if (PSV5_DEBUG_MODE) {
              Logger.log(`    SWAP rejeté: contraintes d'options incompatibles entre ${eleveMobileSource.NOM} et ${candidatCible.NOM}`);
            }
          }
          
          if (eleveAEchangerDansCible) {
            classeSource.eleves.splice(classeSource.eleves.indexOf(eleveMobileSource), 1);
            classeCible.eleves.splice(classeCible.eleves.indexOf(eleveAEchangerDansCible), 1);
            classeSource.eleves.push(eleveAEchangerDansCible);
            classeCible.eleves.push(eleveMobileSource);
            operationsEffectuees.push({ type: "SWAP", eleveA: eleveMobileSource, classeA: classeSource.nom, eleveB: eleveAEchangerDansCible, classeB: classeCible.nom, motif: `StratC1-SF ${classeSource.nom}↔${classeCible.nom}` });
            if (PSV5_DEBUG_MODE) Logger.log(`    SWAP-FACIL ${sexeEnSurplus}: ${eleveMobileSource.ID_ELEVE}(${classeSource.nom}) ↔ ${eleveAEchangerDansCible.ID_ELEVE}(${classeCible.nom})`);
            opEffectueeCeTour = true;
          }
        }
        if (opEffectueeCeTour) {
          [classeSource, classeCible].forEach(c => {
              c.deltaParite = psv5_deltaParite(c.eleves); c.effectif = c.eleves.length;
              const sexeQuiDevraitPartir = c.deltaParite > cfg.PSV5_PARITY_TOLERANCE ? "M" : (c.deltaParite < -cfg.PSV5_PARITY_TOLERANCE ? "F" : null);
              let potentiel = 0;
              if (sexeQuiDevraitPartir) potentiel = c.eleves.filter(e => psv5_isMobile(e, cfg) && e.SEXE === sexeQuiDevraitPartir).length;
              c.scoreUrgence = Math.abs(c.deltaParite) * (1 + potentiel / (cfg.PSV5_POTENTIEL_CORRECTION_FACTOR || 5.0) );
          });
          aChangeQuelqueChoseGlobalement = true;
        }
      }
    }

    // --- COUP 2 : Réduire Surplus Négatifs (F > M) ---
    if (PSV5_DEBUG_MODE) Logger.log(`  COUP 2 (It ${iterationPrincipale}): Réduction des surplus Négatifs (F > M)`);
    workingClasses.sort((a, b) => b.scoreUrgence - a.scoreUrgence); 

    for (const classeSource of workingClasses) {
      if (classeSource.deltaParite >= -cfg.PSV5_PARITY_TOLERANCE) continue; 
      if (classeSource.deltaParite >= (cfg.PSV5_SEUIL_SURPLUS_NEGATIF_URGENT || -3) && classeSource.scoreUrgence < (Math.abs(cfg.PSV5_SEUIL_SURPLUS_NEGATIF_URGENT) * 1.1) ) continue;

      const sexeEnSurplus = "F";
      const aucuneClasseAvecDeficitFilles = workingClasses.every(cl => cl.deltaParite <= cfg.PSV5_PARITY_TOLERANCE || cl.nom === classeSource.nom);

      const ciblesPotentielles = workingClasses
        .filter(c => {
          if (c.nom === classeSource.nom) return false;
          const peutAccueillirPhysiquement = (c.effectif < EFFECTIF_MAX_POUR_MOVE_SIMPLE) ||
                                           (c.effectif <= EFFECTIF_MAX_POUR_SWAP_FACIL && psv5_findMobile(c.eleves, "M", cfg));
          if (!peutAccueillirPhysiquement) return false;

          if (aucuneClasseAvecDeficitFilles) {
            return c.deltaParite > (classeSource.deltaParite + DIFF_DELTA_MIN_TRANSFERT);
          } else {
            return c.deltaParite > -cfg.PSV5_PARITY_TOLERANCE;
          }
        })
        .sort((a, b) => b.deltaParite - a.deltaParite);
      
      if (PSV5_DEBUG_MODE && ciblesPotentielles.length === 0 && classeSource.deltaParite < -cfg.PSV5_PARITY_TOLERANCE) {
        Logger.log(`    COUP 2: Aucune cible trouvée pour ${classeSource.nom} (Δ=${classeSource.deltaParite.toFixed(0)}, ScoreUrg=${classeSource.scoreUrgence.toFixed(1)}). AucuneClasseAvecDeficitFilles: ${aucuneClasseAvecDeficitFilles}`);
      }

      for (const classeCible of ciblesPotentielles) {
        if (classeSource.deltaParite >= -cfg.PSV5_PARITY_TOLERANCE) break;
        
        // Essayer TOUS les élèves mobiles du sexe en surplus
        const elevesMobilesSource = psv5_findAllMobiles(classeSource.eleves, sexeEnSurplus, cfg);
        
        if (elevesMobilesSource.length === 0) {
          if (PSV5_DEBUG_MODE) Logger.log(`    Aucun élève ${sexeEnSurplus} mobile dans ${classeSource.nom}`);
          continue;
        }
        
        let eleveMobileSource = null;
        let optionCompatibleTrouvee = false;
        
        // Essayer chaque élève mobile jusqu'à en trouver un compatible
        for (const candidat of elevesMobilesSource) {
          if (psv5_optionOK(candidat, classeCible.nom, ctx)) {
            eleveMobileSource = candidat;
            optionCompatibleTrouvee = true;
            break;
          } else if (PSV5_DEBUG_MODE) {
            Logger.log(`    REJET OPTION: ${candidat.NOM || candidat.ID_ELEVE} (OPT:${candidat.OPT || 'aucune'}) incompatible avec ${classeCible.nom}`);
          }
        }
        
        if (!optionCompatibleTrouvee) {
          if (PSV5_DEBUG_MODE) Logger.log(`    Aucun élève ${sexeEnSurplus} mobile avec option compatible trouvé dans ${classeSource.nom} pour ${classeCible.nom}`);
          continue;
        }

        let opEffectueeCeTour = false;
        const deltaSourceAvant = classeSource.deltaParite;
        const deltaCibleAvant = classeCible.deltaParite;

        if (classeCible.effectif < EFFECTIF_MAX_POUR_MOVE_SIMPLE) {
          classeSource.eleves.splice(classeSource.eleves.indexOf(eleveMobileSource), 1);
          classeCible.eleves.push(eleveMobileSource);
          operationsEffectuees.push({ type: "MOVE", eleveA: eleveMobileSource, classeA: classeSource.nom, classeB: classeCible.nom, motif: `StratC2-F ${classeSource.nom}(${deltaSourceAvant.toFixed(0)})→${classeCible.nom}(${deltaCibleAvant.toFixed(0)})`});
          if (PSV5_DEBUG_MODE) Logger.log(`    MOVE ${sexeEnSurplus}: ${eleveMobileSource.ID_ELEVE} de ${classeSource.nom}(Δ${deltaSourceAvant.toFixed(0)}) vers ${classeCible.nom}(Δ${deltaCibleAvant.toFixed(0)}), EffCible:${classeCible.effectif+1}`);
          opEffectueeCeTour = true;
        } else if (classeCible.effectif <= EFFECTIF_MAX_POUR_SWAP_FACIL) {
          const sexeOpposeDansCible = "M";
          // Chercher TOUS les candidats possibles pour le swap
          const candidatsCible = psv5_findAllMobiles(classeCible.eleves, sexeOpposeDansCible, cfg);
          let eleveAEchangerDansCible = null;
          
          for (const candidatCible of candidatsCible) {
            // Vérifier que le swap est valide dans les DEUX sens
            if (psv5_optionOK(candidatCible, classeSource.nom, ctx) && 
                psv5_optionOK(eleveMobileSource, classeCible.nom, ctx)) {
              eleveAEchangerDansCible = candidatCible;
              if (PSV5_DEBUG_MODE) {
                Logger.log(`    SWAP validé: ${eleveMobileSource.NOM}(${eleveMobileSource.OPT||'aucune'}) ↔ ${candidatCible.NOM}(${candidatCible.OPT||'aucune'})`);
              }
              break;
            } else if (PSV5_DEBUG_MODE) {
              Logger.log(`    SWAP rejeté: contraintes d'options incompatibles entre ${eleveMobileSource.NOM} et ${candidatCible.NOM}`);
            }
          }
          
          if (eleveAEchangerDansCible) {
            classeSource.eleves.splice(classeSource.eleves.indexOf(eleveMobileSource), 1);
            classeCible.eleves.splice(classeCible.eleves.indexOf(eleveAEchangerDansCible), 1);
            classeSource.eleves.push(eleveAEchangerDansCible);
            classeCible.eleves.push(eleveMobileSource);
            operationsEffectuees.push({ type: "SWAP", eleveA: eleveMobileSource, classeA: classeSource.nom, eleveB: eleveAEchangerDansCible, classeB: classeCible.nom, motif: `StratC2-SF ${classeSource.nom}↔${classeCible.nom}` });
            if (PSV5_DEBUG_MODE) Logger.log(`    SWAP-FACIL ${sexeEnSurplus}: ${eleveMobileSource.ID_ELEVE}(${classeSource.nom}) ↔ ${eleveAEchangerDansCible.ID_ELEVE}(${classeCible.nom})`);
            opEffectueeCeTour = true;
          }
        }
        if (opEffectueeCeTour) {
          [classeSource, classeCible].forEach(c => {
              c.deltaParite = psv5_deltaParite(c.eleves); c.effectif = c.eleves.length;
              const sexeQuiDevraitPartir = c.deltaParite > cfg.PSV5_PARITY_TOLERANCE ? "M" : (c.deltaParite < -cfg.PSV5_PARITY_TOLERANCE ? "F" : null);
              let potentiel = 0;
              if (sexeQuiDevraitPartir) potentiel = c.eleves.filter(e => psv5_isMobile(e, cfg) && e.SEXE === sexeQuiDevraitPartir).length;
              c.scoreUrgence = Math.abs(c.deltaParite) * (1 + potentiel / (cfg.PSV5_POTENTIEL_CORRECTION_FACTOR || 5.0) );
          });
          aChangeQuelqueChoseGlobalement = true;
        }
      }
    }
     if (!aChangeQuelqueChoseGlobalement && PSV5_DEBUG_MODE && iterationPrincipale > 0) Logger.log(`psv5_strategieDeuxCoupsGenerique: Aucune action lors de l'itération ${iterationPrincipale}. Arrêt de la stratégie.`);
  } 
  if (PSV5_DEBUG_MODE && iterationPrincipale === MAX_ITERATIONS_PRINCIPALES) Logger.log(`psv5_strategieDeuxCoupsGenerique: MAX_ITERATIONS_PRINCIPALES (${MAX_ITERATIONS_PRINCIPALES}) atteinte.`);
  
  return operationsEffectuees;
}

// ----------------------------------------------------------------
// 1.D  RÉDUCTION OVERFLOW FINALE (plus stricte)
// ----------------------------------------------------------------
function psv5_reduceOverflowFinal(workingClasses, effectifMaxStrict, cfg, ctx, swapsArr) {
  if (PSV5_DEBUG_MODE) Logger.log(`—— psv5_reduceOverflowFinal (cible max/classe stricte: ${effectifMaxStrict}) ——`);
  let moves = 0;
  let continueReducing = true;
  let safetyBreak = workingClasses.reduce((sum,c) => sum + (c.eleves?.length || 0) ,0) + 10; 

  while (continueReducing && safetyBreak-- > 0) {
    continueReducing = false;
    workingClasses.sort((a, b) => (b.eleves?.length || 0) - (a.eleves?.length || 0)); 

    for (const classeSource of workingClasses) {
      if ((classeSource.eleves?.length || 0) <= effectifMaxStrict) continue;

      const deltaSource = psv5_deltaParite(classeSource.eleves);
      let sexeADeplacerPrioritaire = null;
      if (deltaSource > cfg.PSV5_PARITY_TOLERANCE) sexeADeplacerPrioritaire = "M";
      else if (deltaSource < -cfg.PSV5_PARITY_TOLERANCE) sexeADeplacerPrioritaire = "F";

      let eleveADeplacer = null;
      if (sexeADeplacerPrioritaire) {
        eleveADeplacer = psv5_findMobile(classeSource.eleves, sexeADeplacerPrioritaire, cfg);
      }
      if (!eleveADeplacer) eleveADeplacer = psv5_findMobile(classeSource.eleves, "M", cfg) || psv5_findMobile(classeSource.eleves, "F", cfg);
      
      if (!eleveADeplacer) {
        if (PSV5_DEBUG_MODE) Logger.log(`  ReduceOverflow: ${classeSource.nom} en surplus mais aucun élève mobile trouvé.`);
        continue; 
      }

      const destinationsPossibles = workingClasses
        .filter(cD => cD.nom !== classeSource.nom && (cD.eleves?.length || 0) < effectifMaxStrict && psv5_optionOK(eleveADeplacer, cD.nom, ctx))
        .sort((a, b) => (a.eleves?.length || 0) - (b.eleves?.length || 0)); 

      if (destinationsPossibles.length > 0) {
        const classeDest = destinationsPossibles[0];
        
        swapsArr.push({ type: "MOVE", eleveA: eleveADeplacer, classeA: classeSource.nom, classeB: classeDest.nom, motif: "ReduceOverflowFinal Move" });
        classeSource.eleves.splice(classeSource.eleves.indexOf(eleveADeplacer), 1);
        classeDest.eleves.push(eleveADeplacer);
        
        classeSource.effectif = classeSource.eleves.length; classeSource.deltaParite = psv5_deltaParite(classeSource.eleves);
        classeDest.effectif = classeDest.eleves.length; classeDest.deltaParite = psv5_deltaParite(classeDest.eleves);
        moves++;
        if (PSV5_DEBUG_MODE) Logger.log(`  ReduceOverflowFinal - MOVE: ${eleveADeplacer.ID_ELEVE}(${classeSource.nom}) → ${classeDest.nom}`);
        continueReducing = true;
        break; 
      } else {
         if (PSV5_DEBUG_MODE) Logger.log(`  ReduceOverflow: ${classeSource.nom} en surplus, ${eleveADeplacer.ID_ELEVE} mobile, mais aucune destination trouvée.`);
      }
    }
  }
  if (safetyBreak <= 0 && PSV5_DEBUG_MODE) Logger.log("WARN: psv5_reduceOverflowFinal - Safety break atteint.");
  if (PSV5_DEBUG_MODE) Logger.log(`psv5_reduceOverflowFinal - Moves effectués : ${moves}`);
}

// ================================================================
// 2.  APPLICATION DES SWAPS
// ================================================================
function psv5_AppliquerSwapsSafeEtLog(swapsAApliquer, ctx, cfg) {
  if (PSV5_DEBUG_MODE) Logger.log(`psv5_AppliquerSwapsSafeEtLog: Tentative d'application de ${swapsAApliquer.length} opérations.`);
  if (swapsAApliquer.length === 0) return 0;
  let nbSwapsEffectifs = 0;

  const vraisSwaps = swapsAApliquer.filter(s => s.type === "SWAP" && s.eleveA && s.eleveB && s.eleveA.ID_ELEVE && s.eleveB.ID_ELEVE);
  const movesSimples = swapsAApliquer.filter(s => s.type === "MOVE" && s.eleveA && !s.eleveB && s.eleveA.ID_ELEVE);

  if (vraisSwaps.length > 0) {
    if (typeof V2_Ameliore_AppliquerSwaps === 'function') { 
      if (PSV5_DEBUG_MODE) Logger.log(`  Application de ${vraisSwaps.length} vrais swaps via V2_Ameliore_AppliquerSwaps.`);
      try {
        const swapsFormatesPourV2 = vraisSwaps.map(s => ({
            eleve1: s.eleveA, eleve2: s.eleveB,
            eleve1ID: s.eleveA.ID_ELEVE, eleve2ID: s.eleveB.ID_ELEVE,
            eleve1Nom: s.eleveA.NOM || s.eleveA.ID_ELEVE, eleve2Nom: s.eleveB.NOM || s.eleveB.ID_ELEVE,
            classe1: s.classeA, classe2: s.classeB, 
            oldClasseE1: s.classeA, oldClasseE2: s.classeB,
            newClasseE1: s.classeB, newClasseE2: s.classeA, 
            motif: s.motif || "Parity Intelligent Swap"
        }));
        const appliedCount = V2_Ameliore_AppliquerSwaps(swapsFormatesPourV2, ctx, cfg);
        nbSwapsEffectifs += (typeof appliedCount === 'number' ? appliedCount : swapsFormatesPourV2.length);
      } catch (e) {
        Logger.log(`  ERREUR V2_Ameliore_AppliquerSwaps: ${e.message}\n${e.stack}`);
      }
    } else {
      Logger.log("  WARN: V2_Ameliore_AppliquerSwaps non trouvée. Vrais swaps de ce patch non appliqués par cette méthode.");
    }
  }

  if (movesSimples.length > 0) {
    if (PSV5_DEBUG_MODE) Logger.log(`  Application de ${movesSimples.length} moves simples via psv5_moveStudentRow.`);
    movesSimples.forEach(m => {
      try {
        if (psv5_moveStudentRow(m.eleveA.ID_ELEVE, m.classeA, m.classeB)) {
          nbSwapsEffectifs++;
          if (PSV5_DEBUG_MODE) Logger.log(`    MOVE OK: ${m.eleveA.ID_ELEVE} de ${m.classeA} vers ${m.classeB}`);
        }
      } catch (err) {
        Logger.log(`    ERREUR psv5_moveStudentRow (${m.eleveA.ID_ELEVE}, ${m.classeA}→${m.classeB}): ${err.message}`);
      }
    });
  }
  if (PSV5_DEBUG_MODE) Logger.log(`psv5_AppliquerSwapsSafeEtLog: Total opérations physiquement tentées/réussies: ${nbSwapsEffectifs}`);
  return nbSwapsEffectifs;
}

function psv5_moveStudentRow(idEleve, fromSheetName, toSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fromSh = ss.getSheetByName(fromSheetName);
  const toSh = ss.getSheetByName(toSheetName);
  if (!fromSh) { Logger.log(`ERREUR psv5_moveStudentRow: Feuille source '${fromSheetName}' introuvable.`); return false; }
  if (!toSh) { Logger.log(`ERREUR psv5_moveStudentRow: Feuille destination '${toSheetName}' introuvable.`); return false; }
  const headerRow = 1;
  const headers = fromSh.getRange(headerRow, 1, 1, fromSh.getLastColumn()).getValues()[0];
  const idColIndex = headers.findIndex(h => String(h || "").toUpperCase().replace(/\s+/g, '_') === "ID_ELEVE");
  if (idColIndex === -1) { Logger.log(`ERREUR psv5_moveStudentRow: Colonne 'ID_ELEVE' introuvable dans '${fromSheetName}'. Headers: ${headers.join(',')}`); return false; }
  
  const lastDataRowInSource = fromSh.getLastRow();
  if (lastDataRowInSource <= headerRow) { 
    if (PSV5_DEBUG_MODE) Logger.log(`INFO psv5_moveStudentRow: Feuille source '${fromSheetName}' vide (hors en-tête).`); 
    return false;
  }
  const dataRange = fromSh.getRange(headerRow + 1, 1, lastDataRowInSource - headerRow, fromSh.getLastColumn());
  const values = dataRange.getValues();
  let rowIndexInData = -1; 
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][idColIndex] || "") === String(idEleve)) {
      rowIndexInData = i;
      break;
    }
  }
  if (rowIndexInData === -1) { Logger.log(`ERREUR psv5_moveStudentRow: Élève ID '${idEleve}' introuvable dans '${fromSheetName}'.`); return false; }
  
  const actualSheetRow = rowIndexInData + headerRow + 1;
  const rowValues = fromSh.getRange(actualSheetRow, 1, 1, fromSh.getLastColumn()).getValues()[0];
  
  fromSh.deleteRow(actualSheetRow);
  toSh.appendRow(rowValues);
  return true;
}

// ================================================================
// 3.  DIAGNOSTIC & DEBUG (Points d'entrée UI)
// ================================================================
function psv5_diagnosticPariteGlobal() {
  const ui = SpreadsheetApp.getUi();
  const cfgAppel = getConfig(); 
  psv5_initialiserDebugMode(cfgAppel); 

  try {
    const cfg = psv5_prepareCfg(cfgAppel);
    if (typeof V2_Ameliore_PreparerDonnees_AvecSEXE !== 'function') {
        Logger.log("ERREUR: V2_Ameliore_PreparerDonnees_AvecSEXE non définie.");
        ui.alert("Erreur Configuration", "Fonction de préparation des données manquante (V2_Ameliore_PreparerDonnees_AvecSEXE).", ui.ButtonSet.OK);
        return "Erreur: V2_Ameliore_PreparerDonnees_AvecSEXE non définie.";
    }
    const ctx = V2_Ameliore_PreparerDonnees_AvecSEXE(cfg); 
    if (!ctx || !ctx.classesState) {
        Logger.log("ERREUR: Contexte de données (ctx) ou ctx.classesState non initialisé par la préparation.");
        ui.alert("Erreur Préparation", "Contexte de données invalide après la préparation.", ui.ButtonSet.OK);
        return "Erreur: Contexte de données invalide après préparation.";
    }

    const rep = Object.entries(ctx.classesState).map(([c, elArray]) => {
      const el = Array.isArray(elArray) ? elArray : [];
      const d = psv5_deltaParite(el); 
      const nbF = el.filter(e => e && String(e.SEXE).toUpperCase() === "F").length;
      const effectif = el.length;
      const lock = effectif >= (cfg.PSV5_EFFECTIF_MAX_STRICT) ? "🔒" : "  ";
      return `${lock} ${c.padEnd(10)} Eff:${String(effectif).padStart(2)} | ${String(nbF).padStart(2)}F | ΔParité: ${String(d).padStart(3)}`;
    }).join("\n");
    
    const logOutput = "=== PARITÉ – état global (AVANT correction de ce module) ===\n" + rep;
    Logger.log(logOutput);
    ui.showModalDialog(HtmlService.createHtmlOutput("<pre>" + logOutput.replace(/\n/g, "<br>") + "</pre>").setWidth(480).setHeight(400), "Diagnostic Parité Global");
    return rep;
  } catch (e) {
    Logger.log(`ERREUR psv5_diagnosticPariteGlobal: ${e.message}\n${e.stack}`);
    ui.alert("Erreur Diagnostic", `Erreur: ${e.message}`, ui.ButtonSet.OK);
    return `Erreur: ${e.message}`;
  }
}

function psv5_debugPourquoiPasDeSwaps() {
  const ui = SpreadsheetApp.getUi();
  const cfgAppel = getConfig();
  const oldDebugMode = PSV5_DEBUG_MODE; 
  PSV5_DEBUG_MODE = true; 
  Logger.log("PSV5_DEBUG_MODE forcé à TRUE pour psv5_debugPourquoiPasDeSwaps.");

  try {
    const cfg = psv5_prepareCfg(cfgAppel); 
    if (typeof V2_Ameliore_PreparerDonnees_AvecSEXE !== 'function') { Logger.log("E:Prep Fct"); ui.alert("E:Prep Fct", "", ui.ButtonSet.OK); return []; }
    const ctx = V2_Ameliore_PreparerDonnees_AvecSEXE(cfg);
    if (!ctx || !ctx.classesState) { Logger.log("E:Prep Ctx"); ui.alert("E:Prep Ctx", "", ui.ButtonSet.OK); return []; }

    const ops = psv5_nirvanaV2_CorrectionPariteINTELLIGENTE(ctx, cfg); 
    
    Logger.log(`--- Résultat psv5_debugPourquoiPasDeSwaps ---`);
    Logger.log(`Opérations calculées par psv5_nirvanaV2_CorrectionPariteINTELLIGENTE : ${ops.length}`);
    ops.slice(0, 30).forEach((s, i) => {
      const nomA = s.eleveA?.NOM || s.eleveA?.ID_ELEVE || 'N/A';
      const motif = s.motif || 'N/D';
      let logMsg;
      if (s.type === "SWAP") {
        const nomB = s.eleveB?.NOM || s.eleveB?.ID_ELEVE || 'N/A';
        logMsg = `#${i + 1} ${s.type}: ${nomA} (${s.classeA || 'N/A'}) ↔ ${nomB} (${s.classeB || 'N/A'}) | Motif: ${motif}`;
      } else {
        logMsg = `#${i + 1} ${s.type}: ${nomA} (${s.classeA || 'N/A'}) → ${s.classeB || 'N/A'} | Motif: ${motif}`;
      }
      Logger.log(logMsg);
    });
    if (!ops.length) {
      ui.alert("Debug Swaps", "Aucun swap/move généré. Consultez logs.", ui.ButtonSet.OK);
    } else {
      ui.alert("Debug Swaps", `${ops.length} opérations calculées (non appliquées). Consultez logs.`, ui.ButtonSet.OK);
    }
    return ops;
  } catch (e) { Logger.log(`E debug: ${e.message}\n${e.stack}`); ui.alert("E debug", e.message, ui.ButtonSet.OK); return [];}
  finally { PSV5_DEBUG_MODE = oldDebugMode; }
}

// ================================================================
// 4.  PIPELINE COMPLET (Point d'entrée principal UI pour ce module)
// ================================================================
function psv5_lancerOptimisationCombinee() {
  const ui = SpreadsheetApp.getUi();
  const cfgAppel = getConfig(); 
  psv5_initialiserDebugMode(cfgAppel); 

  const cfg = psv5_prepareCfg(cfgAppel);
  cfg.PATCH_VERSION = "2.5.3"; 
  
  Logger.log(`=== OPTIMISATION COMBINÉE (Parity Intelligent Patch v${cfg.PATCH_VERSION}) – Démarrage ===`);
  SpreadsheetApp.getActiveSpreadsheet().toast("Optimisation Parité Intelligente en cours...", "Statut", -1);

  try {
    if (typeof V2_Ameliore_PreparerDonnees_AvecSEXE !== 'function') {
        Logger.log("ERREUR CRITIQUE: V2_Ameliore_PreparerDonnees_AvecSEXE manquante.");
        ui.alert("Erreur Configuration", "Fonction de préparation des données (V2_Ameliore_PreparerDonnees_AvecSEXE) est manquante.", ui.ButtonSet.OK);
        return { success: false, error: "Préparation des données manquante." };
    }
    const ctx = V2_Ameliore_PreparerDonnees_AvecSEXE(cfg);
    if (!ctx || !ctx.classesState) {
        Logger.log("ERREUR CRITIQUE: Préparation des données n'a pas retourné un contexte valide.");
        ui.alert("Erreur Préparation", "La préparation des données n'a pas fonctionné correctement.", ui.ButtonSet.OK);
        return { success: false, error: "Contexte de données invalide après préparation." };
    }
    if (PSV5_DEBUG_MODE) Logger.log("Préparation des données terminée. Contexte initialisé.");
        
    if (PSV5_DEBUG_MODE) Logger.log("Lancement de psv5_nirvanaV2_CorrectionPariteINTELLIGENTE...");
    const opsParite = psv5_nirvanaV2_CorrectionPariteINTELLIGENTE(ctx, cfg);
    if (PSV5_DEBUG_MODE) Logger.log(`CorrectionPariteINTELLIGENTE a retourné ${opsParite.length} opérations uniques.`);

    let nbApplied = 0;
    if (opsParite.length > 0) {
      if (PSV5_DEBUG_MODE) Logger.log(`Application de ${opsParite.length} opérations de parité...`);
      nbApplied = psv5_AppliquerSwapsSafeEtLog(opsParite, ctx, cfg);
      Logger.log(`✅ ${nbApplied} opérations de parité effectivement appliquées sur les feuilles.`);
      ui.alert("Optimisation Parité Terminée", `${nbApplied} corrections de parité ont été appliquées aux feuilles.\nConsultez les logs pour le détail.`, ui.ButtonSet.OK);
    } else {
      Logger.log("Aucune opération de parité nécessaire ou possible trouvée par le moteur intelligent.");
      ui.alert("Info Parité", "Aucune opération de parité nécessaire ou possible n'a été trouvée.", ui.ButtonSet.OK);
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast("Optimisation Parité Intelligente terminée!", "Statut", 5);
    Logger.log(`=== FIN OPTIMISATION COMBINÉE === (${nbApplied} opérations de parité appliquées)`);
    return { success: true, nbSwapsAppliquesParite: nbApplied };

  } catch (err) {
    Logger.log(`❌ ERREUR FATALE dans psv5_lancerOptimisationCombinee : ${err.message}`);
    Logger.log(err.stack);
    SpreadsheetApp.getActiveSpreadsheet().toast("Erreur Optimisation Parité!", "Statut", 5);
    ui.alert("Erreur Critique", `Une erreur est survenue: ${err.message}\nConsultez les logs.`, ui.ButtonSet.OK);
    return { success: false, error: err.toString() };
  }
}

// ================================================================
// 5.  SOLUTION ULTIME (+ FLEXIBILITE EFFECTIF MAX)
// ================================================================
function psv5_solutionUltimeOverflow() {
  const ui = SpreadsheetApp.getUi();
  const cfgAppel = getConfig();
  psv5_initialiserDebugMode(cfgAppel);

  const cfg = psv5_prepareCfg(cfgAppel);
  cfg.PSV5_EFFECTIF_MAX_CIBLE_INITIAL += (cfg.PSV5_FLEXIBILITE_EFFECTIF_MAX || 1); 
  cfg.PSV5_EFFECTIF_MAX_STRICT += (cfg.PSV5_FLEXIBILITE_EFFECTIF_MAX || 1);
  cfg.PATCH_VERSION = "2.5.3-Overflow";
  
  Logger.log(`=== SOLUTION ULTIME OVERFLOW (Parity Intelligent Patch v${cfg.PATCH_VERSION}) ===`);
  Logger.log(`   EFFECTIF_MAX_CIBLE_INITIAL (pour moves) augmenté à ${cfg.PSV5_EFFECTIF_MAX_CIBLE_INITIAL}. EFFECTIF_MAX_STRICT augmenté à ${cfg.PSV5_EFFECTIF_MAX_STRICT}.`);
  SpreadsheetApp.getActiveSpreadsheet().toast("Solution Ultime (Overflow) en cours...", "Statut", -1);

  try {
    if (typeof V2_Ameliore_PreparerDonnees_AvecSEXE !== 'function') {
        Logger.log("ERREUR CRITIQUE: V2_Ameliore_PreparerDonnees_AvecSEXE manquante.");
        ui.alert("Erreur Configuration", "Fonction de préparation des données manquante.", ui.ButtonSet.OK);
        return { success: false, error: "Préparation des données manquante." };
    }
    const ctx = V2_Ameliore_PreparerDonnees_AvecSEXE(cfg);
    if (!ctx || !ctx.classesState) {
        Logger.log("ERREUR CRITIQUE: Préparation des données (overflow) n'a pas retourné un contexte valide.");
        ui.alert("Erreur Préparation", "La préparation des données (overflow) n'a pas fonctionné.", ui.ButtonSet.OK);
        return { success: false, error: "Contexte de données invalide après préparation (overflow)." };
    }

    const opsParite = psv5_nirvanaV2_CorrectionPariteINTELLIGENTE(ctx, cfg);
    let nbApplied = 0;
    if (opsParite.length > 0) {
      nbApplied = psv5_AppliquerSwapsSafeEtLog(opsParite, ctx, cfg);
      Logger.log(`✅ ${nbApplied} opérations (avec overflow) appliquées.`);
      ui.alert("Solution Ultime Terminée", `${nbApplied} opérations (avec tolérance d'overflow) ont été appliquées.`, ui.ButtonSet.OK);
    } else {
      Logger.log("Aucune opération trouvée même avec la tolérance d'overflow.");
      ui.alert("Info Solution Ultime", "Aucune opération trouvée même avec la tolérance d'overflow.", ui.ButtonSet.OK);
    }
    SpreadsheetApp.getActiveSpreadsheet().toast("Optimisation (Overflow) terminée!", "Statut", 5);
    return { success: true, nbSwapsAppliquesParite: nbApplied };

  } catch (err) {
    Logger.log(`❌ ERREUR FATALE dans psv5_solutionUltimeOverflow : ${err.message}`);
    Logger.log(err.stack);
    SpreadsheetApp.getActiveSpreadsheet().toast("Erreur!", "Statut", 5);
    ui.alert("Erreur Critique (Overflow)", `Une erreur est survenue: ${err.message}`, ui.ButtonSet.OK);
    return { success: false, error: err.toString() };
  }
}


// ==================================================================
// FONCTIONS DE LANCEMENT AGRESSIF ET DIAGNOSTIC (PROPOSÉES PAR L'UTILISATEUR)
// Intégrées et adaptées pour utiliser les préfixes psv5_ et la gestion de PSV5_DEBUG_MODE
// ==================================================================

/**
 * Lancer le patch avec des paramètres AGRESSIFS
 */
function lancerPariteAgressiveAvecConfig() { // Point d'entrée UI
  const ui = SpreadsheetApp.getUi();
  let configBase; 
  try {
    Logger.log("=== LANCEMENT PARITÉ AGRESSIVE AVEC CONFIG FORCÉE ===");
    configBase = getConfig(); // Récupérer la config de base
    
    // Cloner pour ne pas modifier l'objet config global du Spreadsheet si getConfig() retourne une référence
    const configAgresive = JSON.parse(JSON.stringify(configBase));

    configAgresive.DEBUG_MODE_PARITY_STRATEGY = true; 
    
    // Utiliser les clés PSV5_ pour la configuration spécifique de ce module
    configAgresive.PSV5_PARITY_TOLERANCE = configAgresive.PSV5_PARITY_TOLERANCE_AGRESSIF ?? 1;
    configAgresive.PSV5_SEUIL_SURPLUS_POSITIF_URGENT = configAgresive.PSV5_SEUIL_POSITIF_AGRESSIF ?? 3;
    configAgresive.PSV5_SEUIL_SURPLUS_NEGATIF_URGENT = configAgresive.PSV5_SEUIL_NEGATIF_AGRESSIF ?? -3;
    configAgresive.PSV5_EFFECTIF_MAX_CIBLE_INITIAL = configAgresive.PSV5_EFFECTIF_CIBLE_AGRESSIF || (configAgresive.EFFECTIF_MAX || 26);
    configAgresive.PSV5_FLEXIBILITE_EFFECTIF_MAX = configAgresive.PSV5_FLEXIBILITE_AGRESSIF ?? 1;
    // EFFECTIF_MAX_STRICT sera recalculé dans psv5_prepareCfg basé sur CIBLE_INITIAL et FLEXIBILITE
    configAgresive.PSV5_POTENTIEL_CORRECTION_FACTOR = configAgresive.PSV5_POTENTIEL_FACTOR_AGRESSIF || 3.0;
    configAgresive.PSV5_MAX_ITER_STRATEGIE = configAgresive.PSV5_MAX_ITER_AGRESSIF || 10;
    configAgresive.PSV5_DIFF_DELTA_MIN_POUR_TRANSFERT_MEME_SENS = configAgresive.PSV5_DIFF_DELTA_TRANSFERT_AGRESSIF ?? 1; // Plus permissif
    
    // Initialiser le mode debug du module avec cette config agressive
    psv5_initialiserDebugMode(configAgresive); 

    if (PSV5_DEBUG_MODE) { // Log basé sur l'état réel de PSV5_DEBUG_MODE
        Logger.log("Configuration agressive préparée pour le patch:");
        Logger.log(`  - PSV5_PARITY_TOLERANCE : ${configAgresive.PSV5_PARITY_TOLERANCE}`);
        Logger.log(`  - PSV5_SEUIL_SURPLUS_POSITIF_URGENT : ${configAgresive.PSV5_SEUIL_SURPLUS_POSITIF_URGENT}`);
        // ... autres logs de config si besoin ...
    }
    
    const resultat = psv5_lancerOptimisationAvecConfigForcee(configAgresive); // Appel de la fonction qui prend une config
    
    if (resultat.success) {
      ui.alert('Parité Agressive', 
        `✅ Optimisation agressive terminée !\n\n` +
        `${resultat.nbSwapsAppliquesParite || 0} opérations appliquées.\n\n` +
        `Vérifiez les logs pour voir le détail des actions.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Erreur (Agressive)', resultat.error || 'Erreur inconnue lors du lancement agressif.', ui.ButtonSet.OK);
    }
    return resultat;
    
  } catch (e) {
    Logger.log(`❌ Erreur dans lancerPariteAgressiveAvecConfig: ${e.message}`);
    Logger.log(e.stack);
    ui.alert('Erreur Fatale (Agressive)', e.message, ui.ButtonSet.OK);
    return { success: false, error: e.message };
  }
}

/**
 * Version de psv5_lancerOptimisationCombinee qui utilise une config fournie.
 */
function psv5_lancerOptimisationAvecConfigForcee(configForcee) {
  const ui = SpreadsheetApp.getUi();
  // PSV5_DEBUG_MODE est déjà initialisé par l'appelant (lancerPariteAgressiveAvecConfig)
  // psv5_prepareCfg va utiliser configForcee
  const cfg = psv5_prepareCfg(configForcee); 
  cfg.PATCH_VERSION = cfg.PATCH_VERSION || "2.5.3-ForcedConfig"; 
  
  Logger.log(`=== OPTIMISATION COMBINÉE (Config Forcée v${cfg.PATCH_VERSION}) – Démarrage ===`);
  SpreadsheetApp.getActiveSpreadsheet().toast("Optimisation Parité (Config Forcée)...", "Statut", -1);

  try {
    if (typeof V2_Ameliore_PreparerDonnees_AvecSEXE !== 'function') {
      throw new Error("Fonction V2_Ameliore_PreparerDonnees_AvecSEXE manquante.");
    }
    const ctx = V2_Ameliore_PreparerDonnees_AvecSEXE(cfg);
    if (!ctx || !ctx.classesState) {
      throw new Error("Contexte de données invalide après V2_Ameliore_PreparerDonnees_AvecSEXE.");
    }
    if (PSV5_DEBUG_MODE) Logger.log("Préparation des données (Config Forcée) terminée.");
        
    // Afficher l'état initial avec les seuils agressifs
    if (PSV5_DEBUG_MODE) {
        Logger.log("\n=== ÉTAT INITIAL (avec config forcée) ===");
        Object.entries(ctx.classesState).forEach(([classe, eleves]) => {
          const nbF = eleves.filter(e => e.SEXE === 'F').length;
          const nbM = eleves.filter(e => e.SEXE === 'M').length;
          const delta = nbM - nbF;
          const urgencePos = delta > cfg.PSV5_SEUIL_SURPLUS_POSITIF_URGENT;
          const urgenceNeg = delta < cfg.PSV5_SEUIL_SURPLUS_NEGATIF_URGENT;
          Logger.log(`${classe.padEnd(10)}: ${nbF}F/${nbM}M (Δ${delta > 0 ? '+' : ''}${delta}) ${urgencePos ? "🚨POS" : (urgenceNeg ? "🚨NEG" : "")}`);
        });
    }

    if (PSV5_DEBUG_MODE) Logger.log("\nLancement psv5_nirvanaV2_CorrectionPariteINTELLIGENTE (Config Forcée)...");
    const opsParite = psv5_nirvanaV2_CorrectionPariteINTELLIGENTE(ctx, cfg);
    if (PSV5_DEBUG_MODE) Logger.log(`CorrectionPariteINTELLIGENTE (Config Forcée) a retourné ${opsParite.length} ops uniques.`);

    let nbApplied = 0;
    if (opsParite.length > 0) {
      if (PSV5_DEBUG_MODE) Logger.log(`Application de ${opsParite.length} ops parité (Config Forcée)...`);
      nbApplied = psv5_AppliquerSwapsSafeEtLog(opsParite, ctx, cfg);
      Logger.log(`✅ ${nbApplied} ops parité (Config Forcée) appliquées.`);
      ui.alert("Optimisation Parité (Config Forcée) Terminée", 
        `${nbApplied} corrections appliquées avec la config forcée.\n\n` +
        `Seuils utilisés:\n` +
        `- Tolérance: ±${cfg.PSV5_PARITY_TOLERANCE}\n` +
        `- Action Positive dès: Δ +${cfg.PSV5_SEUIL_SURPLUS_POSITIF_URGENT}\n` +
        `- Action Négative dès: Δ ${cfg.PSV5_SEUIL_SURPLUS_NEGATIF_URGENT}\n\n` +
        `Consultez les logs pour le détail.`,
        ui.ButtonSet.OK);
    } else {
      Logger.log("Aucune op parité (Config Forcée) trouvée.");
      ui.alert("Info Parité (Config Forcée)", 
      "Aucune opération de parité trouvée, même avec la configuration forcée.\n\n" +
      "Vérifiez :\n" +
      "- La mobilité des élèves (colonne MOBILITE = LIBRE)\n" +
      "- Les contraintes d'options pour les élèves mobiles\n" +
      "- Les logs détaillés pour les raisons de non-action.",
      ui.ButtonSet.OK);
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast("Optimisation (Config Forcée) terminée!", "Statut", 5);
    Logger.log(`=== FIN OPTIMISATION COMBINÉE (Config Forcée) === (${nbApplied} ops parité appliquées)`);
    return { success: true, nbSwapsAppliquesParite: nbApplied };

  } catch (err) {
    Logger.log(`❌ ERREUR FATALE dans psv5_lancerOptimisationAvecConfigForcee : ${err.message}\n${err.stack}`);
    SpreadsheetApp.getActiveSpreadsheet().toast("Erreur Optimisation Parité!", "Statut", 5);
    ui.alert("Erreur Critique", `Erreur: ${err.message}`, ui.ButtonSet.OK);
    return { success: false, error: err.toString() };
  }
}

/**
 * Diagnostic rapide de la mobilité
 */
function diagnosticMobiliteRapide() {
  const ui = SpreadsheetApp.getUi();
  const cfgAppel = getConfig();
  psv5_initialiserDebugMode(cfgAppel); 

  Logger.log("=== DIAGNOSTIC MOBILITÉ RAPIDE ===");
  try {
    if (typeof V2_Ameliore_PreparerDonnees_AvecSEXE !== 'function') throw new Error("V2_Ameliore_PreparerDonnees_AvecSEXE manquante");
    const config = psv5_prepareCfg(cfgAppel); 
    const ctx = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
    if (!ctx || !ctx.classesState) throw new Error("Contexte de données invalide");

    let logDetail = "";
    Object.entries(ctx.classesState).forEach(([classe, eleves]) => {
      const nbF = eleves.filter(e => e.SEXE === 'F').length;
      const nbM = eleves.filter(e => e.SEXE === 'M').length;
      const delta = nbM - nbF;
      
      const mobilesF = eleves.filter(e => e.SEXE === 'F' && psv5_isMobile(e, config)).length;
      const mobilesM = eleves.filter(e => e.SEXE === 'M' && psv5_isMobile(e, config)).length;
      
      const line = `${classe.padEnd(10)}: ${nbF}F/${nbM}M (Δ${delta >= 0 ? '+' : ''}${delta}) - Mobiles: ${mobilesF}F/${mobilesM}M`;
      Logger.log(line);
      logDetail += line + "\n";
      
      if (Math.abs(delta) > config.PSV5_PARITY_TOLERANCE) { // Comparer à la tolérance configurée
        if (delta > 0 && mobilesM === 0) {
          const warn = `  ⚠️ PROBLÈME: Surplus de ${delta} garçons mais AUCUN mobile !`;
          Logger.log(warn); logDetail += warn + "\n";
        } else if (delta < 0 && mobilesF === 0) {
          const warn = `  ⚠️ PROBLÈME: Surplus de ${-delta} filles mais AUCUNE mobile !`;
          Logger.log(warn); logDetail += warn + "\n";
        }
      }
    });
    ui.showModalDialog(HtmlService.createHtmlOutput("<pre>" + logDetail.replace(/\n/g, "<br>") + "</pre>").setWidth(500).setHeight(400), "Diagnostic Mobilité");
  } catch (e) {
    Logger.log(`ERREUR diagnosticMobiliteRapide: ${e.message}`);
    ui.alert("Erreur Diagnostic Mobilité", e.message, ui.ButtonSet.OK);
  }
}

/**
 * Test avec simulation (sans appliquer)
 */
function testerStrategieDeuxCoupsSansAppliquer() {
  const ui = SpreadsheetApp.getUi();
  Logger.clear();
  Logger.log("=== TEST STRATÉGIE DEUX COUPS (SIMULATION UNIQUEMENT) ===");
  
  const config = getConfig(); // Config de base
  
  // Forcer les paramètres agressifs pour ce test de simulation
  config.DEBUG_MODE_PARITY_STRATEGY = true;
  config.PSV5_PARITY_TOLERANCE = 1;
  config.PSV5_SEUIL_SURPLUS_POSITIF_URGENT = 3;
  config.PSV5_SEUIL_SURPLUS_NEGATIF_URGENT = -3;
  config.PSV5_EFFECTIF_MAX_CIBLE_INITIAL = 26;
  config.PSV5_FLEXIBILITE_EFFECTIF_MAX = 1;
  // PSV5_EFFECTIF_MAX_STRICT sera calculé dans psv5_prepareCfg
  config.PSV5_POTENTIEL_CORRECTION_FACTOR = 3.0;
  config.PSV5_MAX_ITER_STRATEGIE = 10;
  config.PSV5_DIFF_DELTA_MIN_POUR_TRANSFERT_MEME_SENS = 1; 
  
  psv5_initialiserDebugMode(config); 
  const cfgPréparée = psv5_prepareCfg(config); 
  
  try {
    if (typeof V2_Ameliore_PreparerDonnees_AvecSEXE !== 'function') throw new Error("V2_Ameliore_PreparerDonnees_AvecSEXE manquante");
    const ctx = V2_Ameliore_PreparerDonnees_AvecSEXE(cfgPréparée);
    if (!ctx || !ctx.classesState) throw new Error("Contexte de données invalide");

    Logger.log("Configuration utilisée pour la simulation:");
    Logger.log(JSON.stringify({
      Tolérance: cfgPréparée.PSV5_PARITY_TOLERANCE,
      SeuilPos: cfgPréparée.PSV5_SEUIL_SURPLUS_POSITIF_URGENT,
      SeuilNeg: cfgPréparée.PSV5_SEUIL_SURPLUS_NEGATIF_URGENT,
      DiffDeltaMin: cfgPréparée.PSV5_DIFF_DELTA_MIN_POUR_TRANSFERT_MEME_SENS,
      EffMaxCible: cfgPréparée.PSV5_EFFECTIF_MAX_CIBLE_INITIAL,
      FlexEff: cfgPréparée.PSV5_FLEXIBILITE_EFFECTIF_MAX,
      MaxIter: cfgPréparée.PSV5_MAX_ITER_STRATEGIE
    }, null, 2));
  
    const ops = psv5_nirvanaV2_CorrectionPariteINTELLIGENTE(ctx, cfgPréparée);
  
    Logger.log(`\n=== RÉSULTAT SIMULATION ===`);
    Logger.log(`${ops.length} opérations générées (non appliquées).`);
  
    let logDetailOps = "";
    ops.forEach((op, idx) => {
      let line = "";
      if (op.type === "SWAP") {
        line = `${idx + 1}. SWAP: ${op.eleveA.NOM}(${op.classeA}) ↔ ${op.eleveB.NOM}(${op.classeB}) | ${op.motif || ''}`;
      } else {
        line = `${idx + 1}. MOVE: ${op.eleveA.NOM} de ${op.classeA} → ${op.classeB} | ${op.motif || ''}`;
      }
      Logger.log(line);
      logDetailOps += line + "\n";
    });
    
    if (ops.length > 0) {
      ui.showModalDialog(HtmlService.createHtmlOutput("<pre>" + logDetailOps.replace(/\n/g, "<br>") + "</pre>").setWidth(700).setHeight(450), "Opérations Simulées");
    } else {
      ui.alert("Simulation", "Aucune opération générée par la simulation.", ui.ButtonSet.OK);
    }

  } catch (e) {
    Logger.log(`ERREUR testerStrategieDeuxCoupsSansAppliquer: ${e.message}\n${e.stack}`);
    ui.alert("Erreur Simulation", e.message, ui.ButtonSet.OK);
  }
}

/** Fonctions de diagnostic avancées supplémentaires fournies par l'utilisateur */
function diagnosticBlocagesPariteAvance() {
  const ui = SpreadsheetApp.getUi();
  const config = getConfig();
  // Forcer le debug pour cette fonction de diagnostic
  config.DEBUG_MODE_PARITY_STRATEGY = true; 
  psv5_initialiserDebugMode(config);
  
  try {
    const cfg = psv5_prepareCfg(config);
    if (typeof V2_Ameliore_PreparerDonnees_AvecSEXE !== 'function') throw new Error("V2_Ameliore_PreparerDonnees_AvecSEXE manquante");
    const ctx = V2_Ameliore_PreparerDonnees_AvecSEXE(cfg);
    if (!ctx || !ctx.classesState) throw new Error("Contexte invalide");
    
    let rapport = "=== DIAGNOSTIC AVANCÉ DES BLOCAGES DE PARITÉ (v2.5.3) ===\n\n";
    
    let totalF = 0, totalM = 0;
    let classesPlusFilles = 0, classesPlusGarcons = 0, classesEquilibrees = 0;
    
    Object.values(ctx.classesState).forEach(eleves => {
      const nbFClasse = eleves.filter(e => e.SEXE === 'F').length;
      const nbMClasse = eleves.filter(e => e.SEXE === 'M').length;
      totalF += nbFClasse;
      totalM += nbMClasse;
      const delta = nbMClasse - nbFClasse;
      
      if (Math.abs(delta) <= cfg.PSV5_PARITY_TOLERANCE) classesEquilibrees++;
      else if (delta > 0) classesPlusGarcons++;
      else classesPlusFilles++;
    });
    
    rapport += `RÉPARTITION GLOBALE:\n`;
    rapport += `Total: ${totalF}F / ${totalM}M (Δ global: ${totalM - totalF})\n`;
    rapport += `Classes: ${classesPlusGarcons} avec surplus M, ${classesPlusFilles} avec surplus F, ${classesEquilibrees} équilibrées (Tol: ±${cfg.PSV5_PARITY_TOLERANCE})\n\n`;
    
    rapport += `PROBLÈMES STRUCTURELS IDENTIFIÉS:\n`;
    if (classesPlusFilles === 0 && classesPlusGarcons > 0 && Object.keys(ctx.classesState).length === classesPlusGarcons) {
      rapport += `⚠️ TOUTES les classes ont un surplus de GARÇONS (ou sont équilibrées)!\n`;
      rapport += `   → Difficile de rééquilibrer sans classes "receveuses" de garçons (avec déficit M / surplus F).\n`;
      rapport += `   → La stratégie modifiée pour transfert vers "moins pire" est activée.\n\n`;
    } else if (classesPlusGarcons === 0 && classesPlusFilles > 0 && Object.keys(ctx.classesState).length === classesPlusFilles) {
      rapport += `⚠️ TOUTES les classes ont un surplus de FILLES (ou sont équilibrées)!\n`;
      rapport += `   → Idem, stratégie modifiée pour transfert vers "moins pire" activée.\n\n`;
    } else if (classesPlusGarcons > 0 && classesPlusFilles > 0) {
        rapport += `ℹ️ Des classes ont un surplus de M et d'autres un surplus de F. Swaps directs devraient être possibles.\n\n`;
    } else {
        rapport += `ℹ️ Situation de parité globalement équilibrée ou peu de classes déséquilibrées.\n\n`;
    }
    
    rapport += `ANALYSE DE LA MOBILITÉ PAR CLASSE:\n`;
    let totalMobilesF = 0, totalMobilesM = 0;
    
    Object.entries(ctx.classesState).forEach(([classe, eleves]) => {
      const nbF = eleves.filter(e => e.SEXE === 'F').length;
      const nbM = eleves.filter(e => e.SEXE === 'M').length;
      const mobilesF = eleves.filter(e => e.SEXE === 'F' && psv5_isMobile(e, cfg)).length;
      const mobilesM = eleves.filter(e => e.SEXE === 'M' && psv5_isMobile(e, cfg)).length;
      const delta = nbM - nbF;
      
      totalMobilesF += mobilesF;
      totalMobilesM += mobilesM;
      
      rapport += `${classe.padEnd(10)}: ${nbF}F/${nbM}M (Δ${delta >= 0 ? '+' : ''}${delta}) - Mobiles: ${mobilesF}F/${mobilesM}M`;
      
      if (Math.abs(delta) > cfg.PSV5_PARITY_TOLERANCE) {
        if (delta > 0 && mobilesM === 0) rapport += ` ❌ Surplus M mais AUCUN mobile!`;
        else if (delta < 0 && mobilesF === 0) rapport += ` ❌ Surplus F mais AUCUNE mobile!`;
      }
      rapport += `\n`;
    });
    
    rapport += `\nTotal mobiles: ${totalMobilesF}F / ${totalMobilesM}M\n\n`;
    
    rapport += `PARAMÈTRES DE STRATÉGIE ACTUELS (psv5_prepareCfg):\n`;
    rapport += `  Tolérance Parité: ±${cfg.PSV5_PARITY_TOLERANCE}\n`;
    rapport += `  Seuil Surplus Positif Urgent (M): Δ > ${cfg.PSV5_SEUIL_SURPLUS_POSITIF_URGENT}\n`;
    rapport += `  Seuil Surplus Négatif Urgent (F): Δ < ${cfg.PSV5_SEUIL_SURPLUS_NEGATIF_URGENT}\n`;
    rapport += `  Diff. Delta Min pour Transfert Même Sens: ${cfg.PSV5_DIFF_DELTA_MIN_POUR_TRANSFERT_MEME_SENS}\n`;
    rapport += `  Effectif Max Cible Initial (avant flex): ${cfg.PSV5_EFFECTIF_MAX_CIBLE_INITIAL}\n`;
    rapport += `  Flexibilité Effectif Max (pour moves/swaps): +${cfg.PSV5_FLEXIBILITE_EFFECTIF_MAX}\n`;
    rapport += `  Effectif Max Strict (après reduceOverflow): ${cfg.PSV5_EFFECTIF_MAX_STRICT}\n\n`;

    rapport += `SUGGESTIONS BASÉES SUR CE DIAGNOSTIC:\n`;
    if ((classesPlusFilles === 0 && classesPlusGarcons > 0) || (classesPlusGarcons === 0 && classesPlusFilles > 0)) {
      rapport += `- S'assurer que PSV5_DIFF_DELTA_MIN_POUR_TRANSFERT_MEME_SENS (actuel: ${cfg.PSV5_DIFF_DELTA_MIN_POUR_TRANSFERT_MEME_SENS}) est assez petit pour permettre des transferts.\n`;
    }
    if (totalMobilesM < 5 || totalMobilesF < 5) {
      rapport += `- VÉRIFIER LA COLONNE 'MOBILITE' DANS LES FEUILLES ÉLÈVES. Peu d'élèves 'LIBRE'.\n`;
    }
    rapport += `- Si toujours bloqué, utiliser 'LANCER PARITÉ AGRESSIVE' ou 'Solution Ultime (Overflow +)' du menu.\n`;
    rapport += `- Exécuter 'Tester Stratégie (Simu Agressive)' pour voir les opérations que l'algo envisagerait avec des paramètres stricts.\n`;

    Logger.log(rapport);
    const html = HtmlService.createHtmlOutput(`<pre style="font-family: monospace; font-size: 11px;">${rapport.replace(/\n/g, '<br>')}</pre>`)
      .setWidth(850).setHeight(600);
    ui.showModalDialog(html, "Diagnostic Avancé Parité v2.5.3");
    
  } catch (e) {
    Logger.log(`ERREUR diagnosticBlocagesPariteAvance: ${e.message}\n${e.stack}`);
    ui.alert("Erreur Diagnostic", e.message, ui.ButtonSet.OK);
  }
}

function psv5_forcerEquilibrageProgressif() {
  const ui = SpreadsheetApp.getUi();
  const config = getConfig();
  
  config.DEBUG_MODE_PARITY_STRATEGY = true; // Forcer le debug pour cette fonction
  config.PSV5_PARITY_TOLERANCE = 0; 
  config.PSV5_SEUIL_SURPLUS_POSITIF_URGENT = 1; // Agir très vite
  config.PSV5_SEUIL_SURPLUS_NEGATIF_URGENT = -1;
  config.PSV5_MAX_ITER_STRATEGIE = 20; 
  
  psv5_initialiserDebugMode(config);
  
  try {
    Logger.log("=== ÉQUILIBRAGE PROGRESSIF FORCÉ (v2.5.3) ===");
    
    const cfg = psv5_prepareCfg(config);
    if (typeof V2_Ameliore_PreparerDonnees_AvecSEXE !== 'function') throw new Error("V2_Ameliore_PreparerDonnees_AvecSEXE manquante");
    const ctx = V2_Ameliore_PreparerDonnees_AvecSEXE(cfg);
    if (!ctx || !ctx.classesState) throw new Error("Contexte invalide");
    
    const workingClasses = Object.entries(ctx.classesState)
      .map(([nom, eleves]) => ({
        nom, eleves: [...eleves], 
        delta: psv5_deltaParite(eleves), 
        effectif: eleves.length
      }));
    
    let operationsTotal = [];
    let iteration = 0;
    const MAX_ITER_PROGRESSIF = 50; // Sécurité pour cette boucle spécifique
    let aChangeCeTour = true;
    
    while (iteration < MAX_ITER_PROGRESSIF && aChangeCeTour) {
      iteration++;
      aChangeCeTour = false;
      workingClasses.sort((a, b) => b.delta - a.delta); // Du plus grand surplus M au plus grand surplus F
      
      const source = workingClasses[0];
      const cible = workingClasses[workingClasses.length - 1];
      
      if (Math.abs(source.delta) <= cfg.PSV5_PARITY_TOLERANCE && Math.abs(cible.delta) <= cfg.PSV5_PARITY_TOLERANCE) {
        Logger.log(`Équilibrage progressif: Classes extrêmes (${source.nom}, ${cible.nom}) dans la tolérance. Arrêt.`);
        break;
      }
      if (source.delta <= cfg.PSV5_PARITY_TOLERANCE && cible.delta >= -cfg.PSV5_PARITY_TOLERANCE && source.delta >= cible.delta) {
         Logger.log(`Équilibrage progressif: Faible potentiel d'amélioration entre ${source.nom}(Δ${source.delta}) et ${cible.nom}(Δ${cible.delta}). Arrêt.`);
         break;
      }


      const sexeSurplusSource = source.delta > 0 ? "M" : "F";
      const eleveMobileSource = psv5_findMobile(source.eleves, sexeSurplusSource, cfg);
      
      if (!eleveMobileSource) {
        Logger.log(`  It ${iteration}: Pas d'élève ${sexeSurplusSource} mobile dans source ${source.nom}(Δ${source.delta}). Tentative avec la 2e classe la plus déséquilibrée.`);
        if (workingClasses.length > 2) { // Essayer avec la 2e et l'avant-dernière si la première est bloquée
            const altSource = workingClasses[1];
            const altCible = workingClasses[workingClasses.length - 2];
            const altSexeSurplus = altSource.delta > 0 ? "M" : "F";
            const altMobile = psv5_findMobile(altSource.eleves, altSexeSurplus, cfg);
            if (altMobile && psv5_optionOK(altMobile, altCible.nom, ctx) && altCible.effectif < cfg.PSV5_EFFECTIF_MAX_STRICT) {
                 Logger.log(`  MOVE ALT: ${altMobile.NOM} de ${altSource.nom}(Δ${altSource.delta}) → ${altCible.nom}(Δ${altCible.delta})`);
            }
        }
        aChangeCeTour = false; // Pour sortir de la boucle while si on ne peut rien faire.
        continue;
      }
      
      if (cible.effectif >= cfg.PSV5_EFFECTIF_MAX_STRICT) {
        Logger.log(`  It ${iteration}: Cible ${cible.nom} est pleine (${cible.effectif}). Swap facilitateur non implémenté dans cette fonction simplifiée.`);
        aChangeCeTour = false;
        continue;
      }
      
      if (!psv5_optionOK(eleveMobileSource, cible.nom, ctx)) {
        Logger.log(`  It ${iteration}: Option incompatible pour ${eleveMobileSource.NOM} (${sexeSurplusSource}) vers ${cible.nom}.`);
        aChangeCeTour = false;
        continue;
      }
      
      Logger.log(`  It ${iteration} MOVE: ${eleveMobileSource.NOM}(${eleveMobileSource.SEXE}) de ${source.nom}(Δ${source.delta}) → ${cible.nom}(Δ${cible.delta})`);
      operationsTotal.push({
        type: "MOVE", eleveA: eleveMobileSource, classeA: source.nom, classeB: cible.nom,
        motif: `ForcéProg It${iteration} (ΔS${source.delta}→ΔC${cible.delta})`
      });
      
      source.eleves.splice(source.eleves.indexOf(eleveMobileSource),1);
      cible.eleves.push(eleveMobileSource);
      
      source.delta = psv5_deltaParite(source.eleves); source.effectif = source.eleves.length;
      cible.delta = psv5_deltaParite(cible.eleves); cible.effectif = cible.eleves.length;
      aChangeCeTour = true; // Un mouvement a été fait
    }
    if (iteration === MAX_ITER_PROGRESSIF) Logger.log("Équilibrage progressif: MAX_ITER atteinte.");
    
    Logger.log(`\n=== RÉSUMÉ ÉQUILIBRAGE FORCÉ ===`);
    Logger.log(`${operationsTotal.length} opérations planifiées.`);
    
     if (operationsTotal.length > 0) {
      const reponse = ui.alert(
        'Équilibrage Forcé',
        `${operationsTotal.length} mouvements trouvés pour améliorer l'équilibre.\n\n` +
        `Voulez-vous les appliquer ?`,
        ui.ButtonSet.YES_NO
      );
      if (reponse === ui.Button.YES) {
        const nbAppliques = psv5_AppliquerSwapsSafeEtLog(operationsTotal, ctx, cfg);
        ui.alert('Résultat', `${nbAppliques} mouvements appliqués avec succès.`, ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Information', 'Aucun mouvement possible trouvé pour l\'équilibrage forcé progressif.', ui.ButtonSet.OK);
    }
    return operationsTotal;
  } catch (e) {
    Logger.log(`ERREUR psv5_forcerEquilibrageProgressif: ${e.message}\n${e.stack}`);
    ui.alert("Erreur", e.message, ui.ButtonSet.OK);
    return [];
  }
}

function psv5_proposerSwapsManuels() {
  const ui = SpreadsheetApp.getUi();
  const config = getConfig();
  config.DEBUG_MODE_PARITY_STRATEGY = true; // Activer debug pour les helpers
  psv5_initialiserDebugMode(config);
  
  try {
    const cfg = psv5_prepareCfg(config);
    if (typeof V2_Ameliore_PreparerDonnees_AvecSEXE !== 'function') throw new Error("V2_Ameliore_PreparerDonnees_AvecSEXE manquante");
    const ctx = V2_Ameliore_PreparerDonnees_AvecSEXE(cfg);
    if (!ctx || !ctx.classesState) throw new Error("Contexte invalide");
    
    let suggestions = "=== SUGGESTIONS DE SWAPS MANUELS (v2.5.3) ===\n\n";
    suggestions += `Configuration utilisée: Tolérance Parité ±${cfg.PSV5_PARITY_TOLERANCE}\n\n`;
    
    const classesInfos = Object.entries(ctx.classesState).map(([nom,eleves]) => ({
        nom, eleves, delta: psv5_deltaParite(eleves)
    }));
    let nbSuggestions = 0;
    
    for (let i = 0; i < classesInfos.length; i++) {
      for (let j = i + 1; j < classesInfos.length; j++) {
        const classeA = classesInfos[i];
        const classeB = classesInfos[j];
        
        // Condition: déséquilibres opposés et significatifs
        if (classeA.delta * classeB.delta < 0 && 
            Math.abs(classeA.delta) > cfg.PSV5_PARITY_TOLERANCE && 
            Math.abs(classeB.delta) > cfg.PSV5_PARITY_TOLERANCE) {
          
          const sexeSurplusA = classeA.delta > 0 ? "M" : "F";
          const sexeSurplusB = classeB.delta > 0 ? "M" : "F"; // Sera l'opposé de sexeSurplusA

          const mobileA = psv5_findMobile(classeA.eleves, sexeSurplusA, cfg);
          const mobileB = psv5_findMobile(classeB.eleves, sexeSurplusB, cfg); // Cherche le sexe opposé de A dans B
          
          if (mobileA && mobileB && 
              psv5_optionOK(mobileA, classeB.nom, ctx) && 
              psv5_optionOK(mobileB, classeA.nom, ctx)) {
            
            suggestions += `SWAP SUGGÉRÉ #${++nbSuggestions}:\n`;
            suggestions += `  De ${classeA.nom.padEnd(10)} (Δ${classeA.delta}) : Déplacer ${mobileA.NOM} (${mobileA.SEXE}, Opt:${mobileA.OPT||'Aucune'})\n`;
            suggestions += `  Vers ${classeB.nom.padEnd(10)} (Δ${classeB.delta})\n`;
            suggestions += `  ET\n`;
            suggestions += `  De ${classeB.nom.padEnd(10)} (Δ${classeB.delta}) : Déplacer ${mobileB.NOM} (${mobileB.SEXE}, Opt:${mobileB.OPT||'Aucune'})\n`;
            suggestions += `  Vers ${classeA.nom.padEnd(10)} (Δ${classeA.delta})\n`;
            
            // Calculer l'impact
            const deltaA_apres = classeA.delta - (sexeSurplusA === "M" ? 2 : -2); // A perd un M, gagne une F ou inversement
            const deltaB_apres = classeB.delta - (sexeSurplusB === "M" ? 2 : -2);

            suggestions += `  IMPACT PRÉVU:\n`;
            suggestions += `    ${classeA.nom}: Δ ${classeA.delta} → ${deltaA_apres}\n`;
            suggestions += `    ${classeB.nom}: Δ ${classeB.delta} → ${deltaB_apres}\n\n`;
            
            if (nbSuggestions >= 15) break; 
          }
        }
      }
      if (nbSuggestions >= 15) break;
    }
    
    if (nbSuggestions === 0) {
      suggestions += "Aucun swap direct évident trouvé entre classes à déséquilibres opposés.\n";
      suggestions += "Essayez 'Forcer Équilibrage Progressif' ou 'Lancer Parité Agressive'.\n";
      suggestions += "Vérifiez aussi la mobilité des élèves et les contraintes d'options via les diagnostics.\n";
    }
    
    Logger.log(suggestions);
    const html = HtmlService.createHtmlOutput(`<pre style="font-family: monospace; font-size: 11px;">${suggestions.replace(/\n/g, '<br>')}</pre>`)
      .setWidth(750).setHeight(550);
    ui.showModalDialog(html, "Suggestions de Swaps Manuels");
    
  } catch (e) {
    Logger.log(`ERREUR psv5_proposerSwapsManuels: ${e.message}`);
    ui.alert("Erreur Suggestions", e.message, ui.ButtonSet.OK);
  }
}
