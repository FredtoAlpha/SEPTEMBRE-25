// Bandeau maintenance : tous les utilitaires sont centralisés dans Utils.js
// Utilisez uniquement Utils.idx(header, name) et Utils.logAction(...)
/**
 * PHASE 2 - RÉPARTITION UNIVERSELLE SELON LES CODES ASSO/DISSO
 * Version avec VÉRIFICATION FINALE et correction des filtres
 */

/**
 * Détermine si un élève a des contraintes fortes (LV2 autre que ESP ou options)
 */
function aContraintesFortes(eleve) {
  return (eleve.lv2 && eleve.lv2.toUpperCase() !== "ESP" && eleve.lv2.toUpperCase() !== "") || 
         (eleve.options && eleve.options.length > 0);
}

/**
 * Vérifie si une classe peut accueillir un élève selon ses contraintes
 */
function classeCompatible(eleve, classe, classesPourOption) {
  // Si l'élève n'a pas de contraintes, il peut aller partout
  if (!aContraintesFortes(eleve)) return true;
  
  // Vérifier la LV2
  if (eleve.lv2 && eleve.lv2.toUpperCase() !== "ESP" && eleve.lv2.toUpperCase() !== "") {
    const classesLV2 = classesPourOption[eleve.lv2.toUpperCase()];
    if (!classesLV2 || !classesLV2.has(classe)) return false;
  }
  
  // Vérifier chaque option
  for (const opt of eleve.options || []) {
    if (opt) {
      const classesOpt = classesPourOption[opt.toUpperCase()];
      if (!classesOpt || !classesOpt.has(classe)) return false;
    }
  }
  
  return true;
}

function repartirPhase2() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Force un rafraîchissement des données et supprime les filtres
    SpreadsheetApp.flush();
    
    // CORRECTION: Supprimer tous les filtres avant de commencer
    supprimerTousLesFiltres();
    
    // 1. Lire les élèves déjà placés dans les onglets TEST (résultat Phase 1)
    const elevesPlacesPhase1 = lireElevesDesOngletsTest();
    if (!elevesPlacesPhase1 || elevesPlacesPhase1.length === 0) {
      const message = "Aucun élève trouvé dans les onglets TEST - Exécutez d'abord la Phase 1";
      SpreadsheetApp.getActiveSpreadsheet().toast(message, "Phase 2", 10);
      return { success: false, message: message };
    }
    
    Logger.log(`${elevesPlacesPhase1.length} élèves lus depuis les onglets TEST (Phase 1)`);
    
    // 2. Lire tous les élèves depuis CONSOLIDATION
    const tousLesEleves = lireTousLesElevesDepuisConsolidation();
    if (!tousLesEleves || tousLesEleves.length === 0) {
      const message = "Aucun élève trouvé dans CONSOLIDATION";
      SpreadsheetApp.getActiveSpreadsheet().toast(message, "Phase 2", 10);
      return { success: false, message: message };
    }
    
    // 3. Fusionner les données et détecter les langues et options présentes
    const { eleves, languesDetectees, optionsDetectees, classesPourOption } = fusionnerDonnees(tousLesEleves, elevesPlacesPhase1);
    
    Logger.log(`Langues détectées: ${Array.from(languesDetectees).join(', ')}`);
    Logger.log(`Options détectées: ${Array.from(optionsDetectees).join(', ')}`);
    
    // 4. Traiter les codes ASSO
    const resultatAsso = appliquerReglesAsso(eleves, languesDetectees, optionsDetectees, classesPourOption);
    
    // 5. Traiter les codes DISSO (version simplifiée)
    const resultatDisso = appliquerReglesDisso(eleves, languesDetectees, optionsDetectees, classesPourOption);
    
    // 6. NOUVELLE ÉTAPE: Vérification finale et correction des conflits DISSO
    const resultatVerification = verifierEtCorrigerConflitsDisso(eleves, classesPourOption);
    
    // 7. Mettre à jour les onglets TEST
    try {
      Logger.log("Début de l'écriture des résultats dans les onglets TEST...");
      const resultatEcriture = ecrireResultatsDansOngletsTest(eleves);
      Logger.log(`Écriture terminée avec ${resultatEcriture.totalEcrits} élèves écrits`);
      
      // 8. Afficher les résultats
      const message = `Phase 2 terminée: ${resultatAsso.deplacements} élèves associés, ${resultatDisso.deplacements} élèves dissociés, ${resultatVerification.corrections} corrections finales, ${resultatEcriture.totalEcrits} élèves écrits`;
      SpreadsheetApp.getActiveSpreadsheet().toast(message, "Phase 2", 10);
      
      return {
        success: true,
        message: message,
        asso: resultatAsso.deplacements,
        disso: resultatDisso.deplacements,
        corrections: resultatVerification.corrections,
        ecrits: resultatEcriture.totalEcrits
      };
    } catch (errEcriture) {
      Logger.log(`ERREUR lors de l'écriture: ${errEcriture.message}`);
      const message = `Phase 2 partiellement terminée: ${resultatAsso.deplacements} élèves associés, ${resultatDisso.deplacements} élèves dissociés, mais ERREUR d'écriture`;
      SpreadsheetApp.getActiveSpreadsheet().toast(message, "Phase 2", 10);
      
      return {
        success: false,
        message: message,
        asso: resultatAsso.deplacements,
        disso: resultatDisso.deplacements,
        erreur: errEcriture.message
      };
    }
    
  } catch (e) {
    const message = "Erreur: " + e.message;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Phase 2", 10);
    return { success: false, message: message };
  }
}

/**
 * NOUVELLE FONCTION: Supprime tous les filtres pour éviter les erreurs de consolidation
 */
function supprimerTousLesFiltres() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    sheets.forEach(sheet => {
      try {
        // Supprimer le filtre s'il existe
        if (sheet.getFilter()) {
          sheet.getFilter().remove();
          Logger.log(`Filtre supprimé dans ${sheet.getName()}`);
        }
      } catch (e) {
        // Ignorer les erreurs de suppression de filtre
      }
    });
  } catch (e) {
    Logger.log(`Erreur lors de la suppression des filtres: ${e.message}`);
  }
}

/**
 * NOUVELLE FONCTION: Vérification finale et correction des conflits DISSO avec OPTIMISATION GLOBALE
 * Cette fonction cherche la meilleure répartition possible en tenant compte de toutes les contraintes
 */
function verifierEtCorrigerConflitsDisso(eleves, classesPourOption) {
  Logger.log("\n=== VÉRIFICATION FINALE DES CONFLITS DISSO AVEC OPTIMISATION ===");
  
  let corrections = 0;
  
  // Trouver tous les onglets TEST disponibles
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classesPossibles = ss.getSheets()
    .filter(s => s.getName().match(/^[\d°]+\d*TEST$/))
    .map(s => s.getName());
  
  // Lire les contraintes de l'onglet _STRUCTURE si disponible
  const structureConstraints = lireContraintesStructure();
  
  // Identifier tous les élèves avec code DISSO
  const elevesAvecDisso = eleves.filter(e => e.dissoCode);
  
  // Grouper par code DISSO
  const groupesDisso = new Map();
  elevesAvecDisso.forEach(e => {
    if (!groupesDisso.has(e.dissoCode)) {
      groupesDisso.set(e.dissoCode, []);
    }
    groupesDisso.get(e.dissoCode).push(e);
  });
  
  // Pour chaque groupe DISSO, optimiser la répartition
  groupesDisso.forEach((groupe, codeDisso) => {
    if (groupe.length < 2) return;
    
    Logger.log(`\n🔍 Analyse du code ${codeDisso}: ${groupe.length} élèves`);
    
    // Analyser les contraintes du groupe
    const analyseCIntraintes = analyserContraintesGroupe(groupe, classesPourOption);
    Logger.log(`  Contraintes: ${JSON.stringify(analyseCIntraintes)}`);
    
    // Générer toutes les répartitions possibles
    const repartitionsPossibles = genererRepartitionsPossibles(
      groupe, 
      classesPossibles, 
      classesPourOption, 
      structureConstraints
    );
    
    if (repartitionsPossibles.length === 0) {
      Logger.log(`  ❌ Aucune répartition possible pour ${codeDisso}`);
      return; // Passer au groupe suivant dans forEach
    }
    
    // Évaluer chaque répartition
    const repartitionsEvaluees = repartitionsPossibles.map(rep => ({
      repartition: rep,
      score: evaluerRepartition(rep, eleves, structureConstraints, classesPourOption)
    }));
    
    // Trier par score (plus élevé = meilleur)
    repartitionsEvaluees.sort((a, b) => b.score - a.score);
    
    // Appliquer la meilleure répartition
    const meilleureRepartition = repartitionsEvaluees[0];
    Logger.log(`  ✅ Meilleure répartition trouvée (score: ${meilleureRepartition.score})`);
    
    // Appliquer les changements
    meilleureRepartition.repartition.forEach(({eleve, classe}) => {
      const classePrevue = eleve.planifie ? eleve.classeFinale : eleve.classeActuelle;
      
      if (classePrevue !== classe) {
        eleve.planifie = true;
        eleve.classeFinale = classe;
        eleve.raison = `Optimisation ${codeDisso}`;
        corrections++;
        Logger.log(`  ↻ ${eleve.id} (${eleve.nomComplet}) → ${classe}`);
      }
    });
  });
  
  // Vérification finale des conflits résiduels
  const conflitsResiduels = verifierConflitsResiduels(eleves);
  if (conflitsResiduels.length > 0) {
    Logger.log(`\n⚠️ ${conflitsResiduels.length} conflits résiduels détectés, correction en cours...`);
    corrections += corrigerConflitsResiduels(conflitsResiduels, eleves, classesPourOption, structureConstraints);
  }
  
  Logger.log(`\n=== FIN VÉRIFICATION: ${corrections} corrections effectuées ===`);
  
  return { corrections };
}

/**
 * Lit les contraintes depuis l'onglet _STRUCTURE
 */
function lireContraintesStructure() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("_STRUCTURE");
    
    if (!sheet) {
      Logger.log("Onglet _STRUCTURE non trouvé - utilisation des paramètres par défaut");
      return { 
        effectifsCibles: new Map(),
        optionsParClasse: new Map(),
        maxElevesParClasse: 30,
        equilibrageActif: true 
      };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { 
        effectifsCibles: new Map(),
        optionsParClasse: new Map(),
        maxElevesParClasse: 30,
        equilibrageActif: true 
      };
    }
    
    // Identifier les colonnes
    const headers = data[0].map(h => String(h).toUpperCase());
    const colClasseOrigine = headers.indexOf("CLASSE_ORIGINE");
    const colClasseDest = headers.indexOf("CLASSE_DEST");
    const colEffectif = headers.indexOf("EFFECTIF");
    const colOptions = headers.indexOf("OPTIONS");
    
    const contraintes = {
      effectifsCibles: new Map(),
      optionsParClasse: new Map(),
      maxElevesParClasse: 30,
      equilibrageActif: true
    };
    
    // Parcourir les données pour extraire les informations
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const classeDest = colClasseDest !== -1 ? String(row[colClasseDest] || "").trim() : "";
      const effectif = colEffectif !== -1 ? Number(row[colEffectif]) || 28 : 28;
      const options = colOptions !== -1 ? String(row[colOptions] || "").trim() : "";
      
      if (classeDest) {
        // Stocker l'effectif cible
        contraintes.effectifsCibles.set(classeDest, effectif);
        contraintes.maxElevesParClasse = Math.max(contraintes.maxElevesParClasse, effectif);
        
        // Stocker les options disponibles
        if (options) {
          const optionsArray = options.split(/[,;=]/).map(o => {
            // Extraire le nom de l'option (avant le =)
            const match = o.match(/^([A-Z]+)/);
            return match ? match[1].trim() : o.trim();
          }).filter(o => o);
          
          contraintes.optionsParClasse.set(classeDest, new Set(optionsArray));
        }
      }
    }
    
    Logger.log(`_STRUCTURE lu: ${contraintes.effectifsCibles.size} classes configurées`);
    
    return contraintes;
  } catch (e) {
    Logger.log(`Erreur lecture _STRUCTURE: ${e.message}`);
    return { 
      effectifsCibles: new Map(),
      optionsParClasse: new Map(),
      maxElevesParClasse: 30,
      equilibrageActif: true 
    };
  }
}

/**
 * Analyse les contraintes d'un groupe d'élèves
 */
function analyserContraintesGroupe(groupe, classesPourOption) {
  const analyse = {
    totalEleves: groupe.length,
    avecContraintes: 0,
    sansContraintes: 0,
    parLangue: new Map(),
    parOption: new Map()
  };
  
  groupe.forEach(eleve => {
    if (eleve.hasConstraints) {
      analyse.avecContraintes++;
      
      // Analyser LV2
      if (eleve.lv2 && eleve.lv2.toUpperCase() !== "ESP") {
        const lv2 = eleve.lv2.toUpperCase();
        analyse.parLangue.set(lv2, (analyse.parLangue.get(lv2) || 0) + 1);
      }
      
      // Analyser options
      eleve.options.forEach(opt => {
        if (opt) {
          const option = opt.toUpperCase();
          analyse.parOption.set(option, (analyse.parOption.get(option) || 0) + 1);
        }
      });
    } else {
      analyse.sansContraintes++;
    }
  });
  
  return analyse;
}

/**
 * Génère toutes les répartitions possibles pour un groupe
 */
function genererRepartitionsPossibles(groupe, classesPossibles, classesPourOption, structureConstraints) {
  const repartitions = [];
  
  // Fonction récursive pour générer les combinaisons
  function genererRecursif(index, repartitionCourante, classesUtilisees) {
    if (index === groupe.length) {
      // Vérifier que chaque élève est dans une classe différente
      const elevesParClasse = new Map();
      repartitionCourante.forEach(({eleve, classe}) => {
        if (!elevesParClasse.has(classe)) {
          elevesParClasse.set(classe, []);
        }
        elevesParClasse.get(classe).push(eleve);
      });
      
      // Vérifier qu'il n'y a pas plus d'un élève du même code DISSO par classe
      let valide = true;
      elevesParClasse.forEach(elevesClasse => {
        if (elevesClasse.length > 1) {
          valide = false;
        }
      });
      
      if (valide) {
        repartitions.push([...repartitionCourante]);
      }
      return;
    }
    
    const eleve = groupe[index];
    
    // Trouver les classes compatibles pour cet élève
    classesPossibles.forEach(classe => {
      // Vérifier que la classe n'a pas déjà un élève de ce groupe
      if (!classesUtilisees.has(classe) && classeCompatible(eleve, classe, classesPourOption)) {
        repartitionCourante.push({eleve, classe});
        classesUtilisees.add(classe);
        
        genererRecursif(index + 1, repartitionCourante, classesUtilisees);
        
        repartitionCourante.pop();
        classesUtilisees.delete(classe);
      }
    });
  }
  
  // Limiter le nombre de répartitions pour éviter l'explosion combinatoire
  if (groupe.length > 6) {
    // Pour les grands groupes, utiliser une heuristique
    return genererRepartitionsHeuristique(groupe, classesPossibles, classesPourOption);
  }
  
  genererRecursif(0, [], new Set());
  
  return repartitions;
}

/**
 * Génère des répartitions par heuristique pour les grands groupes
 */
function genererRepartitionsHeuristique(groupe, classesPossibles, classesPourOption) {
  const repartitions = [];
  
  // Stratégie 1: Placer d'abord les élèves avec le plus de contraintes
  const groupeTrie = [...groupe].sort((a, b) => {
    const contraintesA = (a.lv2 && a.lv2.toUpperCase() !== "ESP" ? 1 : 0) + a.options.length;
    const contraintesB = (b.lv2 && b.lv2.toUpperCase() !== "ESP" ? 1 : 0) + b.options.length;
    return contraintesB - contraintesA;
  });
  
  // Générer quelques répartitions candidates
  for (let tentative = 0; tentative < 5; tentative++) {
    const repartition = [];
    const classesUtilisees = new Set();
    let valide = true;
    
    for (const eleve of groupeTrie) {
      const classesCompatibles = classesPossibles.filter(c => 
        !classesUtilisees.has(c) && classeCompatible(eleve, c, classesPourOption)
      );
      
      if (classesCompatibles.length === 0) {
        valide = false;
        break;
      }
      
      // Choisir une classe (avec une part d'aléatoire pour diversifier)
      const indexClasse = tentative === 0 ? 0 : Math.floor(Math.random() * classesCompatibles.length);
      const classeChoisie = classesCompatibles[indexClasse];
      
      repartition.push({eleve, classe: classeChoisie});
      classesUtilisees.add(classeChoisie);
    }
    
    if (valide) {
      repartitions.push(repartition);
    }
  }
  
  return repartitions;
}

/**
 * Évalue la qualité d'une répartition
 */
function evaluerRepartition(repartition, tousLesEleves, structureConstraints, classesPourOption) {
  let score = 100; // Score de base
  
  // Compter les élèves par classe après répartition
  const effectifsParClasse = new Map();
  
  // Compter les élèves déjà placés
  tousLesEleves.forEach(e => {
    const classe = e.planifie ? e.classeFinale : e.classeActuelle;
    if (classe) {
      effectifsParClasse.set(classe, (effectifsParClasse.get(classe) || 0) + 1);
    }
  });
  
  // Ajouter les élèves de la répartition candidate
  repartition.forEach(({eleve, classe}) => {
    // Retirer l'élève de son ancienne classe si nécessaire
    const ancienneClasse = eleve.planifie ? eleve.classeFinale : eleve.classeActuelle;
    if (ancienneClasse && ancienneClasse !== classe) {
      effectifsParClasse.set(ancienneClasse, (effectifsParClasse.get(ancienneClasse) || 1) - 1);
    }
    
    effectifsParClasse.set(classe, (effectifsParClasse.get(classe) || 0) + 1);
  });
  
  // Évaluer par rapport aux effectifs cibles de _STRUCTURE
  effectifsParClasse.forEach((effectif, classe) => {
    const effectifCible = structureConstraints.effectifsCibles.get(classe) || 28;
    
    // Pénaliser l'écart par rapport à l'effectif cible
    const ecart = Math.abs(effectif - effectifCible);
    score -= ecart * 5;
    
    // Pénaliser fortement si on dépasse l'effectif cible
    if (effectif > effectifCible) {
      score -= (effectif - effectifCible) * 15; // Pénalité supplémentaire
    }
  });
  
  // Pénaliser les déséquilibres d'effectifs si l'équilibrage est actif
  if (structureConstraints.equilibrageActif) {
    const effectifs = Array.from(effectifsParClasse.values());
    const moyenneEffectifs = effectifs.reduce((a, b) => a + b, 0) / effectifs.length;
    const ecartType = Math.sqrt(
      effectifs.reduce((sum, eff) => sum + Math.pow(eff - moyenneEffectifs, 2), 0) / effectifs.length
    );
    
    score -= ecartType * 10; // Plus l'écart-type est grand, plus on pénalise
  }
  
  // Bonus pour les élèves qui restent dans leur classe actuelle
  repartition.forEach(({eleve, classe}) => {
    const classeActuelle = eleve.planifie ? eleve.classeFinale : eleve.classeActuelle;
    if (classeActuelle === classe) {
      score += 5; // Bonus stabilité
    }
  });
  
  // Bonus pour respecter les préférences des élèves avec contraintes fortes
  repartition.forEach(({eleve, classe}) => {
    if (eleve.hasConstraints) {
      // Vérifier si c'est une des meilleures classes pour ses contraintes
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const classesPossibles = ss.getSheets()
        .filter(s => s.getName().match(/^[\d°]+\d*TEST$/))
        .map(s => s.getName());
      
      const nbClassesCompatibles = classesPossibles.filter(c => 
        classeCompatible(eleve, c, classesPourOption)
      ).length;
      
      if (nbClassesCompatibles <= 2) {
        score += 10; // Bonus pour respecter les contraintes fortes
      }
    }
  });
  
  return score;
}

/**
 * Vérifie s'il reste des conflits après optimisation
 */
function verifierConflitsResiduels(eleves) {
  const conflits = [];
  const elevesParClasseEtCode = new Map();
  
  eleves.forEach(e => {
    if (e.dissoCode) {
      const classe = e.planifie ? e.classeFinale : e.classeActuelle;
      if (classe) {
        const cle = `${classe}-${e.dissoCode}`;
        if (!elevesParClasseEtCode.has(cle)) {
          elevesParClasseEtCode.set(cle, []);
        }
        elevesParClasseEtCode.get(cle).push(e);
      }
    }
  });
  
  elevesParClasseEtCode.forEach((groupe, cle) => {
    if (groupe.length > 1) {
      conflits.push({
        cle: cle,
        eleves: groupe
      });
    }
  });
  
  return conflits;
}

/**
 * Corrige les conflits résiduels
 */
function corrigerConflitsResiduels(conflits, tousLesEleves, classesPourOption, structureConstraints) {
  let corrections = 0;
  
  conflits.forEach(conflit => {
    Logger.log(`\nCorrection conflit résiduel: ${conflit.cle}`);
    
    // Trier les élèves par priorité (fixes en premier)
    conflit.eleves.sort((a, b) => {
      if (a.isFixed && !b.isFixed) return -1;
      if (!a.isFixed && b.isFixed) return 1;
      if (a.hasConstraints && !b.hasConstraints) return -1;
      if (!a.hasConstraints && b.hasConstraints) return 1;
      return 0;
    });
    
    // Déplacer tous sauf le premier
    for (let i = 1; i < conflit.eleves.length; i++) {
      const eleve = conflit.eleves[i];
      
      // Trouver une classe alternative
      const classeAlternative = trouverMeilleureClasseAlternative(
        eleve, 
        tousLesEleves, 
        classesPourOption, 
        structureConstraints
      );
      
      if (classeAlternative) {
        eleve.planifie = true;
        eleve.classeFinale = classeAlternative;
        eleve.raison = `Correction conflit résiduel ${eleve.dissoCode}`;
        corrections++;
        Logger.log(`  ✓ ${eleve.id} → ${classeAlternative}`);
      } else {
        Logger.log(`  ❌ Impossible de déplacer ${eleve.id}`);
      }
    }
  });
  
  return corrections;
}

/**
 * Trouve la meilleure classe alternative pour un élève
 */
function trouverMeilleureClasseAlternative(eleve, tousLesEleves, classesPourOption, structureConstraints) {
  const classeActuelle = eleve.planifie ? eleve.classeFinale : eleve.classeActuelle;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classesPossibles = ss.getSheets()
    .filter(s => s.getName().match(/^[\d°]+\d*TEST$/))
    .map(s => s.getName())
    .filter(c => c !== classeActuelle);
  
  // Filtrer les classes compatibles
  const classesCompatibles = classesPossibles.filter(classe => {
    // Vérifier la compatibilité avec les contraintes
    if (!classeCompatible(eleve, classe, classesPourOption)) {
      return false;
    }
    
    // Vérifier qu'il n'y a pas déjà un élève avec le même code DISSO
    const dejaPresent = tousLesEleves.some(e => {
      const classeE = e.planifie ? e.classeFinale : e.classeActuelle;
      return classeE === classe && e.dissoCode === eleve.dissoCode && e.id !== eleve.id;
    });
    
    return !dejaPresent;
  });
  
  if (classesCompatibles.length === 0) {
    return null;
  }
  
  // Évaluer chaque classe selon plusieurs critères
  let meilleureClasse = null;
  let meilleurScore = -Infinity;
  
  classesCompatibles.forEach(classe => {
    const nbEleves = tousLesEleves.filter(e => {
      const classeE = e.planifie ? e.classeFinale : e.classeActuelle;
      return classeE === classe;
    }).length;
    
    const effectifCible = structureConstraints.effectifsCibles.get(classe) || 28;
    
    // Calculer un score basé sur plusieurs critères
    let score = 100;
    
    // Pénaliser si on dépasse l'effectif cible
    if (nbEleves >= effectifCible) {
      score -= (nbEleves - effectifCible + 1) * 20;
    } else {
      // Bonus pour les classes qui n'ont pas atteint leur effectif cible
      score += (effectifCible - nbEleves) * 5;
    }
    
    // Vérifier si l'élève a des options compatibles avec cette classe
    if (structureConstraints.optionsParClasse.has(classe)) {
      const optionsClasse = structureConstraints.optionsParClasse.get(classe);
      const optionsEleve = new Set([
        ...(eleve.lv2 && eleve.lv2.toUpperCase() !== "ESP" ? [eleve.lv2.toUpperCase()] : []),
        ...eleve.options.map(o => o.toUpperCase())
      ]);
      
      // Bonus si les options de l'élève correspondent aux options de la classe
      optionsEleve.forEach(opt => {
        if (optionsClasse.has(opt)) {
          score += 10;
        }
      });
    }
    
    if (score > meilleurScore) {
      meilleurScore = score;
      meilleureClasse = classe;
    }
  });
  
  return meilleureClasse;
}

////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////

/**
 * Applique les règles ASSO (association) - VERSION ORIGINALE SIMPLIFIÉE
 */
function appliquerReglesAsso(eleves, languesDetectees, optionsDetectees, classesPourOption) {
  Logger.log("Application des règles ASSO...");
  let deplacements = 0;
  let impossibles = 0;
  
  // Trouver tous les onglets TEST disponibles
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classesPossibles = ss.getSheets()
    .filter(s => s.getName().match(/^[\d°]+\d*TEST$/))
    .map(s => s.getName());
  
  // Extraire tous les codes ASSO uniques
  const codesAsso = new Set();
  eleves.forEach(e => {
    if (e.assoCode) codesAsso.add(e.assoCode);
  });
  
  Logger.log(`${codesAsso.size} codes ASSO trouvés`);
  
  // Pour chaque code ASSO
  for (const code of codesAsso) {
    try {
      // Trouver tous les élèves avec ce code
      const groupeAsso = eleves.filter(e => e.assoCode === code);
      
      // Si moins de 2 élèves, rien à faire
      if (groupeAsso.length < 2) continue;
      
      Logger.log(`Code ASSO ${code}: ${groupeAsso.length} élèves`);
      
      // Déterminer les classes possibles pour chaque élève
      const classesPossiblesParEleve = {};
      
      groupeAsso.forEach(eleve => {
        classesPossiblesParEleve[eleve.id] = new Set();
        
        if (eleve.hasConstraints) {
          classesPossibles.forEach(classe => {
            if (classeCompatible(eleve, classe, classesPourOption)) {
              classesPossiblesParEleve[eleve.id].add(classe);
            }
          });
        } else {
          classesPossibles.forEach(classe => {
            classesPossiblesParEleve[eleve.id].add(classe);
          });
        }
      });
      
      // Trouver l'intersection des classes possibles pour tous les élèves
      let classesCommunesPossibles = new Set(classesPossibles);
      
      groupeAsso.forEach(eleve => {
        const possibilites = classesPossiblesParEleve[eleve.id];
        classesCommunesPossibles = new Set(
          [...classesCommunesPossibles].filter(classe => possibilites.has(classe))
        );
      });
      
      Logger.log(`Classes communes possibles pour code ASSO ${code}: ${Array.from(classesCommunesPossibles).join(', ')}`);
      
      // Cas où aucune classe commune n'est possible
      if (classesCommunesPossibles.size === 0) {
        impossibles++;
        Logger.log(`❌ IMPOSSIBLE d'associer les élèves avec code ${code} : contraintes incompatibles`);
        continue;
      }
      
      // Choisir la première classe commune possible (logique simplifiée)
      const classeChoisie = Array.from(classesCommunesPossibles)[0];
      
      // Appliquer l'association
      for (const eleve of groupeAsso) {
        if (!eleve.place || eleve.classeActuelle !== classeChoisie) {
          eleve.planifie = true;
          eleve.classeFinale = classeChoisie;
          eleve.raison = `Associé (${code})`;
          deplacements++;
          Logger.log(`✓ ${eleve.id} → ${classeChoisie}`);
        }
      }
    } catch (error) {
      Logger.log(`❌ ERREUR lors du traitement du code ASSO ${code}: ${error.message}`);
    }
  }
  
  Logger.log(`Application ASSO terminée: ${deplacements} élèves à déplacer, ${impossibles} groupes impossibles`);
  return { deplacements, impossibles };
}

/**
 * Applique les règles DISSO (dissociation) - VERSION SIMPLIFIÉE
 * Ne fait qu'un traitement basique, la vérification finale corrigera les problèmes
 */
function appliquerReglesDisso(eleves, languesDetectees, optionsDetectees, classesPourOption) {
  Logger.log("Application des règles DISSO (traitement basique)...");
  let deplacements = 0;
  
  // Trouver tous les onglets TEST disponibles
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classesPossibles = ss.getSheets()
    .filter(s => s.getName().match(/^[\d°]+\d*TEST$/))
    .map(s => s.getName());
  
  // Extraire tous les codes DISSO uniques
  const codesDisso = new Set();
  eleves.forEach(e => {
    if (e.dissoCode) codesDisso.add(e.dissoCode);
  });
  
  Logger.log(`${codesDisso.size} codes DISSO trouvés (traitement basique)`);
  
  // Pour chaque code DISSO, faire un traitement simple
  for (const code of codesDisso) {
    const groupeDisso = eleves.filter(e => e.dissoCode === code);
    
    if (groupeDisso.length < 2) continue;
    
    Logger.log(`Code DISSO ${code}: ${groupeDisso.length} élèves - traitement basique`);
    
    // Distribution simple round-robin dans les classes compatibles
    let indexClasse = 0;
    
    groupeDisso.forEach(eleve => {
      if (!eleve.place) {
        // Trouver une classe compatible
        let classeChoisie = null;
        
        for (let i = 0; i < classesPossibles.length; i++) {
          const classe = classesPossibles[(indexClasse + i) % classesPossibles.length];
          if (classeCompatible(eleve, classe, classesPourOption)) {
            classeChoisie = classe;
            break;
          }
        }
        
        if (classeChoisie) {
          eleve.planifie = true;
          eleve.classeFinale = classeChoisie;
          eleve.raison = `Dissocié (${code}) - basique`;
          deplacements++;
          indexClasse = (indexClasse + 1) % classesPossibles.length;
          Logger.log(`✓ ${eleve.id} → ${classeChoisie} (basique)`);
        }
      }
    });
  }
  
  Logger.log(`Application DISSO basique terminée: ${deplacements} élèves traités`);
  return { deplacements, impossibles: 0 };
}

// [Fonctions de lecture et fusion restent identiques]

/**
 * Lit tous les élèves déjà placés dans les onglets TEST (résultat Phase 1)
 */
function lireElevesDesOngletsTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eleves = [];
  
  // Trouver tous les onglets TEST
  const sheets = ss.getSheets();
  const testSheets = sheets.filter(sheet => sheet.getName().match(/^[\d°]+\d*TEST$/));
  
  Logger.log(`${testSheets.length} onglets TEST trouvés`);
  
  // Pour chaque onglet TEST
  testSheets.forEach(sheet => {
    const nomClasse = sheet.getName();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return;
    
    const headers = data[0].map(h => String(h).trim());
    const indices = {
      id: headers.indexOf("ID_ELEVE"),
      lv2: headers.indexOf("LV2"),
      opt: headers.indexOf("OPT"),
      nom: headers.indexOf("NOM"),
      prenom: headers.indexOf("PRENOM"),
      nomPrenom: headers.indexOf("NOM & PRENOM") !== -1 ? headers.indexOf("NOM & PRENOM") : headers.indexOf("NOM_PRENOM"),
      sexe: headers.indexOf("SEXE"),
      com: headers.indexOf("COM"),
      tra: headers.indexOf("TRA"),
      part: headers.indexOf("PART"),
      abs: headers.indexOf("ABS"),
      dispo: headers.indexOf("DISPO"),
      asso: headers.indexOf("ASSO"),
      disso: headers.indexOf("DISSO"),
      source: headers.indexOf("SOURCE"),
      fixe: headers.indexOf("FIXE"),
      classeFinale: headers.indexOf("CLASSE_FINALE")
    };
    
    if (indices.id === -1) return;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = String(row[indices.id] || "");
      if (!id) continue;
      
      const nom = indices.nom !== -1 ? String(row[indices.nom] || "") : "";
      const prenom = indices.prenom !== -1 ? String(row[indices.prenom] || "") : "";
      const nomPrenom = indices.nomPrenom !== -1 ? String(row[indices.nomPrenom] || "") : `${nom} ${prenom}`.trim();
      const sexe = indices.sexe !== -1 ? String(row[indices.sexe] || "") : "";
      const lv2 = indices.lv2 !== -1 ? String(row[indices.lv2] || "").trim() : "";
      const optStr = indices.opt !== -1 ? String(row[indices.opt] || "").trim() : "";
      const fixeValeur = indices.fixe !== -1 ? String(row[indices.fixe] || "").trim().toUpperCase() : "";
      const assoCode = indices.asso !== -1 ? String(row[indices.asso] || "").trim() : "";
      const dissoCode = indices.disso !== -1 ? String(row[indices.disso] || "").trim().toUpperCase() : "";
      const source = indices.source !== -1 ? String(row[indices.source] || "").trim() : "";
      const classeFinale = indices.classeFinale !== -1 ? String(row[indices.classeFinale] || "").trim() : "";
      const com = indices.com !== -1 ? row[indices.com] || "" : "";
      const tra = indices.tra !== -1 ? row[indices.tra] || "" : "";
      const part = indices.part !== -1 ? row[indices.part] || "" : "";
      const abs = indices.abs !== -1 ? row[indices.abs] || "" : "";
      const dispo = indices.dispo !== -1 ? row[indices.dispo] || "" : "";
      
      const options = optStr ? optStr.split(",").map(o => o.trim()).filter(o => o) : [];
      
      const eleve = {
        id: id,
        nom: nom,
        prenom: prenom,
        nomComplet: nomPrenom,
        sexe: sexe,
        classeActuelle: nomClasse,
        place: true,
        lv2: lv2,
        options: options,
        com: com,
        tra: tra,
        part: part,
        abs: abs,
        dispo: dispo,
        assoCode: assoCode,
        dissoCode: dissoCode,
        source: source,
        fixe: fixeValeur,
        classeFinale: classeFinale,
        rangee: row,
        indices: indices
      };
      
      eleves.push(eleve);
    }
  });
  
  return eleves;
}

/**
 * Lit tous les élèves depuis CONSOLIDATION
 */
function lireTousLesElevesDepuisConsolidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CONSOLIDATION");
  
  if (!sheet) {
    throw new Error("Onglet CONSOLIDATION introuvable");
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    throw new Error("Onglet CONSOLIDATION vide ou avec seulement des en-têtes");
  }
  
  const headers = data[0].map(h => String(h).trim());
  const indices = {
    id: headers.indexOf("ID_ELEVE"),
    nom: headers.indexOf("NOM"),
    prenom: headers.indexOf("PRENOM"),
    nomPrenom: headers.indexOf("NOM & PRENOM") !== -1 ? headers.indexOf("NOM & PRENOM") : headers.indexOf("NOM_PRENOM"),
    sexe: headers.indexOf("SEXE"),
    lv2: headers.indexOf("LV2"),
    opt: headers.indexOf("OPT"),
    com: headers.indexOf("COM"),
    tra: headers.indexOf("TRA"),
    part: headers.indexOf("PART"),
    abs: headers.indexOf("ABS"),
    dispo: headers.indexOf("DISPO"),
    asso: headers.indexOf("ASSO"),
    disso: headers.indexOf("DISSO"),
    source: headers.indexOf("SOURCE"),
    fixe: headers.indexOf("FIXE"),
    classeFinale: headers.indexOf("CLASSE_FINALE")
  };
  
  if (indices.id === -1) {
    throw new Error("Colonne ID manquante dans CONSOLIDATION");
  }
  
  const eleves = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = String(row[indices.id] || "");
    if (!id) continue;
    
    const nom = indices.nom !== -1 ? String(row[indices.nom] || "") : "";
    const prenom = indices.prenom !== -1 ? String(row[indices.prenom] || "") : "";
    const nomPrenom = indices.nomPrenom !== -1 ? String(row[indices.nomPrenom] || "") : `${nom} ${prenom}`.trim();
    const sexe = indices.sexe !== -1 ? String(row[indices.sexe] || "") : "";
    const lv2 = indices.lv2 !== -1 ? String(row[indices.lv2] || "").trim() : "";
    const optStr = indices.opt !== -1 ? String(row[indices.opt] || "").trim() : "";
    const assoCode = indices.asso !== -1 ? String(row[indices.asso] || "").trim() : "";
    const dissoCode = indices.disso !== -1 ? String(row[indices.disso] || "").trim().toUpperCase() : "";
    const source = indices.source !== -1 ? String(row[indices.source] || "").trim() : "";
    const fixeValeur = indices.fixe !== -1 ? String(row[indices.fixe] || "").trim().toUpperCase() : "";
    const classeFinale = indices.classeFinale !== -1 ? String(row[indices.classeFinale] || "").trim() : "";
    const com = indices.com !== -1 ? row[indices.com] || "" : "";
    const tra = indices.tra !== -1 ? row[indices.tra] || "" : "";
    const part = indices.part !== -1 ? row[indices.part] || "" : "";
    const abs = indices.abs !== -1 ? row[indices.abs] || "" : "";
    const dispo = indices.dispo !== -1 ? row[indices.dispo] || "" : "";
    
    const options = optStr ? optStr.split(",").map(o => o.trim()).filter(o => o) : [];
    
    const eleve = {
      id: id,
      nom: nom,
      prenom: prenom,
      nomComplet: nomPrenom,
      sexe: sexe,
      lv2: lv2,
      options: options,
      com: com,
      tra: tra,
      part: part,
      abs: abs,
      dispo: dispo,
      assoCode: assoCode,
      dissoCode: dissoCode.startsWith('D') ? dissoCode : "",
      source: source,
      fixe: fixeValeur,
      classeFinale: classeFinale,
      place: false,
      classeActuelle: null,
      planifie: false,
      raison: ""
    };
    
    eleves.push(eleve);
  }
  
  return eleves;
}

/**
 * Fusionne les données des élèves de CONSOLIDATION avec ceux déjà placés
 */
function fusionnerDonnees(tousLesEleves, elevesPlaces) {
  const placesParId = {};
  elevesPlaces.forEach(e => {
    placesParId[e.id] = e;
  });
  
  const languesDetectees = new Set();
  const optionsDetectees = new Set();
  const classesPourOption = {};
  
  tousLesEleves.forEach(e => {
    if (e.lv2) languesDetectees.add(e.lv2.toUpperCase());
    e.options.forEach(opt => optionsDetectees.add(opt.toUpperCase()));
    
    if (placesParId[e.id]) {
      const place = placesParId[e.id];
      e.place = true;
      e.classeActuelle = place.classeActuelle;
      e.classeFinale = place.classeActuelle;
      e.rangee = place.rangee;
      e.indices = place.indices;
      
      if (place.fixe) e.fixe = place.fixe;
      if (place.classeFinale) e.classeFinale = place.classeFinale;
      
      if (!e.com && place.com) e.com = place.com;
      if (!e.tra && place.tra) e.tra = place.tra;
      if (!e.part && place.part) e.part = place.part;
      if (!e.abs && place.abs) e.abs = place.abs;
      if (!e.dispo && place.dispo) e.dispo = place.dispo;
      if (!e.source && place.source) e.source = place.source;
      
      if (e.lv2 && e.lv2.toUpperCase() !== "ESP" && e.lv2.toUpperCase() !== "") {
        const langue = e.lv2.toUpperCase();
        if (!classesPourOption[langue]) classesPourOption[langue] = new Set();
        classesPourOption[langue].add(e.classeActuelle);
      }
      
      e.options.forEach(opt => {
        if (opt) {
          const option = opt.toUpperCase();
          if (!classesPourOption[option]) classesPourOption[option] = new Set();
          classesPourOption[option].add(e.classeActuelle);
        }
      });
    }
  });
  
  const idsConnus = new Set(tousLesEleves.map(e => e.id));
  elevesPlaces.forEach(e => {
    if (!idsConnus.has(e.id)) {
      tousLesEleves.push(e);
      if (e.lv2) languesDetectees.add(e.lv2.toUpperCase());
      (e.options || []).forEach(opt => optionsDetectees.add(opt.toUpperCase()));
    }
  });
  
  tousLesEleves.forEach(e => {
    e.hasOptions = e.options.length > 0;
    e.hasConstraints = aContraintesFortes(e);
    
    if (e.fixe === "OUI" || e.fixe === "PREF") {
      e.isFixed = true;
    }
  });
  
  Logger.log(`Fusion terminée: ${tousLesEleves.length} élèves au total`);
  
  return { 
    eleves: tousLesEleves,
    languesDetectees,
    optionsDetectees,
    classesPourOption
  };
}

/**
 * Écrit les résultats dans les onglets TEST
 */
function ecrireResultatsDansOngletsTest(eleves) {
  Logger.log("Écriture des résultats dans les onglets TEST...");
  
  const HEADERS_STANDARD = [
    "ID_ELEVE", "NOM", "PRENOM", "NOM & PRENOM", "SEXE", "LV2", "OPT", 
    "COM", "TRA", "PART", "ABS", "DISPO", "ASSO", "DISSO", 
    "SOURCE", "FIXE", "CLASSE_FINALE"
  ];
  
  const elevesParClasse = new Map();
  
  eleves.filter(e => e.place && !e.planifie).forEach(e => {
    if (!elevesParClasse.has(e.classeActuelle)) {
      elevesParClasse.set(e.classeActuelle, []);
    }
    elevesParClasse.get(e.classeActuelle).push(e);
  });
  
  eleves.filter(e => e.planifie).forEach(e => {
    if (!elevesParClasse.has(e.classeFinale)) {
      elevesParClasse.set(e.classeFinale, []);
    }
    elevesParClasse.get(e.classeFinale).push(e);
  });
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalEcrits = 0;
  let erreurs = 0;
  
  elevesParClasse.forEach((elevesClasse, nomClasse) => {
    try {
      const sheet = ss.getSheetByName(nomClasse);
      if (!sheet) {
        Logger.log(`ERREUR: Onglet ${nomClasse} introuvable`);
        erreurs++;
        return;
      }
      
      const headersExistants = sheet.getRange(1, 1, 1, HEADERS_STANDARD.length).getValues()[0]
        .map(h => String(h).trim());
      
      let headersDifferents = false;
      for (let i = 0; i < HEADERS_STANDARD.length; i++) {
        if (i >= headersExistants.length || headersExistants[i] !== HEADERS_STANDARD[i]) {
          headersDifferents = true;
          break;
        }
      }
      
      if (headersDifferents) {
        sheet.getRange(1, 1, 1, HEADERS_STANDARD.length).setValues([HEADERS_STANDARD]);
        sheet.getRange(1, 1, 1, HEADERS_STANDARD.length).setBackground("#f3f3f3").setFontWeight("bold");
      }
      
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
      }
      
      elevesClasse.sort((a, b) => {
        const nomA = a.nomComplet || `${a.nom} ${a.prenom}`;
        const nomB = b.nomComplet || `${b.nom} ${b.prenom}`;
        return nomA.localeCompare(nomB);
      });
      
      const donnees = [];
      elevesClasse.forEach(e => {
        const ligne = new Array(HEADERS_STANDARD.length).fill("");
        
        ligne[0] = e.id || "";
        ligne[1] = e.nom || "";
        ligne[2] = e.prenom || "";
        ligne[3] = e.nomComplet || `${e.nom} ${e.prenom}`.trim();
        ligne[4] = e.sexe || "";
        ligne[5] = e.lv2 || "";
        ligne[6] = Array.isArray(e.options) ? e.options.join(",") : e.options || "";
        ligne[7] = e.com || "";
        ligne[8] = e.tra || "";
        ligne[9] = e.part || "";
        ligne[10] = e.abs || "";
        ligne[11] = e.dispo || "";
        ligne[12] = e.assoCode || "";
        ligne[13] = e.dissoCode || "";
        ligne[14] = e.source || "";
        ligne[15] = e.fixe || "";
        ligne[16] = e.raison || "";
        
        donnees.push(ligne);
      });
      
      if (donnees.length > 0) {
        sheet.getRange(2, 1, donnees.length, HEADERS_STANDARD.length).setValues(donnees);
        totalEcrits += donnees.length;
        Logger.log(`✓ ${donnees.length} élèves écrits dans ${nomClasse}`);
      }
    } catch (error) {
      Logger.log(`ERREUR lors de l'écriture dans ${nomClasse}: ${error.message}`);
      erreurs++;
    }
  });
  
  return { 
    success: erreurs === 0,
    totalEcrits: totalEcrits,
    erreurs: erreurs
  };
}