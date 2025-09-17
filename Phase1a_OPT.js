// Bandeau maintenance : tous les utilitaires sont centralisés dans Utils.js
// Utilisez uniquement Utils.idx(header, name) et Utils.logAction(...)

/**
 * PHASE 1 - RÉPARTITION DES OPTIONS (VERSION OPTIMISÉE)
 * Version qui lit depuis CONSOLIDATION et répartit dans les onglets TEST
 * Solution centralisée et adaptative qui s'ajuste à toutes les langues et options
 */
function repartirPhase1() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const structureSheet = ss.getSheetByName("_STRUCTURE");
    const consolidationSheet = ss.getSheetByName("CONSOLIDATION");
    
    if (!structureSheet) throw new Error("Onglet _STRUCTURE introuvable");
    if (!consolidationSheet) throw new Error("Onglet CONSOLIDATION introuvable");
    
    // Forcer un rafraîchissement complet du classeur pour être sûr de lire des données à jour
    SpreadsheetApp.flush();
    
    // 1. Lire la structure depuis _STRUCTURE
    const structure = lireStructure(structureSheet);
    if (!structure || structure.classes.length === 0) {
      throw new Error("Structure non valide ou vide");
    }
    
    // 2. Lire tous les élèves depuis CONSOLIDATION
    const tousLesEleves = lireDonneesDepuisConsolidation(consolidationSheet, structure);
    
    // 3. Effacer et recréer tous les onglets TEST nécessaires
    const ongletsTestCrees = preparerOngletsTest(structure, consolidationSheet);
    
    // 4. Répartir les élèves dans les onglets TEST
    const resultat = repartirElevesDansLesClasses(tousLesEleves, structure);
    
    // 5. Afficher résultat
    const message = `Phase 1 terminée: ${resultat.total} élèves répartis (${resultat.detailsType.LV2} langues, ${resultat.detailsType.OPT} options)`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Phase 1", 10);
    
    return { 
      success: true, 
      message: message,
      details: resultat.details 
    };
  } catch (e) {
    const message = "Erreur: " + e.message;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Phase 1", 10);
    return { success: false, message: message };
  }
}

/**
 * Lit la structure de _STRUCTURE
 */
function lireStructure(sheet) {
  Logger.log("Début de lireStructure - Lecture de la feuille _STRUCTURE");
  try {
    // Vérifier que sheet est bien défini
    if (!sheet) {
      Logger.log("ERREUR: Le paramètre sheet est undefined ou null");
      throw new Error("Feuille _STRUCTURE non fournie");
    }
    
    // Lire toutes les données
    const data = sheet.getDataRange().getValues();
    Logger.log(`Données lues: ${data.length} lignes`);
    
    // Vérifier que data contient des données
    if (!data || data.length === 0) {
      Logger.log("ERREUR: Aucune donnée trouvée dans la feuille");
      throw new Error("Feuille _STRUCTURE vide");
    }
    
    // Trouver la ligne d'en-tête
    let headerRow = -1;
    for (let i = 0; i < Math.min(10, data.length); i++) {
      Logger.log(`Analyse ligne ${i}: [${data[i].join(', ')}]`);
      if (data[i][0] === "CLASSE_ORIGINE" && data[i][1] === "CLASSE_DEST") {
        headerRow = i;
        Logger.log(`En-têtes trouvés à la ligne ${i}`);
        break;
      }
    }
    
    if (headerRow === -1) {
      Logger.log("ERREUR: En-têtes CLASSE_ORIGINE et CLASSE_DEST non trouvés");
      throw new Error("En-têtes non trouvés dans _STRUCTURE");
    }
    
    // Trouver les indices des colonnes
    const headers = data[headerRow];
    Logger.log(`En-têtes: [${headers.join(', ')}]`);
    
    const colOrigine = headers.indexOf("CLASSE_ORIGINE");
    const colDest = headers.indexOf("CLASSE_DEST");
    const colEffectif = headers.indexOf("EFFECTIF");
    const colOptions = headers.indexOf("OPTIONS");
    
    Logger.log(`Indices des colonnes: ORIGINE=${colOrigine}, DEST=${colDest}, EFFECTIF=${colEffectif}, OPTIONS=${colOptions}`);
    
    if (colOrigine === -1) {
      Logger.log("ERREUR: Colonne CLASSE_ORIGINE non trouvée");
      throw new Error("Colonne CLASSE_ORIGINE non trouvée");
    }
    if (colDest === -1) {
      Logger.log("ERREUR: Colonne CLASSE_DEST non trouvée");
      throw new Error("Colonne CLASSE_DEST non trouvée");
    }
    if (colOptions === -1) {
      Logger.log("ERREUR: Colonne OPTIONS non trouvée");
      throw new Error("Colonne OPTIONS non trouvée");
    }
    
    // Structure de retour
    const structure = {
      classes: [],
      languesReconnues: new Set(), // Pour collecter dynamiquement les langues
      optionsReconnues: new Set()  // Pour collecter dynamiquement les options
    };
    
    // Parcourir les lignes de données
    Logger.log(`Analyse des données à partir de la ligne ${headerRow + 1}`);
    let lignesTraitees = 0;
    let classesAjoutees = 0;
    
    for (let i = headerRow + 1; i < data.length; i++) {
      const row = data[i];
      lignesTraitees++;
      
      // Vérifier que row contient des données
      if (!row || row.length <= Math.max(colOrigine, colDest, colOptions)) {
        Logger.log(`Ligne ${i} ignorée: données incomplètes`);
        continue;
      }
      
      const origine = row[colOrigine];
      const dest = row[colDest];
      
      if (!origine || !dest) {
        Logger.log(`Ligne ${i} ignorée: origine ou destination vide`);
        continue; // Ignorer les lignes incomplètes
      }
      
      // Créer l'objet classe
      const classe = {
        origine: String(origine).trim(),
        destination: String(dest).trim(),
        nomTest: String(dest).trim() + "TEST",
        effectif: colEffectif !== -1 ? (parseInt(row[colEffectif]) || 0) : 0,
        regles: [] // Règles de répartition (options, langues)
      };
      
      Logger.log(`Classe en cours de traitement: ${classe.origine} -> ${classe.destination} (effectif: ${classe.effectif})`);
      
      // Extraire les options
      if (row[colOptions]) {
        const optionsStr = String(row[colOptions]);
        Logger.log(`Options pour ${classe.destination}: ${optionsStr}`);
        
        try {
          const optionsPairs = optionsStr.split(',');
          Logger.log(`${optionsPairs.length} paires option=quota trouvées`);
          
          for (let j = 0; j < optionsPairs.length; j++) {
            const pair = optionsPairs[j];
            try {
              const parts = pair.split('=').map(p => p.trim());
              if (parts.length === 2) {
                const option = parts[0];
                const quota = parseInt(parts[1]) || 0;
                
                if (option && quota > 0) {
                  // On ne détermine pas encore le type, on l'identifiera après analyse
                  classe.regles.push({
                    option: option,
                    quota: quota,
                    type: null // Sera déterminé plus tard
                  });
                  
                  Logger.log(`Règle ajoutée: ${classe.destination} - ${option}=${quota}`);
                } else {
                  Logger.log(`Règle ignorée (option vide ou quota ≤ 0): ${pair}`);
                }
              } else {
                Logger.log(`Format de paire invalide (pas de '='): ${pair}`);
              }
            } catch (pairError) {
              Logger.log(`Erreur lors du traitement de la paire ${j}: ${pairError.message}`);
            }
          }
        } catch (optionsError) {
          Logger.log(`Erreur lors du traitement des options: ${optionsError.message}`);
        }
      } else {
        Logger.log(`Pas d'options définies pour ${classe.destination}`);
      }
      
      // N'ajouter que si la classe a des règles
      if (classe.regles.length > 0) {
        structure.classes.push(classe);
        classesAjoutees++;
        Logger.log(`Classe ${classe.destination} ajoutée avec ${classe.regles.length} règles`);
      } else {
        Logger.log(`Classe ${classe.destination} ignorée: aucune règle définie`);
      }
    }
    
    Logger.log(`Traitement terminé: ${lignesTraitees} lignes traitées, ${classesAjoutees} classes ajoutées`);
    Logger.log(`Structure finale: ${JSON.stringify(structure)}`);
    
    // Vérification finale
    if (structure.classes.length === 0) {
      Logger.log("AVERTISSEMENT: Aucune classe avec règles n'a été trouvée");
    }
    
    return structure;
  } catch (e) {
    Logger.log(`ERREUR CRITIQUE dans lireStructure: ${e.message}\n${e.stack}`);
    throw e; // Remonter l'erreur
  }
}

/**
 * Lit les données des élèves depuis l'onglet CONSOLIDATION
 */
function lireDonneesDepuisConsolidation(sheet, structure) {
  // Lire les données (rafraîchies)
  SpreadsheetApp.flush(); // Forcer le rafraîchissement
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    throw new Error("Onglet CONSOLIDATION vide ou avec seulement un en-tête");
  }
  
  // Trouver les colonnes importantes
  const headers = data[0].map(h => String(h).trim());
  const indices = {
    id: headers.indexOf("ID_ELEVE"),
    lv2: headers.indexOf("LV2"),
    opt: headers.indexOf("OPT"),
    source: headers.indexOf("SOURCE")
  };
  
  // Vérifier les colonnes essentielles
  if (indices.id === -1) throw new Error("Colonne ID_ELEVE manquante dans CONSOLIDATION");
  if (indices.lv2 === -1) Logger.log("ATTENTION: Colonne LV2 manquante dans CONSOLIDATION");
  if (indices.opt === -1) Logger.log("ATTENTION: Colonne OPT manquante dans CONSOLIDATION");
  
  // Collecter tous les élèves avec leurs données
  const eleves = [];
  
  // Pour collecter dynamiquement les langues et options existantes
  const languesUniques = new Set();
  const optionsUniques = new Set();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = indices.id !== -1 ? String(row[indices.id] || "") : "";
    
    if (!id) continue; // Ignorer les lignes sans ID
    
    const lv2 = indices.lv2 !== -1 ? String(row[indices.lv2] || "").trim() : "";
    const optStr = indices.opt !== -1 ? String(row[indices.opt] || "").trim() : "";
    const options = optStr ? optStr.split(",").map(o => o.trim()) : [];
    const source = indices.source !== -1 ? String(row[indices.source] || "").trim() : "";
    
    // Collecter les langues et options uniques
    if (lv2) languesUniques.add(lv2.toUpperCase());
    options.forEach(opt => {
      if (opt) optionsUniques.add(opt.toUpperCase());
    });
    
    // Ajouter l'élève avec sa ligne complète
    eleves.push({
      id: id,
      lv2: lv2,
      options: options,
      source: source,
      rangee: row // Ligne complète pour copier
    });
  }
  
  // Déterminer maintenant le type de chaque règle (LV2 ou OPT)
  structure.classes.forEach(classe => {
    classe.regles.forEach(regle => {
      const optionUpper = regle.option.toUpperCase();
      
      // Si l'option se trouve dans les langues uniques, c'est une LV2
      if (languesUniques.has(optionUpper)) {
        regle.type = "LV2";
        structure.languesReconnues.add(optionUpper);
      } else {
        regle.type = "OPT";
        structure.optionsReconnues.add(optionUpper);
      }
      
      Logger.log(`Type de règle déterminé: ${classe.destination} - ${regle.option}=${regle.quota} (type: ${regle.type})`);
    });
  });
  
  // Statistiques des élèves par langue et option
  const compteurLangues = {};
  const compteurOptions = {};
  
  eleves.forEach(e => {
    if (e.lv2) {
      compteurLangues[e.lv2] = (compteurLangues[e.lv2] || 0) + 1;
    }
    e.options.forEach(opt => {
      compteurOptions[opt] = (compteurOptions[opt] || 0) + 1;
    });
  });
  
  Logger.log(`Total: ${eleves.length} élèves - Langues: ${JSON.stringify(compteurLangues)} - Options: ${JSON.stringify(compteurOptions)}`);
  
  return eleves;
}

/**
 * Prépare les onglets TEST en les effaçant et recréant
 * Modifié pour créer aussi l'onglet 5°1TEST même sans règles d'options
 */
function preparerOngletsTest(structure, consolidationSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Récupérer l'en-tête depuis CONSOLIDATION
  const headers = consolidationSheet.getRange(1, 1, 1, consolidationSheet.getLastColumn()).getValues()[0];
  
  // Lire la structure directement depuis l'onglet _STRUCTURE pour trouver toutes les classes de destination
  const structureSheet = ss.getSheetByName("_STRUCTURE");
  const toutesLesClassesDest = new Set();
  
  if (structureSheet) {
    const data = structureSheet.getDataRange().getValues();
    // Trouver la ligne d'en-tête
    let headerRow = -1;
    for (let i = 0; i < Math.min(10, data.length); i++) {
      if (data[i][0] === "CLASSE_ORIGINE" && data[i][1] === "CLASSE_DEST") {
        headerRow = i;
        break;
      }
    }
    
    if (headerRow !== -1) {
      const colDest = data[headerRow].indexOf("CLASSE_DEST");
      if (colDest !== -1) {
        // Collecter toutes les classes de destination, même celles sans options
        for (let i = headerRow + 1; i < data.length; i++) {
          const dest = data[i][colDest];
          if (dest && typeof dest === 'string' && dest.trim()) {
            toutesLesClassesDest.add(dest.trim());
          }
        }
      }
    }
  }
  
  Logger.log(`Classes trouvées dans _STRUCTURE: ${Array.from(toutesLesClassesDest).join(", ")}`);
  
  // Créer des onglets TEST pour toutes les classes dans _STRUCTURE
  toutesLesClassesDest.forEach(dest => {
    const nomTest = dest + "TEST";
    let testSheet = ss.getSheetByName(nomTest);
    
    if (testSheet) {
      ss.deleteSheet(testSheet); // Effacer s'il existe déjà
    }
    
    // Créer un nouvel onglet TEST
    testSheet = ss.insertSheet(nomTest);
    
    // Copier l'en-tête
    testSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    testSheet.getRange(1, 1, 1, headers.length).setBackground("#f3f3f3").setFontWeight("bold");
    
    Logger.log(`Onglet ${nomTest} créé pour classe ${dest}`);
  });
  
  return true;
}

/**
 * Répartit les élèves dans les onglets TEST selon les règles définies
 * – version « multi-critères » –
 */
function repartirElevesDansLesClasses(tousLesEleves, structure) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  /* -------------------- 1) VARIABLES GLOBALES -------------------- */
  const elevesPlaces = new Set();          // pour ne poser chaque élève qu'une fois
  let   totalPlaces  = 0;                  // compteur global
  const statsType = { LV2: 0, OPT: 0 };    // stats par type de règle
  const details   = [];                    // reporting détaillé

  /* -------------------- 2) QUOTAS RESTANTS ----------------------- */
  structure.classes.forEach(classe =>
    classe.regles.forEach(r => r.restant = r.quota)
  );

  /* -------------------- 3) PRIORISATION DES CLASSES -------------- */
  const classesAvecPriorite = structure.classes
    .map(classe => ({
      ...classe,
      nombreCriteres: classe.regles.length          // combien de règles pour la classe
    }))
    .sort((a, b) => b.nombreCriteres - a.nombreCriteres);

  /* -------------------- 4) FONCTION DÉCRÉMENT QUOTAS ------------- */
  function decrementeQuotasEtStats(eleve, classe) {
    classe.regles.forEach(regle => {
      if (regle.restant <= 0) return;               // plus de place sur cette règle

      const opt = regle.option.toUpperCase();
      const match = regle.type === "LV2"
        ? eleve.lv2 && eleve.lv2.toUpperCase() === opt
        : eleve.options.some(o => o.toUpperCase() === opt);

      if (match) {
        regle.restant--;
        statsType[regle.type]++;                    // on incrémente LV2 ou OPT
        Logger.log(`   ↳ quota ${regle.type} ${opt} décrémenté → ${regle.restant} restants`);
      }
    });
  }

  /* -------------------- 5) RÉPARTITION --------------------------- */
  classesAvecPriorite.forEach(classe => {
    const sheet = ss.getSheetByName(classe.nomTest);
    if (!sheet) {
      Logger.log(`ERREUR : onglet ${classe.nomTest} introuvable`);
      return;
    }

    let ligne = 2;                                  // 1ʳᵉ ligne dispo après l'en-tête

    tousLesEleves.forEach(eleve => {
      if (elevesPlaces.has(eleve.id)) return;       // déjà casé ailleurs

      /* L'élève satisfait-il AU MOINS UNE règle qui a encore de la place ? */
      const remplitUneRegle = classe.regles.some(r => {
        if (r.restant <= 0) return false;
        const opt = r.option.toUpperCase();
        return r.type === "LV2"
          ? eleve.lv2 && eleve.lv2.toUpperCase() === opt
          : eleve.options.some(o => o.toUpperCase() === opt);
      });

      if (!remplitUneRegle) return;                 // on passe à l'élève suivant

      /* ---------- Placement effectif ---------- */
      sheet.getRange(ligne, 1, 1, eleve.rangee.length)
           .setValues([eleve.rangee]);
      elevesPlaces.add(eleve.id);
      ligne++;
      totalPlaces++;

      decrementeQuotasEtStats(eleve, classe);

      Logger.log(`✓ ${eleve.id} placé dans ${classe.nomTest} (LV2=${eleve.lv2}, OPT=${eleve.options.join(",")})`);
    });

    /* ---------- Reporting détaillé pour la classe ---------- */
    classe.regles.forEach(r => {
      details.push({
        classe : classe.destination,
        option : r.option,
        type   : r.type,
        quota  : r.quota,
        places : r.quota - r.restant
      });
    });
  });

  /* -------------------- 6) RETOUR DES STATISTIQUES --------------- */
  return {
    total      : totalPlaces,
    detailsType: statsType,
    details    : details
  };
}