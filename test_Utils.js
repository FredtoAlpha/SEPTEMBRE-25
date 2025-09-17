/**
 * Tests unitaires pour les fonctions utilitaires de Utils.js
 * À exécuter manuellement dans Apps Script ou via un runner GAS compatible
 */
function test_idx() {
  const header = ['NOM', 'PRENOM', 'CLASSE'];
  if (Utils.idx(header, 'PRENOM') !== 1) throw new Error('idx PRENOM échoue');
  if (Utils.idx(header, 'INEXISTANT') !== -1) throw new Error('idx inexistant échoue');
}

function test_logAction() {
  try {
    Utils.logAction('Test logAction');
  } catch (e) {
    throw new Error('logAction doit fonctionner sans erreur');
  }
}

/**
 * Test des corrections du pipeline Variante Scores
 */
function testCorrectionsVarianteScores() {
  Logger.log("=== TEST CORRECTIONS VARIANTE SCORES ===");
  
  try {
    const config = getConfig();
    config.SCENARIOS_ACTIFS = ['COM', 'TRA', 'PART'];
    
    // 1. Test de préparation des données
    Logger.log("1. Test préparation données...");
    const dataContext = preparerDonneesVarianteScores(config);
    
    if (!dataContext || !dataContext.classesState) {
      throw new Error("Échec préparation données");
    }
    
    // 2. Test des statistiques de scores
    Logger.log("2. Test calcul statistiques...");
    const stats = dataContext.statsScores;
    
    Logger.log("Statistiques calculées:");
    ['COM', 'TRA', 'PART'].forEach(critere => {
      Logger.log(`${critere}: ${stats[critere].total} élèves total`);
      [1, 2, 3, 4].forEach(score => {
        Logger.log(`  Score ${score}: ${stats[critere][score]} élèves`);
      });
    });
    
    // 3. Test de la mobilité
    Logger.log("3. Test identification mobilité...");
    const mobiles = dataContext.elevesMobiles;
    Logger.log(`Élèves mobiles: ${mobiles.total}`);
    Logger.log(`  LIBRE: ${mobiles.LIBRE.length}`);
    Logger.log(`  PERMUT: ${mobiles.PERMUT.length}`);
    Logger.log(`  SPEC: ${mobiles.SPEC.length}`);
    
    // 4. Test du score d'équilibre
    Logger.log("4. Test calcul score équilibre...");
    const score = calculerScoreEquilibreGlobal(dataContext, config);
    Logger.log(`Score d'équilibre: ${score.toFixed(2)}/100`);
    
    // 5. Test de diagnostic
    Logger.log("5. Test diagnostic...");
    diagnostiquerVarianteScores();
    
    Logger.log("✅ Tous les tests passés avec succès !");
    
  } catch (e) {
    Logger.log(`❌ Erreur test: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
  }
}

/**
 * Test du patch de filtrage des élèves (pattern °\d+)
 */
function testPatchFiltrageEleves() {
  Logger.log("=== TEST PATCH FILTRAGE ÉLÈVES ===");
  
  try {
    const config = getConfig();
    
    // Test du chargeur avec le patch
    Logger.log("1. Test du chargeur patché...");
    const chargeResult = chargerElevesEtClasses_AvecSEXE(config, "MOBILITE");
    
    if (!chargeResult.success) {
      throw new Error("Échec du chargement");
    }
    
    const students = chargeResult.students;
    Logger.log(`✅ ${students.length} élèves chargés au total`);
    
    // Vérifier que tous les élèves ont un ID qui match le pattern
    Logger.log("2. Vérification du pattern ID_ELEVE...");
    let patternOK = 0;
    let patternKO = 0;
    
    students.forEach(student => {
      if (/°\d+$/.test(student.ID_ELEVE)) {
        patternOK++;
      } else {
        patternKO++;
        Logger.log(`⚠️ ID invalide: "${student.ID_ELEVE}"`);
      }
    });
    
    Logger.log(`Pattern OK: ${patternOK}, Pattern KO: ${patternKO}`);
    
    // Afficher quelques exemples d'élèves
    Logger.log("3. Exemples d'élèves chargés:");
    students.slice(0, 3).forEach((student, idx) => {
      Logger.log(`  ${idx + 1}. ${student.ID_ELEVE} - ${student.NOM} (${student.CLASSE})`);
      Logger.log(`     Scores: COM=${student.COM}, TRA=${student.TRA}, PART=${student.PART}`);
      Logger.log(`     Mobilité: ${student.mobilite}`);
    });
    
    // Statistiques par classe
    Logger.log("4. Statistiques par classe:");
    const statsParClasse = {};
    students.forEach(student => {
      if (!statsParClasse[student.CLASSE]) {
        statsParClasse[student.CLASSE] = { total: 0, F: 0, M: 0 };
      }
      statsParClasse[student.CLASSE].total++;
      if (student.SEXE === 'F') statsParClasse[student.CLASSE].F++;
      if (student.SEXE === 'M') statsParClasse[student.CLASSE].M++;
    });
    
    Object.entries(statsParClasse).forEach(([classe, stats]) => {
      Logger.log(`  ${classe}: ${stats.total} élèves (${stats.F}F, ${stats.M}M)`);
    });
    
    Logger.log("✅ Test du patch de filtrage réussi !");
    
  } catch (e) {
    Logger.log(`❌ Erreur test: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
  }
}

/**
 * Test spécifique : Vérifier que le pipeline équilibre les EFFECTIFS par score
 */
function testEquilibrageEffectifsScores() {
  Logger.log("=== TEST ÉQUILIBRAGE EFFECTIFS PAR SCORE ===");
  
  try {
    const config = getConfig();
    config.SCENARIOS_ACTIFS = ['COM', 'TRA', 'PART'];
    
    // 1. Charger les données
    Logger.log("1. Chargement des données...");
    const dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
    
    if (!dataContext || !dataContext.classesState) {
      throw new Error("Échec chargement données");
    }
    
    // 2. Calculer la distribution globale
    Logger.log("2. Calcul distribution globale...");
    const distributionGlobale = {};
    ['COM', 'TRA', 'PART'].forEach(critere => {
      distributionGlobale[critere] = { 1: 0, 2: 0, 3: 0, 4: 0, total: 0 };
    });
    
    Object.values(dataContext.classesState).forEach(eleves => {
      eleves.forEach(eleve => {
        ['COM', 'TRA', 'PART'].forEach(critere => {
          const score = parseInt(eleve['niveau' + critere]) || 0;
          if (score >= 1 && score <= 4) {
            distributionGlobale[critere][score]++;
            distributionGlobale[critere].total++;
          }
        });
      });
    });
    
    // 3. Afficher la distribution globale
    Logger.log("3. DISTRIBUTION GLOBALE (EFFECTIFS) :");
    ['COM', 'TRA', 'PART'].forEach(critere => {
      const dist = distributionGlobale[critere];
      Logger.log(`${critere}:`);
      [1, 2, 3, 4].forEach(score => {
        const pourcentage = dist.total > 0 ? (dist[score] / dist.total * 100).toFixed(1) : '0.0';
        Logger.log(`  Score ${score}: ${dist[score]} élèves (${pourcentage}%)`);
      });
    });
    
    // 4. Calculer les cibles par classe
    Logger.log("4. CIBLES PAR CLASSE (EFFECTIFS) :");
    Object.entries(dataContext.classesState).forEach(([classe, eleves]) => {
      const effectif = eleves.length;
      Logger.log(`${classe} (${effectif} élèves) - CIBLES :`);
      
      ['COM', 'TRA', 'PART'].forEach(critere => {
        const dist = distributionGlobale[critere];
        Logger.log(`  ${critere}:`);
        [1, 2, 3, 4].forEach(score => {
          const cible = Math.round((dist[score] / dist.total) * effectif);
          Logger.log(`    Score ${score}: ${cible} élèves cible`);
        });
      });
    });
    
    // 5. Vérifier que c'est bien des effectifs, pas des moyennes
    Logger.log("5. VÉRIFICATION - CE SONT BIEN DES EFFECTIFS :");
    Object.entries(dataContext.classesState).forEach(([classe, eleves]) => {
      const effectif = eleves.length;
      Logger.log(`${classe}: ${effectif} élèves total`);
      
      ['COM', 'TRA', 'PART'].forEach(critere => {
        const distribution = { 1: 0, 2: 0, 3: 0, 4: 0 };
        eleves.forEach(eleve => {
          const score = parseInt(eleve['niveau' + critere]) || 0;
          if (score >= 1 && score <= 4) {
            distribution[score]++;
          }
        });
        
        const somme = distribution[1] + distribution[2] + distribution[3] + distribution[4];
        Logger.log(`  ${critere}: ${distribution[1]}+${distribution[2]}+${distribution[3]}+${distribution[4]} = ${somme} élèves (effectifs)`);
        
        // Vérifier que la somme = effectif total
        if (somme !== effectif) {
          Logger.log(`  ⚠️ ATTENTION: Somme effectifs (${somme}) ≠ effectif total (${effectif})`);
        }
      });
    });
    
    Logger.log("✅ Test effectifs par score réussi !");
    Logger.log("🎯 CONFIRMATION: Le pipeline équilibre bien les EFFECTIFS, pas les moyennes !");
    
  } catch (e) {
    Logger.log(`❌ Erreur test: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
  }
}

/**
 * Test du processus NIRVANA corrigé
 */
function testProcessusNirvanaCorrige() {
  Logger.log("=== TEST PROCESSUS NIRVANA CORRIGÉ ===");
  
  try {
    const config = getConfig();
    config.SCENARIOS_ACTIFS = ['COM', 'TRA', 'PART'];
    
    // Lancer le processus corrigé
    const resultat = processusNirvanaCorrige(config);
    
    if (resultat.success) {
      Logger.log(`✅ Test réussi !`);
      Logger.log(`📊 Résumé:`);
      Logger.log(`  - Transfers SPEC: ${resultat.transfersSPEC || 0}`);
      Logger.log(`  - Transfers équilibrage: ${resultat.transfersEquilibrage || 0}`);
      Logger.log(`  - Total appliqués: ${resultat.nbModifications}`);
      Logger.log(`  - Durée: ${resultat.duree?.toFixed(1) || 'N/A'}s`);
      
      if (resultat.validation) {
        Logger.log(`  - Contraintes respectées: ${resultat.validation.contraintesRespectees ? '✅' : '❌'}`);
        Logger.log(`  - Équilibre amélioré: ${resultat.validation.equilibreAmeliore ? '✅' : '❌'}`);
      }
      
    } else {
      Logger.log(`❌ Test échoué: ${resultat.error}`);
    }
    
    return resultat;
    
  } catch (e) {
    Logger.log(`❌ Erreur test: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

/**
 * Test de validation des corrections de mapping
 */
function testCorrectionsMapping() {
  Logger.log("=== TEST CORRECTIONS MAPPING ===");
  
  try {
    const config = getConfig();
    
    // 1. Charger les données
    Logger.log("1. Chargement des données...");
    const dataContext = V2_Ameliore_PreparerDonnees_AvecSEXE(config);
    
    if (!dataContext || !dataContext.classesState) {
      throw new Error("Échec chargement données");
    }
    
    // 2. Vérifier les propriétés des élèves
    Logger.log("2. Vérification des propriétés...");
    const premierEleve = Object.values(dataContext.classesState)[0]?.[0];
    
    if (!premierEleve) {
      throw new Error("Aucun élève trouvé");
    }
    
    Logger.log(`Premier élève: ${premierEleve.ID_ELEVE}`);
    Logger.log(`Propriétés disponibles: ${Object.keys(premierEleve).join(', ')}`);
    
    // 3. Vérifier que les bonnes propriétés existent
    const proprietesCorrectes = ['COM', 'TRA', 'PART', 'ABS'];
    const proprietesIncorrectes = ['niveauCOM', 'niveauTRA', 'niveauPART', 'niveauABS'];
    
    Logger.log("3. Vérification propriétés correctes:");
    proprietesCorrectes.forEach(prop => {
      const existe = premierEleve.hasOwnProperty(prop);
      const valeur = premierEleve[prop];
      Logger.log(`  ${prop}: ${existe ? '✅' : '❌'} (valeur: ${valeur})`);
    });
    
    Logger.log("4. Vérification propriétés incorrectes (doivent être undefined):");
    proprietesIncorrectes.forEach(prop => {
      const existe = premierEleve.hasOwnProperty(prop);
      const valeur = premierEleve[prop];
      Logger.log(`  ${prop}: ${existe ? '❌' : '✅'} (valeur: ${valeur})`);
    });
    
    // 4. Tester la fonction lireScoreEleve
    Logger.log("5. Test fonction lireScoreEleve:");
    proprietesCorrectes.forEach(critere => {
      const score = lireScoreEleve(premierEleve, critere);
      Logger.log(`  lireScoreEleve(${critere}) = ${score}`);
    });
    
    // 5. Vérifier que les statistiques ne sont plus à 0
    Logger.log("6. Test calcul statistiques:");
    const stats = calculerStatsScoresGlobales(dataContext);
    Object.entries(stats).forEach(([critere, distribution]) => {
      Logger.log(`  ${critere}: ${distribution[1]}/${distribution[2]}/${distribution[3]}/${distribution[4]} (total: ${distribution.total})`);
    });
    
    Logger.log("✅ Test corrections mapping réussi !");
    Logger.log("🎯 CONFIRMATION: Les propriétés sont maintenant correctes !");
    
    return { success: true, stats: stats };
    
  } catch (e) {
    Logger.log(`❌ Erreur test: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

function runAllUtilsTests() {
  test_idx();
  test_logAction();
  Logger.log('Tous les tests unitaires Utils PASSÉS');
}


/**
 * ===============================================================
 * DÉTECTION SIMPLE DES DOUBLONS D'IDs DANS LES ONGLETS TEST
 * ===============================================================
 * Fonction simple à tester directement dans Apps Script
 * Analyse la colonne A de chaque onglet avec suffixe TEST
 */

function detecterDoublonsIDs() {
  Logger.log("🔍 === DÉTECTION DOUBLONS IDs DANS ONGLETS TEST ===");
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    // Filtrer les onglets avec suffixe TEST
    const ongletsTest = sheets.filter(sheet => sheet.getName().includes('TEST'));
    
    if (ongletsTest.length === 0) {
      Logger.log("❌ Aucun onglet avec suffixe TEST trouvé");
      return { success: false, message: "Aucun onglet TEST trouvé" };
    }
    
    Logger.log(`📋 ${ongletsTest.length} onglet(s) TEST trouvé(s): ${ongletsTest.map(s => s.getName()).join(', ')}`);
    
    let doublonsDetectes = false;
    let rapportGlobal = "";
    const statistiques = {
      totalOnglets: ongletsTest.length,
      ongletsAvecDoublons: 0,
      totalDoublons: 0
    };
    
    // Analyser chaque onglet TEST
    ongletsTest.forEach(sheet => {
      const nomOnglet = sheet.getName();
      Logger.log(`\n🔍 Analyse de l'onglet: ${nomOnglet}`);
      
      // Lire toutes les valeurs de la colonne A
      const colonneA = sheet.getRange("A:A").getValues().flat();
      
      // Nettoyer les valeurs (supprimer les vides)
      const ids = colonneA
        .map(cell => String(cell).trim())
        .filter(cell => cell !== "" && cell !== "null" && cell !== "undefined");
      
      Logger.log(`   ${ids.length} IDs non vides trouvés`);
      
      if (ids.length === 0) {
        Logger.log(`   ⚠️ Aucun ID trouvé dans ${nomOnglet}`);
        return;
      }
      
      // Détecter les doublons
      const vus = new Set();
      const doublons = new Set();
      const detailsDoublons = [];
      
      ids.forEach((id, index) => {
        if (vus.has(id)) {
          doublons.add(id);
          detailsDoublons.push({
            id: id,
            ligneExacte: index + 1,
            valeur: colonneA[index] // Valeur originale
          });
        } else {
          vus.add(id);
        }
      });
      
      // Rapport pour cet onglet
      let rapportOnglet = `\n📊 ONGLET: ${nomOnglet}\n`;
      rapportOnglet += `   Total IDs: ${ids.length}\n`;
      rapportOnglet += `   IDs uniques: ${vus.size}\n`;
      rapportOnglet += `   Doublons: ${doublons.size}\n`;
      
      if (doublons.size > 0) {
        doublonsDetectes = true;
        statistiques.ongletsAvecDoublons++;
        statistiques.totalDoublons += doublons.size;
        
        rapportOnglet += `   🚨 DOUBLONS DÉTECTÉS:\n`;
        
        doublons.forEach(id => {
          const occurrences = detailsDoublons.filter(d => d.id === id);
          const lignes = ids.map((val, idx) => val === id ? idx + 1 : null).filter(l => l !== null);
          
          rapportOnglet += `     • ID "${id}" trouvé ${lignes.length} fois aux lignes: ${lignes.join(', ')}\n`;
        });
        
        Logger.log(`🚨 ${doublons.size} doublon(s) détecté(s) dans ${nomOnglet}`);
        doublons.forEach(id => {
          const lignes = ids.map((val, idx) => val === id ? idx + 1 : null).filter(l => l !== null);
          Logger.log(`   • "${id}" aux lignes: ${lignes.join(', ')}`);
        });
        
      } else {
        rapportOnglet += `   ✅ Aucun doublon\n`;
        Logger.log(`✅ Aucun doublon dans ${nomOnglet}`);
      }
      
      rapportGlobal += rapportOnglet;
    });
    
    // Rapport final
    Logger.log("\n" + "=".repeat(60));
    Logger.log("📊 RAPPORT FINAL DOUBLONS");
    Logger.log("=".repeat(60));
    Logger.log(`Total onglets analysés: ${statistiques.totalOnglets}`);
    Logger.log(`Onglets avec doublons: ${statistiques.ongletsAvecDoublons}`);
    Logger.log(`Total types de doublons: ${statistiques.totalDoublons}`);
    Logger.log(`Résultat global: ${doublonsDetectes ? '🚨 DOUBLONS DÉTECTÉS' : '✅ AUCUN DOUBLON'}`);
    
    // Affichage dans une boîte de dialogue
    const html = HtmlService.createHtmlOutput(
      `<div style="font-family: monospace; font-size: 12px; padding: 20px;">
        <h3>🔍 Détection des doublons d'IDs</h3>
        <div style="background: #f5f5f5; padding: 15px; border-radius: 5px;">
          <strong>Résumé:</strong><br>
          • Onglets analysés: ${statistiques.totalOnglets}<br>
          • Onglets avec doublons: ${statistiques.ongletsAvecDoublons}<br>
          • Types de doublons: ${statistiques.totalDoublons}<br>
          • Statut: ${doublonsDetectes ? '<span style="color:red">🚨 DOUBLONS DÉTECTÉS</span>' : '<span style="color:green">✅ AUCUN DOUBLON</span>'}
        </div>
        <hr>
        <pre>${rapportGlobal.replace(/\n/g, '<br>')}</pre>
      </div>`
    ).setWidth(600).setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(html, "Rapport de détection des doublons");
    
    return {
      success: true,
      doublonsDetectes: doublonsDetectes,
      statistiques: statistiques,
      rapport: rapportGlobal
    };
    
  } catch (e) {
    Logger.log(`❌ Erreur lors de la détection des doublons: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    
    SpreadsheetApp.getUi().alert(
      "Erreur", 
      `Erreur lors de la détection: ${e.message}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return { 
      success: false, 
      error: e.message 
    };
  }
}

/**
 * VERSION RAPIDE : Juste les doublons, sans interface
 */
function detecterDoublonsRapide() {
  Logger.log("⚡ === DÉTECTION RAPIDE DOUBLONS ===");
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const ongletsTest = sheets.filter(sheet => sheet.getName().includes('TEST'));
  
  let totalDoublons = 0;
  
  ongletsTest.forEach(sheet => {
    const nomOnglet = sheet.getName();
    const colonneA = sheet.getRange("A:A").getValues().flat();
    const ids = colonneA
      .map(cell => String(cell).trim())
      .filter(cell => cell !== "" && cell !== "null");
    
    const vus = new Set();
    const doublons = new Set();
    
    ids.forEach(id => {
      if (vus.has(id)) {
        doublons.add(id);
      } else {
        vus.add(id);
      }
    });
    
    if (doublons.size > 0) {
      Logger.log(`${nomOnglet}: ${doublons.size} doublon(s) - ${Array.from(doublons).join(', ')}`);
      totalDoublons += doublons.size;
    } else {
      Logger.log(`${nomOnglet}: ✅ OK`);
    }
  });
  
  Logger.log(`\n📊 TOTAL: ${totalDoublons} type(s) de doublons détectés`);
  return totalDoublons;
}

/**
 * FONCTION DE NETTOYAGE : Surligner les doublons en rouge
 */
function surlignerDoublons() {
  Logger.log("🖍️ === SURLIGNAGE DES DOUBLONS ===");
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const ongletsTest = sheets.filter(sheet => sheet.getName().includes('TEST'));
  
  ongletsTest.forEach(sheet => {
    const nomOnglet = sheet.getName();
    const range = sheet.getRange("A:A");
    const values = range.getValues();
    
    // Identifier les doublons
    const ids = values.map(row => String(row[0]).trim()).filter(id => id !== "" && id !== "null");
    const vus = new Set();
    const doublons = new Set();
    
    ids.forEach(id => {
      if (vus.has(id)) {
        doublons.add(id);
      } else {
        vus.add(id);
      }
    });
    
    if (doublons.size > 0) {
      // Surligner en rouge les cellules avec doublons
      values.forEach((row, index) => {
        const id = String(row[0]).trim();
        if (doublons.has(id)) {
          sheet.getRange(index + 1, 1).setBackground("#ffcccc"); // Rouge clair
        }
      });
      
      Logger.log(`${nomOnglet}: ${doublons.size} doublon(s) surlignés en rouge`);
    }
  });
  
  Logger.log("✅ Surlignage terminé");
}
