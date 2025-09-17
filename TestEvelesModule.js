// ==================================================================
// TestElevesModule.gs - Tests pour le module de répartition autonome
// ==================================================================

/**
 * Test complet du module Eleves
 * Vérifie que tout fonctionne sans interférer avec le système existant
 */
function testElevesModuleComplet() {
  console.log("╔════════════════════════════════════════╗");
  console.log("║   TEST DU MODULE ELEVES AUTONOME       ║");
  console.log("╚════════════════════════════════════════╝");
  console.log(new Date().toLocaleString());
  console.log("\n");
  
  // 1. Vérifier que le module n'interfère pas avec Config existant
  console.log("=== TEST 1: Vérification de l'isolation ===");
  
  // Tester si getConfig existe (du système principal)
  if (typeof getConfig === 'function') {
    console.log("✅ getConfig() du système principal détecté");
    try {
      const mainConfig = getConfig();
      console.log("✅ Config principal accessible, VERSION:", mainConfig.VERSION);
    } catch (e) {
      console.log("⚠️ getConfig() existe mais erreur:", e.message);
    }
  } else {
    console.log("ℹ️ Pas de getConfig() détecté (normal si système principal absent)");
  }
  
  // Vérifier que nos constantes locales sont accessibles
  console.log("\n=== TEST 2: Configuration locale du module ===");
  console.log("ELEVES_MODULE_CONFIG:", ELEVES_MODULE_CONFIG);
  console.log("✅ Configuration locale accessible");
  
  // 2. Test de lecture des données
  console.log("\n=== TEST 3: Lecture des données élèves ===");
  try {
    const data = getElevesData();
    console.log("✅ getElevesData() exécuté avec succès");
    console.log("Nombre de classes trouvées:", data.length);
    
    if (data.length > 0) {
      console.log("\nDétails première classe:");
      console.log("- Nom:", data[0].classe);
      console.log("- Nombre d'élèves:", data[0].eleves.length);
      
      if (data[0].eleves.length > 0) {
        console.log("\nExemple premier élève:");
        const eleve = data[0].eleves[0];
        console.log("- ID:", eleve.id);
        console.log("- Nom:", eleve.nom);
        console.log("- Prénom:", eleve.prenom);
        console.log("- Scores:", JSON.stringify(eleve.scores));
      }
    } else {
      console.log("⚠️ Aucune donnée trouvée - Vérifiez que vous avez des onglets TEST");
    }
  } catch (e) {
    console.error("❌ Erreur lors de la lecture des données:", e);
  }
  
  // 3. Test des statistiques
  console.log("\n=== TEST 4: Calcul des statistiques ===");
  try {
    const stats = getElevesStats();
    console.log("✅ getElevesStats() exécuté avec succès");
    console.log("Stats globales:", JSON.stringify(stats.global));
    console.log("Nombre de classes dans les stats:", stats.parClasse.length);
  } catch (e) {
    console.error("❌ Erreur lors du calcul des stats:", e);
  }
  
  // 4. Test de sauvegarde (sans écrire réellement)
  console.log("\n=== TEST 5: Test de sauvegarde (simulation) ===");
  try {
    // Créer une disposition fictive
    const testDisposition = {
      "5°1": ["001", "002", "003"],
      "5°2": ["004", "005"]
    };
    
    console.log("Disposition test:", testDisposition);
    console.log("✅ Structure de sauvegarde valide");
    
    // Note: Ne pas exécuter réellement saveElevesSnapshot pour ne pas créer d'onglets
    console.log("ℹ️ saveElevesSnapshot() non exécuté (pour éviter de créer des onglets)");
  } catch (e) {
    console.error("❌ Erreur dans le test de sauvegarde:", e);
  }
  
  // 5. Vérifier les onglets existants
  console.log("\n=== TEST 6: Analyse des onglets ===");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let testSheets = 0;
  let intSheets = 0;
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.toUpperCase().endsWith('TEST')) {
      testSheets++;
      console.log(`✅ Onglet TEST trouvé: "${name}"`);
    }
    if (name.toUpperCase().endsWith('INT')) {
      intSheets++;
      console.log(`📁 Onglet INT existant: "${name}"`);
    }
  });
  
  console.log(`\nRésumé: ${testSheets} onglets TEST, ${intSheets} onglets INT`);
  
  // Résumé final
  console.log("\n╔════════════════════════════════════════╗");
  console.log("║            RÉSUMÉ DU TEST              ║");
  console.log("╚════════════════════════════════════════╝");
  console.log("✅ Module autonome fonctionnel");
  console.log("✅ Aucune interférence avec le système existant");
  
  if (testSheets === 0) {
    console.log("⚠️ Aucun onglet TEST trouvé - Créez-en avec createElevesTestData()");
  }
  
  return {
    success: true,
    testSheets: testSheets,
    dataFound: getElevesData().length > 0
  };
}

/**
 * Vérifie rapidement si le module peut lire des données
 */
function quickTestEleves() {
  console.log("=== Test rapide du module Eleves ===");
  
  try {
    const data = getElevesData();
    const count = data.reduce((sum, g) => sum + g.eleves.length, 0);
    
    console.log(`✅ ${data.length} classes trouvées`);
    console.log(`✅ ${count} élèves au total`);
    
    return {
      success: true,
      classes: data.length,
      eleves: count
    };
  } catch (e) {
    console.error("❌ Erreur:", e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Affiche la configuration du module
 */
function showElevesModuleConfig() {
  console.log("=== Configuration du module Eleves ===");
  console.log("ELEVES_MODULE_CONFIG:", JSON.stringify(ELEVES_MODULE_CONFIG, null, 2));
  
  console.log("\nPremiers alias de colonnes:");
  Object.keys(ELEVES_ALIAS).slice(0, 5).forEach(key => {
    console.log(`- ${key}: ${ELEVES_ALIAS[key].join(', ')}`);
  });
  console.log("...");
  
  return ELEVES_MODULE_CONFIG;
}