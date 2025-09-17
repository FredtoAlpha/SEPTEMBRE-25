/**
 * Fonctions de diagnostic pour comprendre pourquoi 0 swaps
 */

// 1. Vérifier la configuration
function diagnostiquerConfiguration() {
  Logger.log("=== DIAGNOSTIC CONFIGURATION ===");
  const config = getConfig();
  
  Logger.log("NIVEAU: " + config.NIVEAU);
  Logger.log("OPTIONS pour ce niveau: " + JSON.stringify(config.OPTIONS));
  Logger.log("LV2_OPTIONS: " + JSON.stringify(config.LV2_OPTIONS));
  Logger.log("OPTIONS_TO_TRACK_IN_BILAN: " + JSON.stringify(config.OPTIONS_TO_TRACK_IN_BILAN));
  
  // Vérifier _CONFIG
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("_CONFIG");
  if (configSheet) {
    const data = configSheet.getDataRange().getValues();
    Logger.log("\n=== Contenu _CONFIG ===");
    data.forEach((row, i) => {
      if (row[0] && (row[0] === "OPT" || row[0] === "LV2")) {
        Logger.log(`Ligne ${i+1}: ${row[0]} = ${row[1]}`);
      }
    });
  }
}

// 2. Vérifier les pools d'options
function diagnostiquerPools() {
  Logger.log("\n=== DIAGNOSTIC POOLS D'OPTIONS ===");
  const config = getConfig();
  const niveau = determinerNiveauActifCache();
  const structureResult = chargerStructureEtOptions(niveau, config);
  
  if (structureResult.success) {
    const optionPools = buildOptionPools(structureResult.structure, config);
    Logger.log("Pools construits: " + JSON.stringify(optionPools, null, 2));
  } else {
    Logger.log("ERREUR chargement structure: " + structureResult.message);
  }
}

// 3. Vérifier la mobilité des élèves
function diagnostiquerMobilite() {
  Logger.log("\n=== DIAGNOSTIC MOBILITÉ ===");
  const config = getConfig();
  const result = chargerElevesEtClasses(config, "MOBILITE");
  
  if (result.success) {
    const stats = {
      FIXE: 0,
      SPEC: 0,
      PERMUT: 0,
      CONDI: 0,
      LIBRE: 0,
      total: 0
    };
    
    result.students.forEach(eleve => {
      const mobilite = eleve.mobilite || 'LIBRE';
      stats[mobilite] = (stats[mobilite] || 0) + 1;
      stats.total++;
    });
    
    Logger.log("Répartition mobilité:");
    Object.keys(stats).forEach(key => {
      const pct = stats.total > 0 ? (stats[key] / stats.total * 100).toFixed(1) : 0;
      Logger.log(`  ${key}: ${stats[key]} (${pct}%)`);
    });
    
    // Afficher quelques exemples
    Logger.log("\nExemples d'élèves:");
    result.students.slice(0, 5).forEach(e => {
      Logger.log(`  ID: ${e.ID_ELEVE}, Mobilité: ${e.mobilite}, Option: ${e.optionKey}, Classe: ${e.CLASSE}`);
    });
  }
}

// 4. Tester les contraintes entre deux élèves
function testerContraintesSwap() {
  Logger.log("\n=== TEST CONTRAINTES SWAP ===");
  const config = getConfig();
  const niveau = determinerNiveauActifCache();
  
  // Charger tout
  const structureResult = chargerStructureEtOptions(niveau, config);
  const elevesResult = chargerElevesEtClasses(config, "MOBILITE");
  
  if (!structureResult.success || !elevesResult.success) {
    Logger.log("ERREUR chargement données");
    return;
  }
  
  const optionPools = buildOptionPools(structureResult.structure, config);
  const dissocMap = buildDissocCountMap(elevesResult.classesMap);
  
  // Trouver deux élèves de classes différentes
  const classes = Object.keys(elevesResult.classesMap);
  if (classes.length < 2) {
    Logger.log("Moins de 2 classes!");
    return;
  }
  
  const eleve1 = elevesResult.classesMap[classes[0]][0];
  const eleve2 = elevesResult.classesMap[classes[1]][0];
  
  if (eleve1 && eleve2) {
    Logger.log(`\nTest swap entre:`);
    Logger.log(`  E1: ${eleve1.ID_ELEVE} (${eleve1.CLASSE}, Mob: ${eleve1.mobilite}, Opt: ${eleve1.optionKey})`);
    Logger.log(`  E2: ${eleve2.ID_ELEVE} (${eleve2.CLASSE}, Mob: ${eleve2.mobilite}, Opt: ${eleve2.optionKey})`);
    
    const ok = respecteContraintes(
      eleve1, eleve2, 
      elevesResult.students, 
      structureResult.structure, 
      structureResult.optionsNiveau, 
      optionPools, 
      dissocMap
    );
    
    Logger.log(`\nRésultat: ${ok ? "SWAP POSSIBLE" : "SWAP IMPOSSIBLE"}`);
  }
}

// 5. Diagnostic complet
function diagnosticComplet() {
  Logger.log("==========================================================");
  Logger.log("DIAGNOSTIC COMPLET V14 - POURQUOI 0 SWAPS ?");
  Logger.log("==========================================================");
  
  diagnostiquerConfiguration();
  diagnostiquerPools();
  diagnostiquerMobilite();
  testerContraintesSwap();
  
  Logger.log("\n=== RECOMMANDATIONS ===");
  Logger.log("1. Vérifiez que tous les élèves ne sont pas FIXE/SPEC");
  Logger.log("2. Vérifiez que les options dans _STRUCTURE correspondent aux onglets TEST");
  Logger.log("3. Vérifiez que OPT et LV2 sont bien définis dans _CONFIG");
  Logger.log("4. Vérifiez que la colonne MOBILITE existe dans les onglets TEST");
}