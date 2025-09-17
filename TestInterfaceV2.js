// ==================================================================
// TestInterfaceV2.gs - Tests rapides pour la nouvelle interface
// ==================================================================

/**
 * Test rapide de l'interface V2
 * Vérifie que tout est prêt pour le déploiement
 */
function testInterfaceV2() {
  console.log("=== TEST INTERFACE V2 ===");
  console.log(new Date().toLocaleString());
  
  // 1. Vérifier les données
  console.log("\n1. Test getElevesData:");
  try {
    const data = getElevesData();
    console.log(`✅ ${data.length} classes trouvées`);
    
    let totalEleves = 0;
    data.forEach(g => {
      console.log(`   - ${g.classe}: ${g.eleves.length} élèves`);
      totalEleves += g.eleves.length;
    });
    console.log(`✅ Total: ${totalEleves} élèves`);
  } catch (e) {
    console.error("❌ Erreur getElevesData:", e);
  }
  
  // 2. Vérifier les règles
  console.log("\n2. Test getStructureRules:");
  try {
    const rules = getStructureRules();
    const nbRules = Object.keys(rules).length;
    
    if (nbRules > 0) {
      console.log(`✅ ${nbRules} classes avec règles`);
      Object.entries(rules).forEach(([classe, rule]) => {
        console.log(`   - ${classe}: capacité=${rule.capacity}, quotas=${Object.keys(rule.quotas).length}`);
      });
    } else {
      console.log("⚠️ Aucune règle trouvée - Créez l'onglet _STRUCTURE");
    }
  } catch (e) {
    console.error("❌ Erreur getStructureRules:", e);
  }
  
  // 3. Vérifier les colonnes importantes
  console.log("\n3. Vérification des colonnes:");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheets = ss.getSheets().filter(s => s.getName().toUpperCase().endsWith('TEST'));
  
  if (testSheets.length > 0) {
    const firstSheet = testSheets[0];
    const headers = firstSheet.getRange(1, 1, 1, firstSheet.getLastColumn()).getValues()[0];
    
    console.log(`Analyse de "${firstSheet.getName()}":`);
    
    // Vérifier les colonnes critiques
    const criticalCols = ['ID', 'NOM', 'SEXE', 'LV2', 'OPT', 'COM', 'TRA', 'PART', 'ABS'];
    criticalCols.forEach(col => {
      const found = headers.some(h => String(h).toUpperCase().includes(col));
      console.log(`   ${found ? '✅' : '❌'} Colonne ${col}`);
    });
  }
  
  console.log("\n=== RÉSUMÉ ===");
  console.log("Si tous les tests sont ✅, l'interface est prête !");
  console.log("Sinon, corrigez les ❌ avant de déployer.");
}

/**
 * Créer des données de test avec scores variés
 */
function createTestDataWithScores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "5°TEST_DEMO";
  
  // Supprimer si existe
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  
  // Créer nouvel onglet
  const sheet = ss.insertSheet(sheetName);
  
  const headers = [
    "ID_ELEVE", "NOM", "PRENOM", "SEXE", "LV2", "OPT", 
    "COM", "TRA", "PART", "ABS", "DISSO", "ASSO", 
    "SOURCE", "DISPO", "MOBILITE"
  ];
  
  // Données variées pour tester l'affichage
  const testData = [
    headers,
    // Élèves avec différents scores
    ["001", "MARTIN", "Sophie", "F", "ESP", "CHAV", 4, 4, 4, 1, "1", "", "6°A", "", "LIBRE"],
    ["002", "DURAND", "Lucas", "M", "ITA", "", 3, 3, 2, 2, "", "3", "6°B", "PAI", "LIBRE"],
    ["003", "BERNARD", "Emma", "F", "ALL", "EURO", 4, 3, 3, 1, "2", "", "6°A", "", "LIBRE"],
    ["004", "PETIT", "Noah", "M", "ESP", "BIL", 2, 2, 4, 3, "", "1", "6°C", "", "LIBRE"],
    ["005", "ROBERT", "Léa", "F", "ITA", "GREC", 1, 3, 3, 4, "", "", "6°B", "PAP", "LIBRE"],
    ["006", "MOREAU", "Jules", "M", "LATIN", "CHAV", 4, 4, 4, 0, "1", "2", "6°D", "", "LIBRE"],
    ["007", "ROUSSEAU", "Alice", "F", "ESP", "", 3, 2, 1, 2, "", "", "6°A", "", "LIBRE"],
    ["008", "LAMBERT", "Tom", "M", "ITA", "EURO", 4, 3, 4, 1, "3", "", "6°C", "GEVASCO", "LIBRE"]
  ];
  
  // Écrire les données
  sheet.getRange(1, 1, testData.length, headers.length).setValues(testData);
  
  // Formater
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#5b21b6')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  sheet.autoResizeColumns(1, headers.length);
  
  console.log(`✅ Onglet "${sheetName}" créé avec ${testData.length - 1} élèves`);
  console.log("Scores variés pour tester l'affichage des pastilles colorées");
  
  return sheetName;
}

/**
 * Affiche un aperçu des scores pour vérifier les couleurs
 */
function previewScoreColors() {
  console.log("\n=== APERÇU DES COULEURS DE SCORES ===");
  console.log("Score 4 : 🟢 Vert foncé (#006400)");
  console.log("Score 3 : 🟢 Vert clair (#22c55e)");
  console.log("Score 2 : 🟡 Jaune (#fbbf24)");
  console.log("Score 1 : 🔴 Rouge (#dc2626)");
  console.log("\nChaque score apparaît comme une pastille ronde avec la lettre (C/T/P/A)");
}