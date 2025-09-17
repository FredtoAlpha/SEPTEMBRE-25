// Harmonisation : tous les utilitaires (logAction, idx, etc.) sont centralisés dans Utils.js
// Suppression de toute définition locale, tous les appels utilisent la version centrale
// Exemple : Utils.logAction("...")

/**
 * Génère les valeurs NOM_PRENOM et ID_ELEVE pour toutes les feuilles pertinentes
 * Puis masque les colonnes A, B et C pour n'afficher que la colonne D (NOM_PRENOM)
 */
function genererNomPrenomEtID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Confirmation de l'utilisateur
  const reponse = ui.alert(
    'Générer NOM_PRENOM et ID_ELEVE',
    'Cette action va générer les valeurs NOM_PRENOM et ID_ELEVE, puis masquer les colonnes A, B et C pour ne laisser visible que la colonne NOM_PRENOM.\n\nContinuer ?',
    ui.ButtonSet.YES_NO
  );
  
  if (reponse !== ui.Button.YES) return;
  
  Logger.log("=== Début génération NOM_PRENOM et ID_ELEVE ===");
  
  // Récupérer les feuilles sources
  let sheets;
  try {
    sheets = getSourceSheets();
    Logger.log(`${sheets.length} feuilles sources trouvées`);
  } catch (e) {
    Logger.log(`ERREUR récupération feuilles sources: ${e}`);
    sheets = [];
  }
  
  if (!sheets || sheets.length === 0) {
    ui.alert("Aucune feuille source trouvée", "Impossible de trouver des feuilles à traiter.", ui.ButtonSet.OK);
    return;
  }
  
  let totalNomPrenom = 0;
  let totalIDs = 0;
  
  // Pour chaque feuille, traiter les données
  sheets.forEach(sheet => {
    try {
      const sheetName = sheet.getName();
      Logger.log(`Traitement de l'onglet: ${sheetName}`);
      
      // Trouver les colonnes
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const colIdEleve = headers.indexOf('ID_ELEVE') + 1;
      const colNom = headers.indexOf('NOM') + 1;
      const colPrenom = headers.indexOf('PRENOM') + 1;
      const colNomPrenom = headers.indexOf('NOM_PRENOM') + 1;
      
      // Vérifier que toutes les colonnes nécessaires existent
      if (colIdEleve <= 0 || colNom <= 0 || colPrenom <= 0 || colNomPrenom <= 0) {
        Logger.log(`ERREUR: Une ou plusieurs colonnes requises manquantes sur ${sheetName}`);
        return;
      }
      
      // Lire les données existantes
      const lastRow = Math.max(sheet.getLastRow(), 2);
      const data = sheet.getRange(2, 1, lastRow - 1, Math.max(colNom, colPrenom, colNomPrenom, colIdEleve)).getValues();
      
      // Préparer le préfixe ID basé sur le nom de l'onglet
      const prefixeID = sheetName.replace(/\s+/g, ''); 
      
      // Trouver le plus grand ID existant pour cette classe
      let maxNumero = 0;
      const idsExistants = new Set(); // Pour vérifier les collisions d'IDs
      
      data.forEach(row => {
        const idActuel = String(row[colIdEleve - 1] || '');
        if (idActuel.startsWith(prefixeID)) {
          // Ajouter à la liste des IDs existants
          idsExistants.add(idActuel);
          
          // Extraire le numéro après le préfixe
          const match = idActuel.match(new RegExp(`^${prefixeID}(\\d+)$`));
          if (match && match[1]) {
            const numero = parseInt(match[1], 10);
            if (!isNaN(numero) && numero > maxNumero) {
              maxNumero = numero;
            }
          }
        }
      });
      
      // Déterminer le prochain numéro d'ID
      let prochainNumero = maxNumero > 0 ? maxNumero + 1 : 1001;
      
      Logger.log(`  Préfixe ID: ${prefixeID}, Plus grand ID existant: ${maxNumero}, Prochain ID: ${prochainNumero}`);
      
      // Parcourir les données et effectuer les mises à jour
      let compteurMAJ = 0;
      let compteurID = 0;
      
      data.forEach((row, idx) => {
        const rowIndex = idx + 2; // +2 car idx commence à 0 et on saute l'en-tête
        const nom = String(row[colNom - 1] || '').trim();
        const prenom = String(row[colPrenom - 1] || '').trim();
        const idActuel = String(row[colIdEleve - 1] || '').trim();
        
        // Ne traiter que les lignes avec nom et prénom
        if (nom && prenom) {
          // Mettre à jour NOM_PRENOM
          const nomPrenom = `${nom} ${prenom}`;
          sheet.getRange(rowIndex, colNomPrenom).setValue(nomPrenom);
          compteurMAJ++;
          
          // Générer un ID si non existant ou incorrect
          if (!idActuel || !idActuel.startsWith(prefixeID)) {
            // Générer un ID unique non utilisé
            let nouvelID;
            do {
              nouvelID = `${prefixeID}${prochainNumero}`;
              prochainNumero++;
            } while (idsExistants.has(nouvelID));
            
            // Appliquer le nouvel ID et l'ajouter à la liste des utilisés
            sheet.getRange(rowIndex, colIdEleve).setValue(nouvelID);
            idsExistants.add(nouvelID);
            compteurID++;
          }
        } else if (row[colNomPrenom - 1]) {
          // Si NOM_PRENOM a une valeur mais NOM ou PRENOM est vide, effacer NOM_PRENOM
          sheet.getRange(rowIndex, colNomPrenom).setValue("");
        }
      });
      
      // NOUVELLE PARTIE : Masquer les colonnes A, B et C
      try {
        // Vérifier que colIdEleve, colNom et colPrenom correspondent bien à A, B et C
        const colonnesAMasquer = [];
        if (colIdEleve === 1) colonnesAMasquer.push(1); // A
        if (colNom === 2) colonnesAMasquer.push(2); // B
        if (colPrenom === 3) colonnesAMasquer.push(3); // C
        
        // Si les colonnes sont bien A, B et C, les masquer
        if (colonnesAMasquer.length > 0) {
          sheet.hideColumns(1, 3); // Masquer les colonnes A, B et C
          Logger.log(`  Colonnes A, B et C masquées dans ${sheetName}`);
        } else {
          // Si les colonnes ne sont pas A, B et C, on masque les colonnes spécifiques
          if (colIdEleve > 0) sheet.hideColumns(colIdEleve, 1);
          if (colNom > 0) sheet.hideColumns(colNom, 1);
          if (colPrenom > 0) sheet.hideColumns(colPrenom, 1);
          Logger.log(`  Colonnes ID_ELEVE, NOM et PRENOM masquées dans ${sheetName}`);
        }
      } catch (e) {
        Logger.log(`  ERREUR lors du masquage des colonnes sur ${sheetName}: ${e}`);
      }
      
      Logger.log(`${sheetName}: ${compteurMAJ} NOM_PRENOM générés, ${compteurID} ID_ELEVE créés`);
      totalNomPrenom += compteurMAJ;
      totalIDs += compteurID;
      
    } catch (e) {
      Logger.log(`ERREUR sur ${sheet?.getName() || 'feuille inconnue'}: ${e}`);
    }
  });
  
  // Traiter la feuille de CONSOLIDATION si elle existe
  const consolidationSheet = ss.getSheetByName('CONSOLIDATION');
  if (consolidationSheet) {
    try {
      Logger.log("Mise à jour de l'onglet CONSOLIDATION...");
      // Pour CONSOLIDATION, on ne met à jour que NOM_PRENOM, pas les ID
      const headers = consolidationSheet.getRange(1, 1, 1, consolidationSheet.getLastColumn()).getValues()[0];
      const colNom = headers.indexOf('NOM') + 1;
      const colPrenom = headers.indexOf('PRENOM') + 1;
      const colNomPrenom = headers.indexOf('NOM_PRENOM') + 1;
      const colIdEleve = headers.indexOf('ID_ELEVE') + 1;
      
      if (colNom > 0 && colPrenom > 0 && colNomPrenom > 0) {
        const lastRow = Math.max(consolidationSheet.getLastRow(), 2);
        const data = consolidationSheet.getRange(2, 1, lastRow - 1, Math.max(colNom, colPrenom, colNomPrenom)).getValues();
        
        let compteurMAJ = 0;
        data.forEach((row, idx) => {
          const rowIndex = idx + 2;
          const nom = String(row[colNom - 1] || '').trim();
          const prenom = String(row[colPrenom - 1] || '').trim();
          
          if (nom && prenom) {
            const nomPrenom = `${nom} ${prenom}`;
            consolidationSheet.getRange(rowIndex, colNomPrenom).setValue(nomPrenom);
            compteurMAJ++;
          } else if (row[colNomPrenom - 1]) {
            consolidationSheet.getRange(rowIndex, colNomPrenom).setValue("");
          }
        });
        
        // Masquer aussi les colonnes dans CONSOLIDATION
        try {
          // Masquer les colonnes spécifiques
          if (colIdEleve > 0) consolidationSheet.hideColumns(colIdEleve, 1);
          if (colNom > 0) consolidationSheet.hideColumns(colNom, 1);
          if (colPrenom > 0) consolidationSheet.hideColumns(colPrenom, 1);
          Logger.log("  Colonnes ID_ELEVE, NOM et PRENOM masquées dans CONSOLIDATION");
        } catch (e) {
          Logger.log(`  ERREUR lors du masquage des colonnes sur CONSOLIDATION: ${e}`);
        }
        
        Logger.log(`CONSOLIDATION: ${compteurMAJ} NOM_PRENOM générés`);
        totalNomPrenom += compteurMAJ;
      }
    } catch (e) {
      Logger.log(`ERREUR sur CONSOLIDATION: ${e}`);
    }
  }
  
  Logger.log(`=== Fin génération NOM_PRENOM et ID_ELEVE (${totalNomPrenom} NOM_PRENOM, ${totalIDs} IDs) ===`);
  
  // Afficher un message de confirmation
  ui.alert('Opération terminée', 
           `Résultats:\n- ${totalIDs} identifiants uniques générés\n- ${totalNomPrenom} noms+prénoms combinés\n- Colonnes ID_ELEVE, NOM et PRENOM masquées\n\n` +
           'Consultez les journaux pour plus de détails.', 
           ui.ButtonSet.OK);
  
  // Enregistrer l'opération dans le journal
  try { 
    Utils.logAction(`Génération NOM_PRENOM (${totalNomPrenom}) et ID_ELEVE (${totalIDs}) + masquage colonnes`); 
  } catch (e) { 
    Logger.log("Note: fonction logAction non disponible"); 
  }
}
