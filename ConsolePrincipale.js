/**
 * Ouvre la console principale de répartition
 * VERSION UNIQUE - Assurez-vous qu'aucune autre définition n'existe dans le projet.
 */
function ouvrirConsole() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Console.html')
      .setWidth(400)
      //.setHeight(600) // Hauteur souvent mieux gérée par le contenu ou omise
      .setTitle('Console de Répartition');
    SpreadsheetApp.getUi().showSidebar(html);
    Logger.log("Sidebar 'Console.html' ouverte.");
  } catch (e) {
      Logger.log(`Erreur ouverture Console.html: ${e.message} ${e.stack}`);
      SpreadsheetApp.getUi().alert(`Impossible d'ouvrir la console principale: ${e.message}`);
  }
}

/* La deuxième définition de ouvrirConsole a été supprimée d'ici */

/**
 * Fonction d'interface pour le bouton OPTIONS (Console.html)
 * Fait appel à repartirPhase1() et formate le résultat pour l'affichage.
 */
function repartirOptions() {
  try {
    Logger.log("Appel de repartirPhase1...");
    const resultat = repartirPhase1(); // Dépendance Externe

    if (resultat && resultat.success) {
      // Formatage message succès (vérifier si 'details' existe)
      let message = `Phase 1 (Options) exécutée avec succès.\n`;
      if (resultat.details && Array.isArray(resultat.details)) {
          message += `${resultat.details.length} règles appliquées.\n`;
          const totalPlaces = resultat.details.reduce((sum, d) => sum + (d.places || 0), 0);
          message += `Total: ${totalPlaces} élèves placés.\n\n`;

          // Optionnel: Formatage détaillé (peut être trop long pour une alerte)
          /*
          const classeGroups = {};
          resultat.details.forEach(d => {
            if (!classeGroups[d.classe]) classeGroups[d.classe] = [];
            classeGroups[d.classe].push(d);
          });
          for (const classe in classeGroups) {
            message += `${classe}:\n`;
            classeGroups[classe].forEach(d => {
              message += `- ${d.option || '?'} (${d.type || '?'}): ${d.places || 0}/${d.quota || '?'}\n`;
            });
            message += '\n';
          }
          */
      } else {
          message += "Aucun détail fourni sur les règles appliquées.";
      }
      return message;
    } else {
      // Gestion erreur ou échec
      const errorMessage = resultat?.message || "Erreur inconnue ou résultat invalide de Phase 1.";
      Logger.log(`Échec ou erreur Phase 1: ${errorMessage}`);
      return `Erreur lors de l'exécution de la Phase 1 (Options): ${errorMessage}`;
    }
  } catch (e) {
    Logger.log(`Erreur technique dans repartirOptions: ${e.message} ${e.stack}`);
    return "Erreur technique Phase 1: " + e.message;
  }
}

/**
 * Fonction d'interface pour le bouton CODES (Console.html)
 * Fait appel à repartirPhase2() et formate le résultat pour l'affichage.
 */
function repartirCodes() {
  try {
    Logger.log("Appel de repartirPhase2...");
    const resultat = repartirPhase2(); // Dépendance Externe

    if (resultat && resultat.success) {
      return `Phase 2 (Codes) exécutée avec succès:\n- ${resultat.asso || 0} élèves associés traités\n- ${resultat.disso || 0} élèves dissociés traités\n- ${resultat.ecrits || 0} élèves placés/modifiés par codes.`;
    } else {
      const errorMessage = resultat?.message || "Erreur inconnue ou résultat invalide de Phase 2.";
       Logger.log(`Échec ou erreur Phase 2: ${errorMessage}`);
      return `Erreur lors de l'exécution de la Phase 2 (Codes): ${errorMessage}`;
    }
  } catch (e) {
    Logger.log(`Erreur technique dans repartirCodes: ${e.message} ${e.stack}`);
    return "Erreur technique Phase 2: " + e.message;
  }
}


/**
 * Fonction d'interface pour le bouton PARITÉ (Console.html)
 * Fait appel à executerParite() et formate le résultat pour l'affichage.
 */
function repartirParite() {
  try {
    Logger.log("Appel de executerParite...");
    const resultat = executerParite(); // Dépendance Externe

    if (resultat && resultat.success) {
      return `Phase 3 (Parité/Capacité) exécutée avec succès.\n${resultat.nbAjoutes || 0} élèves placés/ajustés.\n\n${resultat.message || ""}`;
    } else {
      const errorMessage = resultat?.message || "Erreur inconnue ou résultat invalide de Phase 3.";
      Logger.log(`Échec ou erreur Phase 3: ${errorMessage}`);
      // Gérer aussi le cas où resultat est null/undefined mais pas une erreur technique
      if (resultat === null || resultat === undefined) {
           return "Phase 3 (Parité) terminée sans retour de statut explicite.";
      }
      return `Erreur lors de l'exécution de la Phase 3 (Parité): ${errorMessage}`;
    }
  } catch (e) {
    Logger.log(`Erreur technique dans repartirParite: ${e.message} ${e.stack}`);
    return "Erreur technique Phase 3: " + e.message;
  }
}

/**
 * Fonction d'interface pour le bouton "Répartir Tout" (Console.html)
 * Exécute successivement les trois phases principales de répartition.
 */
function executerPhases1a3() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let messageTotal = "";
    let phaseOk = true;
    let toastMessage = "";

    // Phase 1
    toastMessage = "Exécution Phase 1 (Options)...";
    ss.toast(toastMessage, "Répartition Complète", 5);
    Logger.log(toastMessage);
    const resultatPhase1 = repartirPhase1(); // Dépendance Externe
    if (resultatPhase1 && resultatPhase1.success) {
        const places = (resultatPhase1.details && Array.isArray(resultatPhase1.details))
                       ? resultatPhase1.details.reduce((sum, d) => sum + (d.places || 0), 0) : 0;
        const regles = (resultatPhase1.details && Array.isArray(resultatPhase1.details))
                       ? resultatPhase1.details.length : 0;
        messageTotal += `✓ Phase 1: ${regles} règles, ${places} élèves placés.\n`;
    } else {
        messageTotal += `❌ Échec Phase 1: ${resultatPhase1?.message || "Erreur inconnue"}\n`;
        phaseOk = false;
    }

    // Phase 2 (si phase 1 ok)
    if(phaseOk) {
        toastMessage = "Exécution Phase 2 (Codes)...";
        ss.toast(toastMessage, "Répartition Complète", 5);
        Logger.log(toastMessage);
        const resultatPhase2 = repartirPhase2(); // Dépendance Externe
        if (resultatPhase2 && resultatPhase2.success) {
            messageTotal += `✓ Phase 2: ${resultatPhase2.asso || 0} asso, ${resultatPhase2.disso || 0} disso, ${resultatPhase2.ecrits || 0} placés.\n`;
        } else {
            messageTotal += `❌ Échec Phase 2: ${resultatPhase2?.message || "Erreur inconnue"}\n`;
            phaseOk = false;
        }
    }

    // Phase 3 (si phases précédentes ok)
    if(phaseOk) {
        toastMessage = "Exécution Phase 3 (Parité/Capacité)...";
        ss.toast(toastMessage, "Répartition Complète", 5);
        Logger.log(toastMessage);
        const resultatPhase3 = executerParite(); // Dépendance Externe
        if (resultatPhase3 && resultatPhase3.success) {
            messageTotal += `✓ Phase 3: ${resultatPhase3.nbAjoutes || 0} élèves placés.\n`;
        } else {
            messageTotal += `❌ Échec Phase 3: ${resultatPhase3?.message || "Erreur inconnue"}\n`;
            phaseOk = false;
        }
    }

    const finalMessage = phaseOk ? `Répartition complète terminée avec succès!\n\n${messageTotal}` : `Répartition incomplète ou échouée.\n\n${messageTotal}`;
    ss.toast(phaseOk ? "Répartition terminée !" : "Répartition échouée/incomplète.", "Répartition Complète - Fin", 10);
    Logger.log(finalMessage);
    return finalMessage; // Retourne le message pour l'alert() dans Console.html

  } catch (e) {
    Logger.log(`Erreur majeure dans executerPhases1a3: ${e.message} ${e.stack}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Erreur Technique: ${e.message}`, "Répartition Complète - ERREUR", 10);
    return "Erreur technique lors de la répartition complète: " + e.message;
  }
}

/**
 * Ouvre l'interface de l'assistant d'optimisation (Phase 4)
 * Cette fonction est appelée par le bouton "LANCER ASSISTANT" dans la section OPTIMISATION
 * de la Console.html
 */
function lancerAssistantOptimisationUI() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Phase4UI.html')
      .setWidth(450)
      .setTitle("Assistant d'Optimisation - Phase 4");
    
    SpreadsheetApp.getUi().showSidebar(html);
    Logger.log("Sidebar 'Phase4UI.html' ouverte.");
  } catch (e) {
    Logger.log(`Erreur ouverture Phase4UI.html: ${e.message} ${e.stack}`);
    SpreadsheetApp.getUi().alert(`Impossible d'ouvrir l'assistant d'optimisation: ${e.message}`);
  }
}
/** Affiche la sidebar d'optimisation V11 */
function showOptimisationSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Phase4UI.html') // Nom du fichier HTML V11
      .setTitle("Optimisation V11")
      .setWidth(450); // Largeur UI V11
    SpreadsheetApp.getUi().showSidebar(html);
  } catch(e) {
     Logger.log(`Erreur ouverture Sidebar Phase 4 V11: ${e.message} ${e.stack}`);
     SpreadsheetApp.getUi().alert(`Erreur: Impossible d'ouvrir l'interface d'optimisation.\nFichier HTML 'Phase4UI.html' manquant ou invalide ?\n(${e.message})`);
  }
}

/** Affiche la sidebar de finalisation Phase 5 */
function showFinalisationSidebar() {
   try {
    const html = HtmlService.createHtmlOutputFromFile('FinalisationUI.html') // Nom du fichier HTML Phase 5
      .setTitle("Finalisation Phase 5")
      .setWidth(550); // Largeur UI Phase 5
    SpreadsheetApp.getUi().showSidebar(html);
  } catch(e) {
     Logger.log(`Erreur ouverture Sidebar Phase 5: ${e.message} ${e.stack}`);
     SpreadsheetApp.getUi().alert(`Erreur: Impossible d'ouvrir l'interface de finalisation.\nFichier HTML 'FinalisationUI.html' manquant ou invalide ?\n(${e.message})`);
  }
}
// Code côté Google Apps Script (à ajouter à votre fichier .gs)

/************************************************************
 *  ConsolePrincipale.gs
 *  — ANCIEN NOM :  afficherStatistiquesSources()
 *  — NOUVEAU NOM : ouvrirStatistiquesSources()
 ************************************************************/
function ouvrirStatistiquesSources() {
  const html = HtmlService.createHtmlOutputFromFile('StatistiquesDashboard')
                .setWidth(800)
                .setHeight(600)
                .setTitle('Statistiques Classes Sources');

  const params = {
    viewType: 'sources',
    title   : 'Statistiques des Classes Sources'
  };

  const data = obtenirDonneesDetailleesPlageEleves('sources');

  PropertiesService.getScriptProperties().setProperty(
    'TEMP_STATS_DATA',
    JSON.stringify({ params, data })
  );

  SpreadsheetApp.getUi().showModalDialog(html,
    'Statistiques des Classes Sources');
}

/**
 * Affiche les statistiques des classes TEST
 */
function afficherStatistiquesTest() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('StatistiquesDashboard')
      .setWidth(900)  // Harmoniser la taille
      .setHeight(600)
      .setTitle('Statistiques Classes TEST');
    
    const params = {
      viewType: 'test',
      title: 'Statistiques des Classes TEST'
    };
    
    const data = obtenirDonneesDetailleesPlageEleves('test');
    
    PropertiesService.getScriptProperties().setProperty('TEMP_STATS_DATA', JSON.stringify({
      params: params,
      data: data
    }));
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Statistiques des Classes TEST');
    return { success: true };  // Ajouter le retour
  } catch (e) {
    Logger.log("Erreur afficherStatistiquesTest: " + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Affiche les statistiques des classes DEF
 */
function afficherStatistiquesDef() {
  const html = HtmlService.createHtmlOutputFromFile('StatistiquesDashboard')
    .setWidth(800)
    .setHeight(600)
    .setTitle('Statistiques Classes DEF');
  
  const params = {
    viewType: 'def',
    title: 'Statistiques des Classes DEF'
  };
  
  // Utiliser obtenirDonneesDetailleesPlageEleves au lieu de obtenirDonneesStatistiques
  const data = obtenirDonneesDetailleesPlageEleves('def');
  
  PropertiesService.getScriptProperties().setProperty('TEMP_STATS_DATA', JSON.stringify({
    params: params,
    data: data
  }));
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Statistiques des Classes DEF');
}
// ---- Vues internes du dashboard ---------------------------------
/*************************************************************
 *  Setters ultra-légers pour la modale Statistiques
 *  → renvoient IMMÉDIATEMENT le JSON attendu par handleDataLoad
 *************************************************************/
function setViewDef() {
  return {
    params: {viewType:'def', title:'Statistiques Classes DEF'},
    data  : obtenirDonneesDetailleesPlageEleves('def')
  };
}

function setViewTest() {
  return {
    params: {viewType:'test', title:'Statistiques Classes TEST'},
    data  : obtenirDonneesDetailleesPlageEleves('test')
  };
}

function setViewSources() {
  return {
    params: {viewType:'sources', title:'Statistiques Classes SOURCES'},
    data  : obtenirDonneesDetailleesPlageEleves('sources')
  };
}

function setViewCompare() {
  return {
    params  : {viewType:'compare', title:'Comparaison TEST / DEF'},
    dataTest: obtenirDonneesDetailleesPlageEleves('test'),
    dataDef : obtenirDonneesDetailleesPlageEleves('def')
  };
}

// ---- utilitaire commun ------------------------------------------
function setTempStatsAndReturn(params, payload){
  PropertiesService.getScriptProperties().setProperty(
    'TEMP_STATS_DATA',
    JSON.stringify({params, ...payload})
  );
  return {success:true};      // rien d’autre !
}

/**
 * Compare les statistiques entre classes TEST et DEF
 */
function comparerStatistiques() {
  const html = HtmlService.createHtmlOutputFromFile('StatistiquesDashboard')
    .setWidth(800)
    .setHeight(600)
    .setTitle('Comparaison TEST vs DEF');
  
  const params = {
    viewType: 'compare',
    title: 'Comparaison TEST vs DEF'
  };
  
  // Utiliser obtenirDonneesDetailleesPlageEleves au lieu de obtenirDonneesStatistiques
  const dataTest = obtenirDonneesDetailleesPlageEleves('test');
  const dataDef = obtenirDonneesDetailleesPlageEleves('def');
  
  PropertiesService.getScriptProperties().setProperty('TEMP_STATS_DATA', JSON.stringify({
    params: params,
    dataTest: dataTest,
    dataDef: dataDef
  }));
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Comparaison TEST vs DEF');
}

/**
 * Fonction pour récupérer les données stockées temporairement
 * Cette fonction sera appelée par la page HTML
 */
function getStatsDatatemp() {
  const jsonData = PropertiesService.getScriptProperties().getProperty('TEMP_STATS_DATA');
  if (!jsonData) {
    return { error: "Aucune donnée disponible" };
  }
  return JSON.parse(jsonData);
}

/**
 * Fonction qui récupère les données statistiques depuis les feuilles
 * @param {string} type - 'sources', 'test', ou 'def'
 */
function obtenirDonneesStatistiques(type) {
  // Cette fonction doit être implémentée selon votre structure de données
  // Voici un exemple de structure
  
  // Simuler des données pour les besoins de l'exemple
  // Dans une implémentation réelle, vous devriez extraire ces données de vos feuilles
  
  let prefix;
  switch(type) {
    case 'sources': prefix = 'SRC'; break;
    case 'test': prefix = 'TST'; break;
    case 'def': prefix = 'DEF'; break;
    default: prefix = 'SRC';
  }
  
  // Structure d'exemple (à remplacer par vos données réelles)
  return [
    { class: `${prefix}-A`, scores: [12, 20, 8, 3], female: 22, male: 21 },
    { class: `${prefix}-B`, scores: [9, 14, 12, 5], female: 18, male: 22 },
    { class: `${prefix}-C`, scores: [6, 18, 10, 2], female: 15, male: 21 }
  ];
}
/**
 * Fonction qui récupère les données complètes pour toutes les vues de statistiques
 * Cette fonction sera appelée par la page HTML
 */
function getCompleteStatsData() {
  try {
    // Utilisation de obtenirDonneesDetailleesPlageEleves au lieu de obtenirDonneesDetaillees
    const sourcesData = obtenirDonneesDetailleesPlageEleves('sources');
    const testData = obtenirDonneesDetailleesPlageEleves('test');
    const defData = obtenirDonneesDetailleesPlageEleves('def');
    
    return {
      success: true,
      data: {
        sources: sourcesData,
        test: testData,
        def: defData
      }
    };
  } catch(e) {
    return {
      success: false,
      error: e.message || "Erreur lors de la récupération des données"
    };
  }
}

/**
 * Fonction alternative de récupération des données
 * Qui examine explicitement la plage des élèves (sans les stats)
 */
  function obtenirDonneesDetailleesPlageEleves(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = [];
  
  // Filtrer les onglets selon le type demandé
  let prefixFilter;
switch(type) {
  case 'sources': 
    prefixFilter = (name) => isSourceClassName(name); 
    break;
  case 'test': 
    prefixFilter = (name) => {
      name = String(name).toUpperCase();
      return name.includes('TEST') || name.includes('TST');
    }; 
    break;
  case 'def': 
    prefixFilter = (name) => {
      name = String(name).toUpperCase();
      return name.includes('DEF');
    }; 
    break;
  default: 
    prefixFilter = () => true;
}
  
  // Parcourir tous les onglets
  const sheets = ss.getSheets();
  for (let sheet of sheets) {
    const sheetName = sheet.getName();
    
    // Filtrer selon le type
    if (!prefixFilter(sheetName)) continue;
    
    try {
      // Récupérer toutes les données
      const range = sheet.getDataRange();
      const values = range.getValues();
      
      // Vérifier s'il y a assez de lignes
      if (values.length < 2) continue;
      
      // Trouver les colonnes importantes
      // Trouver les colonnes importantes
const headers = values[0].map(h => String(h || '').toUpperCase());

// Détection améliorée de la colonne sexe
let sexeIndex = headers.findIndex(h =>
  h === 'F' || h === 'M' || h === 'SEXE' || h.includes('GENRE') ||
  h.includes('F/M') || h.includes('M/F') || h.includes('F/G'));

// Si non trouvé, parcourir les premières lignes pour détecter F/M
if (sexeIndex === -1) {
  for (let i = 0; i < headers.length; i++) {
    let foundF = false, foundM = false;
    for (let row = 1; row < Math.min(6, values.length); row++) {
      const cell = String(values[row][i] || '').trim().toUpperCase();
      if (cell === 'F') foundF = true;
      if (cell === 'M' || cell === 'H' || cell === 'G') foundM = true;
    }
    if (foundF && foundM) {
      sexeIndex = i;
      break;
    }
  }
}

// Détection améliorée des autres colonnes
const comIndex = headers.findIndex(h => h === 'COM' || h === 'COMPORTEMENT' || h === 'COMP');
const traIndex = headers.findIndex(h => h === 'TRA' || h === 'TRAVAIL' || h === 'TRAV');
const partIndex = headers.findIndex(h => h === 'PART' || h === 'PARTICIPATION');
const absIndex = headers.findIndex(h => h === 'ABS' || h === 'ABSENCE' || h === 'ABSENTÉISME');
      
      // Si on ne trouve pas les colonnes essentielles, passer à la feuille suivante
      if (comIndex === -1 && traIndex === -1 && partIndex === -1 && absIndex === -1) continue;
      
      // Initialiser les compteurs
      const scores = {
        com: [0, 0, 0, 0],
        tra: [0, 0, 0, 0],
        part: [0, 0, 0, 0],
        abs: [0, 0, 0, 0]
      };
      
      let female = 0;
      let male = 0;
      
      // Trouver où finissent les données des élèves
      let lastStudentRow = values.length;
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        // Chercher une ligne vide ou une ligne avec des formules de statistiques
        if (row.every(cell => cell === '') || 
            row.some(cell => String(cell).includes('MOYENNE') || 
                           String(cell).includes('TOTAL') || 
                           String(cell).includes('STAT'))) {
          lastStudentRow = i;
          break;
        }
      }
      
      // Compter les élèves et leurs scores
      for (let i = 1; i < lastStudentRow; i++) {
        const row = values[i];
        
        // S'assurer que c'est une ligne d'élève (présence d'au moins un score)
        const hasScores = (comIndex !== -1 && row[comIndex]) || 
                         (traIndex !== -1 && row[traIndex]) || 
                         (partIndex !== -1 && row[partIndex]) || 
                         (absIndex !== -1 && row[absIndex]);
                         
        if (!hasScores) continue;
        
        // Compter par sexe
        if (sexeIndex !== -1) {
          const sexe = String(row[sexeIndex]).trim().toUpperCase();
          if (sexe === 'F') female++;
          else if (sexe === 'M' || sexe === 'H' || sexe === 'G') male++;
        }
        
        // Compter les scores
        if (comIndex !== -1 && row[comIndex]) {
          const score = parseInt(row[comIndex]);
          if (score >= 1 && score <= 4) scores.com[score-1]++;
        }
        
        if (traIndex !== -1 && row[traIndex]) {
          const score = parseInt(row[traIndex]);
          if (score >= 1 && score <= 4) scores.tra[score-1]++;
        }
        
        if (partIndex !== -1 && row[partIndex]) {
          const score = parseInt(row[partIndex]);
          if (score >= 1 && score <= 4) scores.part[score-1]++;
        }
        
        if (absIndex !== -1 && row[absIndex]) {
          const score = parseInt(row[absIndex]);
          if (score >= 1 && score <= 4) scores.abs[score-1]++;
        }
      }
      
      // N'ajouter que les feuilles avec des élèves
      const totalEleves = scores.com.reduce((a, b) => a + b, 0) || 
                         scores.tra.reduce((a, b) => a + b, 0) || 
                         scores.part.reduce((a, b) => a + b, 0) || 
                         scores.abs.reduce((a, b) => a + b, 0);
      
      if (totalEleves > 0) {
        result.push({
          class: sheetName,
          scores: scores,
          female: female,
          male: male
        });
      }
    } catch (e) {
      Logger.log(`Erreur traitement ${sheetName}: ${e.message}`);
    }
  }
  
  return result;
}
/**
 * Récupère la liste de tous les onglets du classeur avec leur état de visibilité
 * @return {Object} Un objet avec un statut de succès et la liste des onglets
 */
function getListeOnglets() {
  try {
    // Récupérer le classeur actif et forcer la mise à jour du cache
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.flush(); // Force l'application des modifications en attente
    
    // Récupérer tous les onglets, y compris les nouveaux
    const sheets = ss.getSheets();
    
    // Créer un tableau pour stocker les informations des onglets
    const onglets = [];
    
    // Log pour débogage
    Logger.log(`Récupération de ${sheets.length} onglets au total`);
    
    // Parcourir tous les onglets et collecter leurs informations
    sheets.forEach(function(sheet) {
      const id = sheet.getSheetId();
      const nom = sheet.getName();
      const masque = sheet.isSheetHidden();
      
      // Log détaillé pour débogage
      Logger.log(`Onglet trouvé: "${nom}" (ID: ${id}) - Masqué: ${masque}`);
      
      onglets.push({
        id: id.toString(),
        nom: nom,
        masque: masque
      });
    });
    
    // Tri alphabétique pour une meilleure organisation
    onglets.sort((a, b) => a.nom.localeCompare(b.nom, undefined, {numeric: true, sensitivity: 'base'}));
    
    // Retourner les données avec un flag de succès
    return {
      success: true,
      onglets: onglets,
      total: onglets.length
    };
  } catch (error) {
    // Log d'erreur détaillé
    Logger.log(`Erreur dans getListeOnglets: ${error.message} | Stack: ${error.stack}`);
    
    // En cas d'erreur, retourner un objet avec les détails de l'erreur
    return {
      success: false,
      message: "Erreur lors de la récupération des onglets: " + error.message
    };
  }
}

/**
 * Modifie la visibilité des onglets selon les paramètres fournis
 * @param {Array} visibilites - Tableau d'objets avec les IDs des onglets et leur état de visibilité
 * @return {Object} Un objet avec un statut de succès et le nombre d'onglets modifiés
 */
function modifierVisibiliteOnglets(visibilites) {
  try {
    // Récupérer le classeur actif
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let modifies = 0;
    
    // Créer un mapping des feuilles par ID pour un accès plus rapide
    const sheets = ss.getSheets();
    const sheetsMap = {};
    sheets.forEach(function(sheet) {
      sheetsMap[sheet.getSheetId().toString()] = sheet;
    });
    
    // Parcourir les changements de visibilité demandés
    visibilites.forEach(function(item) {
      try {
        const sheet = sheetsMap[item.id];
        
        if (sheet) {
          const estActuellementMasque = sheet.isSheetHidden();
          const doitEtreMasque = !item.visible;
          
          // Ne faire le changement que si nécessaire
          // (l'état actuel est différent de l'état souhaité)
          if (estActuellementMasque !== doitEtreMasque) {
            if (doitEtreMasque) {
              sheet.hideSheet();
            } else {
              sheet.showSheet();
            }
            modifies++;
          }
        }
      } catch (e) {
        Logger.log("Erreur avec l'onglet " + item.id + ": " + e.message);
      }
    });
    
    return {
      success: true,
      modifies: modifies,
      message: `${modifies} onglet(s) modifié(s) avec succès.`
    };
  } catch (error) {
    return {
      success: false,
      message: "Erreur lors de la modification des onglets: " + error.message
    };
  }
}
/**
 * Vérifie le mot de passe administrateur en utilisant le mot de passe stocké dans _CONFIG B3
 * @param {string} motDePasse - Le mot de passe à vérifier
 * @return {Object} Résultat de la vérification
 */
function verifierMotDePasseAdmin(motDePasse) {
  try {
    // Récupérer le mot de passe stocké dans _CONFIG B3
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("_CONFIG");
    
    if (!configSheet) {
      return { success: false, message: "Feuille _CONFIG introuvable" };
    }
    
    const motDePasseStocke = configSheet.getRange("B3").getValue();
    
    // Vérification du mot de passe
    return {
      success: motDePasse === motDePasseStocke
    };
  } catch (error) {
    Logger.log("Erreur dans verifierMotDePasseAdmin: " + error.message);
    return { 
      success: false,
      message: "Erreur lors de la vérification: " + error.message
    };
  }
}
/**
 * Vérifie le mot de passe pour l'accès aux onglets protégés
 * Utilise le même mot de passe que la fonction verifierMotDePasseAdmin
 * @param {string} motDePasse - Le mot de passe à vérifier
 * @return {Object} Résultat de la vérification
 */
function verifierMotDePasseOnglets(motDePasse) {
  try {
    // Récupérer le mot de passe stocké dans _CONFIG B3
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("_CONFIG");
    
    if (!configSheet) {
      return { success: false, message: "Feuille _CONFIG introuvable" };
    }
    
    const motDePasseStocke = configSheet.getRange("B3").getValue();
    Logger.log("Vérification mot de passe onglets: " + motDePasse + " vs stocké: " + motDePasseStocke);
    
    // Vérification du mot de passe
    return {
      success: motDePasse === motDePasseStocke
    };
  } catch (error) {
    Logger.log("Erreur dans verifierMotDePasseOnglets: " + error.message);
    return { 
      success: false,
      message: "Erreur lors de la vérification: " + error.message
    };
  }
}
/**
 * Force le rechargement complet de la liste des onglets en vidant le cache
 * @return {Object} La nouvelle liste des onglets
 */
function rechargerOnglets() {
  try {
    // Force le vidage du cache Google Apps Script
    CacheService.getScriptCache().remove("LISTE_ONGLETS_CACHE");
    
    // Actualise toutes les modifications en cours
    SpreadsheetApp.flush();
    
    // Attendre brièvement pour s'assurer que l'API a bien pris en compte les modifications
    Utilities.sleep(500);
    
    // Récupérer la liste fraîche des onglets
    const result = getListeOnglets();
    
    return {
      ...result,
      message: `Liste rechargée avec succès (${result.onglets?.length || 0} onglets)`
    };
  } catch (error) {
    Logger.log(`Erreur dans rechargerOnglets: ${error.message}`);
    return {
      success: false,
      message: "Erreur lors du rechargement des onglets: " + error.message
    };
  }
}
/**
 * Fonction d'interface pour le bouton CODES RÉSERVATION de la console
 * Formate le résultat pour un affichage optimal dans l'interface
 */
function analyserCodesReservationConsole() {
  try {
    Logger.log("=== ANALYSE CODES RÉSERVATION - CONSOLE ===");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Lire les classes sources
    const classeSources = lireClassesSources();
    
    if (classeSources.length === 0) {
      return "❌ Aucune classe source trouvée.\n\nVérifiez l'onglet _STRUCTURE :\n• Colonne A doit contenir 'SOURCE'\n• Colonne B doit contenir les noms de classes (ex: 4°1, 5°3)";
    }
    
    // 2. Analyser les codes
    const statsGlobales = analyserCodesDansClasses(classeSources);
    
    // 3. Formater pour la console (version compacte)
    let message = `📊 CODES RÉSERVATION - ${classeSources.length} classe${classeSources.length > 1 ? 's' : ''} analysée${classeSources.length > 1 ? 's' : ''}\n`;
    message += `════════════════════════════════════\n\n`;
    
    // Codes D compacts
    const codesD = Object.keys(statsGlobales.codesD).sort();
    if (codesD.length > 0) {
      message += `🔴 CODES D: ${codesD.map(code => `${code}=${statsGlobales.codesD[code]}`).join(' • ')}\n`;
    } else {
      message += `🔴 CODES D: Aucun\n`;
    }
    
    // Codes A compacts
    const codesA = Object.keys(statsGlobales.codesA).sort();
    if (codesA.length > 0) {
      message += `🟢 CODES A: ${codesA.map(code => `${code}=${statsGlobales.codesA[code]}`).join(' • ')}\n`;
    } else {
      message += `🟢 CODES A: Aucun\n`;
    }
    
    // Totaux
    const totalD = codesD.reduce((sum, code) => sum + statsGlobales.codesD[code], 0);
    const totalA = codesA.reduce((sum, code) => sum + statsGlobales.codesA[code], 0);
    message += `\n📈 TOTAUX: ${totalD} codes D • ${totalA} codes A\n`;
    
    // Classes analysées
    message += `\n📋 Classes: ${classeSources.join(', ')}`;
    
    // Toast notification pour feedback rapide
    ss.toast(`Codes analysés: ${totalD} codes D, ${totalA} codes A`, "Codes Réservation", 5);
    
    Logger.log("Analyse codes réservation terminée avec succès");
    return message;
    
  } catch (e) {
    Logger.log(`❌ ERREUR analyserCodesReservationConsole: ${e.message}`);
    return `❌ Erreur lors de l'analyse des codes :\n\n${e.message}`;
  }
}

/**
 * Fonction de diagnostic pour vérifier la structure
 * Peut être appelée manuellement depuis l'éditeur de script
 */
function diagnostiquerStructureCodesReservation() {
  Logger.log("=== DIAGNOSTIC STRUCTURE CODES RÉSERVATION ===");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Vérifier _STRUCTURE
  const structureSheet = ss.getSheetByName("_STRUCTURE");
  if (!structureSheet) {
    Logger.log("❌ Onglet _STRUCTURE introuvable");
    return;
  }
  
  Logger.log("✅ Onglet _STRUCTURE trouvé");
  
  const data = structureSheet.getDataRange().getValues();
  Logger.log(`📊 Dimensions _STRUCTURE: ${data.length} lignes × ${data[0].length} colonnes`);
  
  // Analyser les en-têtes
  if (data.length > 0) {
    Logger.log(`📋 En-têtes: ${data[0].join(' | ')}`);
  }
  
  // Chercher les sources
  let sourcesFound = 0;
  for (let i = 1; i < data.length; i++) {
    const type = String(data[i][0] || '').trim().toUpperCase();
    const nomClasse = String(data[i][1] || '').trim();
    
    if (type === 'SOURCE') {
      sourcesFound++;
      const formatValide = /^[3-6]°[1-9]\d*$/.test(nomClasse);
      Logger.log(`${formatValide ? '✅' : '⚠️'} Source ${sourcesFound}: "${nomClasse}" ${formatValide ? '(format valide)' : '(format invalide)'}`);
      
      // Vérifier si l'onglet existe
      const sheet = ss.getSheetByName(nomClasse);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        Logger.log(`   📊 Onglet "${nomClasse}": ${lastRow} lignes × ${lastCol} colonnes`);
        
        if (lastCol >= 14) {
          Logger.log(`   ✅ Colonnes M(13) et N(14) accessibles`);
        } else {
          Logger.log(`   ❌ Pas assez de colonnes (besoin de 14, trouvé ${lastCol})`);
        }
      } else {
        Logger.log(`   ❌ Onglet "${nomClasse}" introuvable`);
      }
    }
  }
  
  Logger.log(`📈 RÉSUMÉ: ${sourcesFound} source${sourcesFound > 1 ? 's' : ''} trouvée${sourcesFound > 1 ? 's' : ''} dans _STRUCTURE`);
  Logger.log("=== FIN DIAGNOSTIC ===");
}