/**
 * Présentation.gs
 * Gère la création et la mise en forme de l'onglet d'accueil "ACCUEIL".
 */

/**
 * Crée un onglet de présentation professionnel.
 * Fonction principale divisée en sous-fonctions.
 * NE DOIT PAS bloquer l'exécution globale de l'initialisation.
 */
function creerOngletPresentation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi(); // Pour l'alerte de confirmation

  // Vérifier si l'onglet existe déjà
  let presentationSheet = ss.getSheetByName("ACCUEIL");

  if (presentationSheet) {
    // Confirmation non bloquante si exécuté depuis un menu par exemple
    // Si appelé par initialiserSysteme, la suppression a déjà eu lieu potentiellement.
    // On garde la confirmation au cas où la fonction est appelée isolément.
    const response = ui.alert(
      "Onglet ACCUEIL existant",
      "L'onglet ACCUEIL existe déjà. Voulez-vous le recréer ?\n(Attention: Annuler ici peut laisser le classeur dans un état incohérent si lancé depuis l'initialisation complète)",
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      Logger.log("Suppression de l'ancien onglet ACCUEIL.");
      try {
          // S'assurer qu'il reste au moins une autre feuille avant de supprimer
          if (ss.getSheets().length > 1) {
            ss.deleteSheet(presentationSheet);
            presentationSheet = null; // Réinitialiser la variable
          } else {
              Logger.log("Impossible de supprimer ACCUEIL, c'est la dernière feuille.");
              ui.alert("Impossible de recréer ACCUEIL car c'est la seule feuille restante.");
              return; // Arrêter la fonction ici
          }
      } catch (e) {
          Logger.log(`Erreur lors de la suppression de ACCUEIL: ${e}`);
          ui.alert(`Erreur lors de la suppression de l'ancien onglet ACCUEIL: ${e.message}`);
          return; // Arrêter si la suppression échoue
      }
    } else {
       Logger.log("Annulation de la recréation de l'onglet ACCUEIL.");
      return; // L'utilisateur a annulé
    }
  }

  // Créer un nouvel onglet et le placer en première position
  Logger.log("Création du nouvel onglet ACCUEIL.");
  presentationSheet = ss.insertSheet("ACCUEIL", 0); // Position 0 = première feuille

  // Définir les couleurs utilisées dans tout le script
  const couleurs = {
    principale: "#1a73e8", // Bleu Google
    secondaire: "#e8f0fe", // Bleu très clair
    accent: "#fbbc04", // Jaune Google
    bordure: "#dadce0", // Gris bordure Google
    texteBlanc: "#ffffff",
    texteNoir: "#000000"
  };

  // Obtenir le niveau depuis _CONFIG (fonction doit exister !)
  let niveau = obtenirNiveauActuel(); // Appel sans argument 'ss' si elle utilise getConfig()

  // Appeler les sous-fonctions pour construire l'onglet
  // Encapsuler dans un try...catch global pour cette fonction
  try {
      // Étape 1: Préparer la feuille (mise en forme de base, colonnes)
      preparerFeuillePresentation(presentationSheet, couleurs);

      // Étape 2: Créer la bannière avec titre
      creerBanniereTitre(presentationSheet, niveau, couleurs);

      // Étape 3: Créer la section menu
      creerSectionMenu(presentationSheet, couleurs);

      // Étape 4: Créer la section À propos
      creerSectionAPropos(presentationSheet, couleurs);

      // Étape 5: Créer la section Démarrage rapide
      const rowAfterDemarrage = creerSectionDemarrageRapide(presentationSheet, couleurs);

      // Étape 6: Créer la section Documentation
      const rowAfterDocumentation = creerSectionDocumentation(presentationSheet, rowAfterDemarrage, couleurs);

      // Étape 7: Créer le bouton d'action et pied de page
      creerBoutonEtPiedDePage(presentationSheet, rowAfterDocumentation, couleurs);

      // Étape 8: Tentative finale de mise en forme du menu (ne bloque pas)
      tentativeFinaleFormatMenu(presentationSheet, couleurs);

      Logger.log("Onglet ACCUEIL créé et formaté avec succès.");

      // --- SUPPRESSION DE L'ACTIVATION ET DE L'ALERTE BLOQUANTE ---
      // L'activation de l'onglet et le message final seront gérés par le processus d'initialisation principal.
      // ss.setActiveSheet(presentationSheet); // <- Supprimé/Commenté
      // ui.alert(...) // <- Supprimé/Commenté

  } catch (e_main) {
      Logger.log(`ERREUR MAJEURE lors de la création de l'onglet ACCUEIL: ${e_main.toString()} - Stack: ${e_main.stack}`);
      ui.alert(`Une erreur s'est produite lors de la création de l'onglet ACCUEIL: ${e_main.message}`);
      // Essayer de supprimer l'onglet potentiellement corrompu ?
      try {
          if (presentationSheet && ss.getSheetByName(presentationSheet.getName())) {
              if (ss.getSheets().length > 1) ss.deleteSheet(presentationSheet);
          }
      } catch (e_del) { /* Ignorer erreur de suppression */ }
  }
}

// =========================================================================
//                SOUS-FONCTIONS DE creerOngletPresentation
// =========================================================================

/**
 * Obtient le niveau actuel configuré en utilisant getConfig().
 * @return {string} Le niveau configuré ou une valeur par défaut.
 */
function obtenirNiveauActuel() {
  const config = getConfig();
  return config.NIVEAU;
}

/**
 * Prépare la feuille avec les paramètres de base.
 */
function preparerFeuillePresentation(sheet, couleurs) {
  Logger.log(" - Préparation de la feuille ACCUEIL...");
  sheet.setHiddenGridlines(true);
  // Appliquer fond blanc seulement à la zone utilisée probable (optimisation)
  // Au lieu de 1000 lignes / 26 colonnes, limiter si possible
  sheet.getRange("A1:G50").setBackground(couleurs.texteBlanc || "white"); // Zone estimée
  sheet.setTabColor(couleurs.principale); // Couleur de l'onglet

  // Ajuster la largeur des colonnes
  sheet.setColumnWidth(1, 20);    // Marge gauche (A)
  sheet.setColumnWidth(2, 150);   // Contenu 1 (B)
  sheet.setColumnWidth(3, 150);   // Contenu 2 (C)
  sheet.setColumnWidth(4, 150);   // Contenu 3 (D)
  sheet.setColumnWidth(5, 150);   // Contenu 4 (E)
  sheet.setColumnWidth(6, 150);   // Contenu 5 (F)
  sheet.setColumnWidth(7, 20);    // Marge droite (G)
  // Masquer les colonnes inutilisées à droite (à partir de H)
  if (sheet.getMaxColumns() > 7) {
      sheet.deleteColumns(8, sheet.getMaxColumns() - 7);
  }
   // Masquer les lignes inutilisées en bas (à partir de 51)
   if (sheet.getMaxRows() > 50) {
      sheet.deleteRows(51, sheet.getMaxRows() - 50);
   }
}

/**
 * Crée la bannière avec titre.
 */
function creerBanniereTitre(sheet, niveau, couleurs) {
  Logger.log(" - Création bannière titre...");
  sheet.setRowHeight(1, 80);
  const rangeTitre = sheet.getRange("B1:F1"); // Colonnes B à F, Ligne 1
  rangeTitre.merge();
  rangeTitre.setBackground(couleurs.principale);
  rangeTitre.setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID);

  // Sécurisation de la variable niveau pour éviter les erreurs si elle est undefined
  const niveauStr = (typeof niveau === "string" && niveau) ? niveau : "?";
  const titre = `SYSTÈME DE RÉPARTITION DES CLASSES\nNIVEAU ${niveauStr.toUpperCase()} - ${getConfig().VERSION || 'V?'}`; // Utiliser getConfig().VERSION

  const celleTitre = sheet.getRange("B1"); // Cellule B1 contient la valeur après fusion
  celleTitre.setValue(titre)
            .setFontColor(couleurs.texteBlanc)
            .setFontWeight("bold")
            .setFontSize(18)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");

  sheet.setRowHeight(2, 20); // Espacement
}

/**
 * Crée la section menu.
 */
function creerSectionMenu(sheet, couleurs) {
   Logger.log(" - Création section menu...");
  sheet.setRowHeight(3, 70);
  const rangeMenu = sheet.getRange("B3:F3");
  rangeMenu.merge();
  rangeMenu.setBackground(couleurs.secondaire);
  rangeMenu.setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID);

  const texteMenu = "LOCALISATION DU MENU:\n\nFichier  Édition  Affichage  [...]  Extensions  Aide         ►  🎓 Répartition  ◄";

  const celleMenu = sheet.getRange("B3");
  celleMenu.setValue(texteMenu)
           .setFontSize(12)
           .setVerticalAlignment("middle")
           .setHorizontalAlignment("center")
           .setWrap(true); // Permettre retour à la ligne si nécessaire

  sheet.setRowHeight(4, 20); // Espacement
}

/**
 * Applique le format RichText au menu (isolée pour robustesse).
 */
function appliquerFormatRichTextMenu(sheet, couleurs) {
   Logger.log("   - Tentative application RichText au menu...");
  try {
      const celleMenu = sheet.getRange("B3");
      const texteMenu = String(celleMenu.getValue()); // Forcer en chaîne
      const indexRepartDebut = texteMenu.indexOf("🎓 Répartition");
      const indexRepartFin = texteMenu.lastIndexOf("◄"); // Trouver la fin

      if (indexRepartDebut !== -1 && indexRepartFin > indexRepartDebut) {
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(texteMenu)
            .setTextStyle(indexRepartDebut, indexRepartFin + 1, // Inclure le dernier caractère
                         SpreadsheetApp.newTextStyle()
                            .setBold(true)
                            .setForegroundColor(couleurs.principale)
                            .setFontSize(14)
                            .build())
            .build();
          celleMenu.setRichTextValue(richText);
          Logger.log("   - Format RichText appliqué au menu.");
      } else {
          Logger.log("   - Marqueurs '🎓 Répartition' / '◄' non trouvés pour RichText.");
      }
  } catch(e) {
       Logger.log(`   - Erreur RichText menu: ${e}`);
  }
}

/**
 * Crée la section À propos.
 */
function creerSectionAPropos(sheet, couleurs) {
  Logger.log(" - Création section À Propos...");
  // Titre
  sheet.setRowHeight(5, 40);
  const rangeTitre = sheet.getRange("B5:F5");
  rangeTitre.merge()
            .setBackground(couleurs.principale)
            .setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID)
            .setValue("À PROPOS DU SYSTÈME")
            .setFontWeight("bold")
            .setFontSize(14)
            .setFontColor(couleurs.texteBlanc)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");

  // Contenu
  sheet.setRowHeight(6, 140); // Ajuster hauteur si besoin
  const rangeContenu = sheet.getRange("B6:F6");
  const texteAPropos =
    "  Ce système permet:\n\n" +
    "  • Répartition semi-automatisée des élèves en classes.\n" +
    "  • Gestion des options et des contraintes (associations/dissociations).\n" +
    "  • Équilibrage selon divers critères (sexe, LV2, options, notes...). \n" +
    "  • Optimisation pour des classes homogènes et équilibrées.\n" +
    "  • Génération de bilans et statistiques pour analyse.\n\n" +
    "  Utilisez le menu '🎓 Répartition' pour accéder aux fonctionnalités.";
  rangeContenu.merge()
              .setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID)
              .setValue(texteAPropos)
              .setFontSize(11)
              .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
              .setVerticalAlignment("top")
              .setHorizontalAlignment("left")
              .setBackground(couleurs.texteBlanc); // Assurer fond blanc

  sheet.setRowHeight(7, 20); // Espacement
}

/**
 * Crée la section Démarrage rapide.
 * @return {number} La ligne suivant cette section.
 */
function creerSectionDemarrageRapide(sheet, couleurs) {
  Logger.log(" - Création section Démarrage Rapide...");
  let row = 8; // Ligne de départ
  // Titre
  sheet.setRowHeight(row, 40);
  sheet.getRange(row, 2, 1, 5).merge() // B8:F8
      .setBackground(couleurs.principale).setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID)
      .setValue("DÉMARRAGE RAPIDE").setFontWeight("bold").setFontSize(14).setFontColor(couleurs.texteBlanc)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  row++; // 9

  const etapes = [
    { i: "1️⃣", t: "INITIALISATION", d: "Vous venez de le faire (ou utilisez le menu). Définit la base." },
    { i: "2️⃣", t: "CONFIGURATION", d: "Ajuster la structure (_STRUCTURE), pondérations (via Console/Admin)." },
    { i: "3️⃣", t: "DONNÉES ÉLÈVES", d: "Importer/Coller les listes dans les onglets sources (ex: 6°1)." },
    { i: "4️⃣", t: "CONSOLE", d: "Utiliser la Console de Répartition pour consolider, répartir, optimiser." },
    { i: "5️⃣", t: "FINALISATION", d: "Générer les onglets DEF, bilans et statistiques." }
  ];

  // En-têtes Tableau
  sheet.getRange(row, 2).setValue("ÉTAPE").setBackground(couleurs.secondaire).setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID); // B9
  sheet.getRange(row, 3, 1, 4).merge().setValue("DESCRIPTION").setBackground(couleurs.secondaire).setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID); // C9:F9
  row++; // 10

  // Lignes Tableau
  etapes.forEach(etape => {
    sheet.setRowHeight(row, 35); // Hauteur pour chaque étape
    sheet.getRange(row, 2).setValue(`${etape.i} ${etape.t}`).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID); // B10, B11...
    sheet.getRange(row, 3, 1, 4).merge().setValue("  " + etape.d).setHorizontalAlignment("left").setVerticalAlignment("middle").setWrap(true).setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID); // C10:F10, C11:F11...
    row++;
  });

  sheet.setRowHeight(row, 20); // Espacement
  row++;
  return row; // Retourne la prochaine ligne disponible
}

/**
 * Crée la section Documentation.
 * @param {Sheet} sheet La feuille.
 * @param {number} startRow La ligne où commencer.
 * @param {Object} couleurs Les couleurs.
 * @return {number} La ligne suivant cette section.
 */
function creerSectionDocumentation(sheet, startRow, couleurs) {
  Logger.log(" - Création section Documentation...");
  let row = startRow;
  // Titre
  sheet.setRowHeight(row, 40);
  sheet.getRange(row, 2, 1, 5).merge() // Bx:Fx
      .setBackground(couleurs.principale).setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID)
      .setValue("DOCUMENTATION & AIDE").setFontWeight("bold").setFontSize(14).setFontColor(couleurs.texteBlanc)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  row++;

  // Contenu (exemple simple)
   sheet.setRowHeight(row, 60);
   const texteDoc = "🔗 Lien vers la documentation complète (à insérer)\n" +
                    "🆘 Contacter le support / l'administrateur en cas de problème.";
   sheet.getRange(row, 2, 1, 5).merge()
       .setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID)
       .setValue(texteDoc)
       .setWrap(true)
       .setVerticalAlignment("middle")
       .setHorizontalAlignment("center")
       .setBackground(couleurs.texteBlanc);
  row++;


  /* // --- Section plus détaillée (si liens disponibles) ---
  const documents = [ // Mettez vos vrais liens ici
    { n: "Manuel Utilisateur", u: "#", d: "Guide complet du système." },
    { n: "Guide Rapide", u: "#", d: "Étapes essentielles." },
    { n: "FAQ / Aide", u: "#", d: "Réponses aux questions fréquentes." }
  ];
  // En-têtes Tableau
  sheet.getRange(row, 2).setValue("DOCUMENT").setBackground(couleurs.secondaire).setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(row, 3, 1, 4).merge().setValue("DESCRIPTION").setBackground(couleurs.secondaire).setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID);
  row++;
  // Lignes Tableau
  documents.forEach(doc => {
     sheet.setRowHeight(row, 30);
    const cellLien = sheet.getRange(row, 2);
    if (doc.u !== "#") { // Si URL valide
        cellLien.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(doc.n).setLinkUrl(doc.u).build());
    } else {
        cellLien.setValue(doc.n);
    }
    cellLien.setFontColor(couleurs.principale).setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(row, 3, 1, 4).merge().setValue("  " + doc.d).setHorizontalAlignment("left").setVerticalAlignment("middle").setWrap(true).setBorder(true, true, true, true, false, false, couleurs.bordure, SpreadsheetApp.BorderStyle.SOLID);
    row++;
  });
  */

  sheet.setRowHeight(row, 20); // Espacement
  row++;
  return row; // Prochaine ligne disponible
}

/**
 * Crée le bouton d'action et le pied de page.
 */
function creerBoutonEtPiedDePage(sheet, startRow, couleurs) {
  Logger.log(" - Création bouton action et pied de page...");
  let row = startRow;

  // Bouton/Indication Action
  sheet.setRowHeight(row, 50);
  sheet.getRange(row, 2, 1, 5).merge() // Bx:Fx
      .setBackground(couleurs.accent).setBorder(true, true, true, true, false, false, '#e0a800', SpreadsheetApp.BorderStyle.SOLID)
      .setValue("UTILISEZ LE MENU '🎓 Répartition' POUR CONTINUER")
      .setFontWeight("bold").setFontSize(14).setFontColor(couleurs.texteBlanc) // Texte Blanc sur Jaune
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  row++;

  // Pied de page
  sheet.setRowHeight(row, 30);
  sheet.getRange(row, 2, 1, 5).merge() // Bx+1:Fx+1
      .setValue(`Système de Répartition ${getConfig().VERSION || 'V?'} - ${new Date().getFullYear()}`)
      .setFontStyle("italic").setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle");
  row++;
}


/**
 * Dernière tentative de mise en forme du menu RichText.
 * Exécutée à la fin de la création de l'onglet.
 */
function tentativeFinaleFormatMenu(sheet, couleurs) {
  // Cette fonction est appelée par creerOngletPresentation
  // et applique le formatage RichText au menu.
  // Le code précédent était correct, on le garde.
   Logger.log(" - Tentative finale format menu (RichText)...");
  try {
    // SpreadsheetApp.flush(); // Pas forcément nécessaire ici

    const celleMenu = sheet.getRange("B3");
    const texteMenu = String(celleMenu.getValue());
    const indexRepartDebut = texteMenu.indexOf("🎓 Répartition");
     const indexRepartFin = texteMenu.lastIndexOf("◄");

    if (indexRepartDebut !== -1 && indexRepartFin > indexRepartDebut) {
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(texteMenu)
        .setTextStyle(indexRepartDebut, indexRepartFin + 1,
                     SpreadsheetApp.newTextStyle()
                        .setBold(true)
                        .setForegroundColor(couleurs.principale)
                        .setFontSize(14)
                        .build())
        .build();
      celleMenu.setRichTextValue(richText);
       Logger.log("   - RichText appliqué.");
    } else {
        Logger.log("   - Marqueurs RichText non trouvés.");
    }

    // Ajouter une note informative peut toujours être utile
    const noteMenu = "Utilisez le menu '🎓 Répartition' ci-dessus pour accéder aux fonctions principales.";
    sheet.getRange("B3").setNote(noteMenu);
    // Logger.log("   - Note ajoutée à la cellule menu.");

  } catch (e) {
    Logger.log("   - Erreur lors tentative finale format menu: " + e);
  }
}