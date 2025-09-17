function testeur_V2_preparerDonneesInitiales_Complet() {
  try {
    Logger.log("=== DÉBUT TEST V2_preparerDonneesInitiales_Complet ===");
    const config = getConfig(); // APPELLE VOTRE getConfig() MIS À JOUR

    // Vérification minimale de la config V2 pour ce test
    if (!config.V2_MOBILITES_CONSIDEREES_FIXES || !config.V2_CRITERES_A_EQUILIBRER) {
        Logger.log("ERREUR: config.V2_MOBILITES_CONSIDEREES_FIXES ou config.V2_CRITERES_A_EQUILIBRER manquante.");
        Logger.log("       Veuillez vérifier votre constante CONFIG dans Config.gs et/ou votre feuille _CONFIG.");
        return;
    }
    Logger.log(`  Config V2_MOBILITES_CONSIDEREES_FIXES: ${(config.V2_MOBILITES_CONSIDEREES_FIXES || []).join(', ')}`);
    Logger.log(`  Config V2_CRITERES_A_EQUILIBRER: ${(config.V2_CRITERES_A_EQUILIBRER || []).join(', ')}`);
    Logger.log(`  Config V2_SHEET_NAME_OPTI_STRUCTURE: ${config.V2_SHEET_NAME_OPTI_STRUCTURE}`);


    const dataContext = V2_preparerDonneesInitiales(config); // APPEL DE LA FONCTION À TESTER

    if (dataContext) {
      Logger.log("SUCCÈS: V2_preparerDonneesInitiales a retourné un dataContext.");
      Logger.log("  Vérifiez la feuille '" + (config.V2_SHEET_NAME_OPTI_STRUCTURE || "_NirvanaV2_Calculs") + "' pour le détail des calculs.");

      // Afficher quelques éléments clés du dataContext pour vérification rapide
      Logger.log(`  Nombre d'élèves valides: ${dataContext.elevesValides?.length}`);
      Logger.log(`  Noms des classes: ${dataContext.classNames?.join(', ')}`);

      if (dataContext.classNames && dataContext.classNames.length > 0) {
          const exempleClasse = dataContext.classNames[0];
          Logger.log(`  Exemple CountsState pour ${exempleClasse}: ${JSON.stringify(dataContext.countsState?.[exempleClasse], null, 2)}`);
          Logger.log(`  Exemple TargetsStateExtremes pour ${exempleClasse}: ${JSON.stringify(dataContext.targetsStateExtremes?.[exempleClasse], null, 2)}`);
          Logger.log(`  Exemple DeltasStateExtremes pour ${exempleClasse}: ${JSON.stringify(dataContext.deltasStateExtremes?.[exempleClasse], null, 2)}`);
      }
    } else {
      Logger.log("ÉCHEC: V2_preparerDonneesInitiales a retourné null ou une erreur a été levée avant de retourner.");
    }
    Logger.log("=== FIN TEST V2_preparerDonneesInitiales_Complet ===");

  } catch (e) {
    Logger.log(`ERREUR CRITIQUE dans testeur_V2_preparerDonneesInitiales_Complet: ${e.message}\n${e.stack}`);
  }
}
function diagnosticClasses() {
  const cfg   = getConfig();
  const ctx   = V2_Ameliore_PreparerDonnees(cfg);

  Logger.log('=== DIAGNOSTIC CLASSES ===');
  Object.keys(ctx.classesState).forEach(c => {
    const n   = ctx.classesState[c].length;
    const has = !!ctx.ciblesParClasse[c];
    Logger.log(
      JSON.stringify(c) + ' : ' + n + ' élèves,   cibles définies : ' + (has ? 'OUI' : 'NON')
    );
  });
}