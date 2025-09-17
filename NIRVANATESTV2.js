function diagnosticNirvanaScores() {
  // 1. Préparation contextuelle
  const config = getConfig();
  const dataContext = V2_Ameliore_PreparerDonnees(config);

  Logger.log("=== DIAGNOSTIC NIRVANA SCORES ===");

  // 2. Affichage des classes et du 1er élève de chaque classe
  Logger.log("--- CLASSES DÉTECTÉES ---");
  Object.entries(dataContext.classesState).forEach(([classe, eleves]) => {
    Logger.log(`Classe ${classe} (${eleves.length} élèves)`);
    if (eleves.length > 0) {
      Logger.log("  Premier élève: " + JSON.stringify(eleves[0], null, 2));
    }
  });

  // 3. Affichage mobilité et score pour chaque élève (raccourci sur 10 élèves)
  Logger.log("--- MOBILITÉ ET SCORES (1ers élèves) ---");
  let preview = 0;
  Object.values(dataContext.classesState).flat().slice(0, 10).forEach(e => {
    Logger.log(
      `${e.ID_ELEVE} | MOBILITE: ${e.MOBILITE} | COM: ${e.COM} | TRA: ${e.TRA} | PART: ${e.PART} | ABS: ${e.ABS} | LV2: ${e.LV2} | OPT: ${e.OPT}`
    );
    preview++;
  });
  if (preview === 0) Logger.log("Aucun élève détecté !");

  // 4. Calcul des effectifs cibles et actuels
  const effectifsCibles = calculerEffectifsCibles_Ameliore(dataContext, config);
  const effectifsActuels = calculerEffectifsActuels(dataContext);

  Logger.log("--- EFFECTIFS CIBLES (extrait) ---");
  Object.entries(effectifsCibles).forEach(([classe, cible]) => {
    Logger.log(`Classe ${classe}: ${JSON.stringify(cible.effectifsCibles, null, 2)}`);
  });

  Logger.log("--- EFFECTIFS ACTUELS (extrait) ---");
  Object.entries(effectifsActuels).forEach(([classe, actuel]) => {
    Logger.log(`Classe ${classe}: ${JSON.stringify(actuel, null, 2)}`);
  });

  // 5. Diagnostic des écarts et scores qui devraient "bouger"
  const ecarts = calculerEcartsEffectifs(effectifsActuels, effectifsCibles);
  let swapsPotentiels = 0;
  Logger.log("--- ÉCARTS DÉTECTÉS ---");
  Object.entries(ecarts).forEach(([classe, crits]) => {
    Object.entries(crits).forEach(([critere, scores]) => {
      Object.entries(scores).forEach(([score, obj]) => {
        if (Math.abs(obj.ecart) > 0) {
          swapsPotentiels++;
          Logger.log(
            `Classe ${classe} - ${critere} score ${score}: Ecart ${obj.ecart} (Actuel: ${obj.actuel} / Cible: ${obj.cible})`
          );
        }
      });
    });
  });

  if (swapsPotentiels === 0) {
    Logger.log("⚠️ Aucun écart détecté : soit tes effectifs sont déjà parfaitement équilibrés, soit les colonnes sont mal lues !");
  } else {
    Logger.log(`Nombre de scores déséquilibrés (où un swap devrait être tenté) : ${swapsPotentiels}`);
  }

  // 6. Vérification de la mobilité réelle détectée
  const elevesMobiles = Object.values(dataContext.classesState).flat().filter(e => estEleveMobile(e, dataContext));
  Logger.log(`Nombre d'élèves mobiles détectés : ${elevesMobiles.length} / ${Object.values(dataContext.classesState).flat().length}`);

  // 7. Résumé diagnostic
  if (elevesMobiles.length === 0) Logger.log("⚠️ Aucun élève détecté comme mobile (colonne T = MOBILITE ?) !");
  Logger.log("=== FIN DIAGNOSTIC NIRVANA SCORES ===");
}