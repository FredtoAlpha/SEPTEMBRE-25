/**
 * ===============================================================
 * RÉÉQUILIBRAGE FORCÉ DES EFFECTIFS
 * ===============================================================
 * 
 * Fonction d'urgence pour corriger les déséquilibres importants
 * Priorité absolue : ÉQUILIBRER LES EFFECTIFS
 */

function reequilibrerEffectifsForce() {
  try {
    Logger.log('\n' + '═'.repeat(70));
    Logger.log('🚨 RÉÉQUILIBRAGE FORCÉ DES EFFECTIFS');
    Logger.log('═'.repeat(70));
    
    const config = getConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Analyser la situation actuelle
    Logger.log('\n📊 ANALYSE DES EFFECTIFS ACTUELS...');
    const analyseSituation = analyserEffectifsActuels();
    
    Logger.log('\n📈 SITUATION ACTUELLE:');
    analyseSituation.classes.forEach(c => {
      Logger.log(`   ${c.nom}: ${c.effectif} élèves ${c.effectif > analyseSituation.moyenneIdeal + 2 ? '⚠️ SURCHARGE' : ''}`);
    });
    Logger.log(`   Moyenne idéale: ${analyseSituation.moyenneIdeal}`);
    Logger.log(`   Écart max: ${analyseSituation.ecartMax}`);
    
    if (analyseSituation.ecartMax <= 2) {
      Logger.log('\n✅ Les effectifs sont déjà équilibrés (écart ≤ 2)');
      return;
    }
    
    // 2. Identifier les déséquilibres
    const classesProblemes = identifierClassesProblemes(analyseSituation);
    
    // 3. Corriger les déséquilibres
    let totalDeplacements = 0;
    
    classesProblemes.surchargees.forEach(classeSource => {
      const nbADeplacer = classeSource.effectif - analyseSituation.moyenneIdeal - 1;
      Logger.log(`\n🔄 Traitement ${classeSource.nom}: ${nbADeplacer} élèves à déplacer`);
      
      const deplacements = deplacerElevesSurcharge(
        classeSource.nom, 
        nbADeplacer, 
        classesProblemes.sousEffectif,
        analyseSituation
      );
      
      totalDeplacements += deplacements;
    });
    
    // 4. Vérification finale
    const analyseFinal = analyserEffectifsActuels();
    Logger.log('\n📊 RÉSULTAT FINAL:');
    analyseFinal.classes.forEach(c => {
      Logger.log(`   ${c.nom}: ${c.effectif} élèves`);
    });
    Logger.log(`   Écart final: ${analyseFinal.ecartMax}`);
    
    Logger.log(`\n✅ Rééquilibrage terminé: ${totalDeplacements} déplacements effectués`);
    
    SpreadsheetApp.getUi().alert(
      'Rééquilibrage terminé',
      `${totalDeplacements} élèves déplacés.\nÉcart final: ${analyseFinal.ecartMax}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    Logger.log(`❌ ERREUR: ${e.message}`);
    Logger.log(e.stack);
    SpreadsheetApp.getUi().alert('Erreur', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Analyse les effectifs actuels directement depuis les feuilles
 */
function analyserEffectifsActuels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().includes('TEST'));
  
  const classes = [];
  let totalEleves = 0;
  
  sheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    // Compter les lignes avec un ID_ELEVE non vide
    let effectif = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim() !== '') {
        effectif++;
      }
    }
    
    classes.push({
      nom: sheet.getName(),
      effectif: effectif,
      sheet: sheet
    });
    
    totalEleves += effectif;
  });
  
  const moyenneIdeal = Math.round(totalEleves / classes.length);
  const effectifs = classes.map(c => c.effectif);
  const ecartMax = Math.max(...effectifs) - Math.min(...effectifs);
  
  return {
    classes: classes.sort((a, b) => b.effectif - a.effectif),
    totalEleves: totalEleves,
    moyenneIdeal: moyenneIdeal,
    ecartMax: ecartMax
  };
}

/**
 * Identifie les classes surchargées et en sous-effectif
 */
function identifierClassesProblemes(analyse) {
  const surchargees = [];
  const sousEffectif = [];
  
  analyse.classes.forEach(classe => {
    if (classe.effectif > analyse.moyenneIdeal + 1) {
      surchargees.push(classe);
    } else if (classe.effectif < analyse.moyenneIdeal - 1) {
      sousEffectif.push(classe);
    }
  });
  
  return { surchargees, sousEffectif };
}

/**
 * Déplace des élèves d'une classe surchargée
 */
function deplacerElevesSurcharge(classeSource, nbADeplacer, classesDestination, analyse) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSource = ss.getSheetByName(classeSource);
  
  if (!sheetSource) {
    Logger.log(`❌ Feuille ${classeSource} introuvable`);
    return 0;
  }
  
  // Charger les données de la classe source
  const dataSource = sheetSource.getDataRange().getValues();
  const headers = dataSource[0].map(h => String(h).trim().toUpperCase());
  
  // Indices des colonnes importantes
  const indices = {
    id: headers.indexOf('ID_ELEVE'),
    nom: headers.indexOf('NOM'),
    mobilite: headers.indexOf('MOBILITE'),
    lv2: headers.indexOf('LV2'),
    opt: headers.indexOf('OPT'),
    com: headers.indexOf('COM'),
    asso: headers.indexOf('ASSO'),
    disso: headers.indexOf('DISSO')
  };
  
  // Identifier les élèves déplaçables (LIBRE et CONDI uniquement)
  const elevesDeplacables = [];
  
  for (let i = 1; i < dataSource.length; i++) {
    const ligne = dataSource[i];
    if (!ligne[indices.id] || String(ligne[indices.id]).trim() === '') continue;
    
    const mobilite = indices.mobilite !== -1 ? String(ligne[indices.mobilite] || 'LIBRE').toUpperCase() : 'LIBRE';
    
    // On ne déplace que LIBRE et CONDI
    if (mobilite === 'LIBRE' || mobilite === 'CONDI') {
      elevesDeplacables.push({
        ligne: i,
        data: ligne,
        id: ligne[indices.id],
        nom: ligne[indices.nom],
        mobilite: mobilite,
        com: parseInt(ligne[indices.com]) || 4,
        asso: ligne[indices.asso] || '',
        disso: ligne[indices.disso] || '',
        lv2: ligne[indices.lv2] || '',
        opt: ligne[indices.opt] || ''
      });
    }
  }
  
  Logger.log(`   → ${elevesDeplacables.length} élèves déplaçables trouvés`);
  
  // Trier par COM décroissant (on déplace d'abord les élèves faciles)
  elevesDeplacables.sort((a, b) => b.com - a.com);
  
  // Déplacer les élèves
  let nbDeplaces = 0;
  const elevesASupprimer = [];
  
  for (const eleve of elevesDeplacables) {
    if (nbDeplaces >= nbADeplacer) break;
    
    // Trouver la meilleure destination
    let meilleureDestination = null;
    let minEffectif = analyse.moyenneIdeal + 1;
    
    for (const classeDestObj of classesDestination) {
      if (classeDestObj.effectif < minEffectif) {
        // Vérifier les contraintes basiques
        if (peutDeplacerVers(eleve, classeDestObj.nom, ss)) {
          meilleureDestination = classeDestObj;
          minEffectif = classeDestObj.effectif;
        }
      }
    }
    
    if (meilleureDestination) {
      // Effectuer le déplacement
      const sheetDest = ss.getSheetByName(meilleureDestination.nom);
      if (sheetDest) {
        // Ajouter à la destination
        const lastRowDest = sheetDest.getLastRow();
        sheetDest.getRange(lastRowDest + 1, 1, 1, eleve.data.length).setValues([eleve.data]);
        
        // Marquer pour suppression de la source
        elevesASupprimer.push(eleve.ligne);
        
        // Mettre à jour les effectifs
        meilleureDestination.effectif++;
        
        Logger.log(`      → ${eleve.nom} déplacé vers ${meilleureDestination.nom}`);
        nbDeplaces++;
      }
    }
  }
  
  // Supprimer les élèves déplacés de la source (en partant de la fin)
  elevesASupprimer.sort((a, b) => b - a);
  elevesASupprimer.forEach(ligne => {
    sheetSource.deleteRow(ligne + 1); // +1 car les indices commencent à 0
  });
  
  return nbDeplaces;
}

/**
 * Vérifie si un élève peut être déplacé vers une classe
 */
function peutDeplacerVers(eleve, classeDestination, ss) {
  const sheet = ss.getSheetByName(classeDestination);
  if (!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const dissoIndex = headers.indexOf('DISSO');
  
  // Vérifier les contraintes DISSO
  if (eleve.disso && dissoIndex !== -1) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][dissoIndex] === eleve.disso) {
        return false; // Conflit DISSO
      }
    }
  }
  
  // Pour l'instant, on ne vérifie que DISSO
  // On pourrait ajouter d'autres vérifications si nécessaire
  
  return true;
}

// ===============================================================
// FONCTION ALTERNATIVE : AFFICHER UNIQUEMENT L'ANALYSE
// ===============================================================

function analyserEffectifsUniquement() {
  const analyse = analyserEffectifsActuels();
  
  Logger.log('\n' + '═'.repeat(70));
  Logger.log('📊 ANALYSE DES EFFECTIFS');
  Logger.log('═'.repeat(70));
  
  Logger.log('\n📈 RÉPARTITION ACTUELLE:');
  analyse.classes.forEach(c => {
    const ecart = c.effectif - analyse.moyenneIdeal;
    const signe = ecart >= 0 ? '+' : '';
    const alerte = Math.abs(ecart) > 2 ? ' ⚠️' : '';
    
    Logger.log(`   ${c.nom}: ${c.effectif} élèves (${signe}${ecart})${alerte}`);
  });
  
  Logger.log(`\n📊 STATISTIQUES:`);
  Logger.log(`   Total élèves: ${analyse.totalEleves}`);
  Logger.log(`   Moyenne idéale: ${analyse.moyenneIdeal}`);
  Logger.log(`   Écart maximum: ${analyse.ecartMax}`);
  
  // Identifier les élèves FIXE/SPEC par classe
  Logger.log('\n🔒 ÉLÈVES IMMOBILES PAR CLASSE:');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  analyse.classes.forEach(classe => {
    const sheet = ss.getSheetByName(classe.nom);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toUpperCase());
    const mobIndex = headers.indexOf('MOBILITE');
    
    if (mobIndex === -1) return;
    
    let nbFixe = 0, nbSpec = 0;
    for (let i = 1; i < data.length; i++) {
      const mob = String(data[i][mobIndex] || '').toUpperCase();
      if (mob === 'FIXE') nbFixe++;
      if (mob === 'SPEC') nbSpec++;
    }
    
    if (nbFixe > 0 || nbSpec > 0) {
      Logger.log(`   ${classe.nom}: ${nbFixe} FIXE, ${nbSpec} SPEC`);
    }
  });
  
  // Recommandations
  if (analyse.ecartMax > 3) {
    Logger.log('\n⚠️ RECOMMANDATION: Écart trop important, rééquilibrage nécessaire');
    Logger.log('   Exécutez reequilibrerEffectifsForce() pour corriger');
  } else if (analyse.ecartMax > 2) {
    Logger.log('\n⚠️ RECOMMANDATION: Écart modéré, rééquilibrage conseillé');
  } else {
    Logger.log('\n✅ Les effectifs sont bien équilibrés');
  }
}

// ===============================================================
// MENU PERSONNALISÉ
// ===============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🔧 Rééquilibrage')
    .addItem('📊 Analyser les effectifs', 'analyserEffectifsUniquement')
    .addItem('⚖️ Rééquilibrer (forcé)', 'reequilibrerEffectifsForce')
    .addSeparator()
    .addItem('📋 Documentation', 'afficherDocumentationReequilibrage')
    .addToUi();
}

function afficherDocumentationReequilibrage() {
  const html = `
    <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 500px;">
      <h2>🔧 Rééquilibrage des Effectifs</h2>
      
      <h3>📊 Analyser les effectifs</h3>
      <p>Affiche la répartition actuelle des élèves par classe et identifie les déséquilibres.</p>
      
      <h3>⚖️ Rééquilibrer (forcé)</h3>
      <p>Déplace automatiquement des élèves des classes surchargées vers les classes en sous-effectif.</p>
      <ul>
        <li>Ne déplace que les élèves LIBRE et CONDI</li>
        <li>Respecte les contraintes DISSO</li>
        <li>Privilégie le déplacement des élèves avec bon comportement (COM élevé)</li>
      </ul>
      
      <h3>⚠️ Important</h3>
      <p>Les élèves FIXE et SPEC ne sont JAMAIS déplacés.</p>
      <p>Faites une sauvegarde avant d'utiliser le rééquilibrage forcé.</p>
    </div>
  `;
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(600).setHeight(400),
    'Documentation Rééquilibrage'
  );
}

/*
===============================================================
UTILISATION
===============================================================

1. ANALYSE SIMPLE:
   analyserEffectifsUniquement()
   → Affiche l'état actuel sans rien modifier

2. RÉÉQUILIBRAGE FORCÉ:
   reequilibrerEffectifsForce()
   → Déplace automatiquement des élèves pour équilibrer

3. DEPUIS LE MENU:
   Menu "🔧 Rééquilibrage" dans Google Sheets

PRINCIPE:
- Priorité absolue aux effectifs équilibrés
- Ne déplace que LIBRE et CONDI
- Respecte les contraintes DISSO
- Déplace d'abord les élèves avec bon comportement

===============================================================
*/