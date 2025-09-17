function ouvrirInterfaceDeplacement() {
  const html = HtmlService.createHtmlOutputFromFile("interface_deplacement")
    .setWidth(400)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, "Déplacer ou échanger un élève");
}

function deplacerEleve(id, classeDest) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().includes("TEST"));
  let ligneTrouvee = null;
  let feuilleSource = null;

  for (let sheet of sheets) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === id) {
        ligneTrouvee = data[i];
        feuilleSource = sheet;
        sheet.deleteRow(i + 1);
        break;
      }
    }
    if (ligneTrouvee) break;
  }

  if (!ligneTrouvee) return;
  const feuilleCible = ss.getSheetByName(classeDest);
  feuilleCible.appendRow(ligneTrouvee);
  SpreadsheetApp.flush();
  V2_calculerStatistiquesFinales();
}

function echangerEleves(id1, id2) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().includes("TEST"));
  let ligne1 = null, feuille1 = null, index1 = -1;
  let ligne2 = null, feuille2 = null, index2 = -1;

  for (let sheet of sheets) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const currentId = String(data[i][0]).trim();
      if (currentId === id1) {
        ligne1 = data[i]; feuille1 = sheet; index1 = i;
      } else if (currentId === id2) {
        ligne2 = data[i]; feuille2 = sheet; index2 = i;
      }
    }
  }

  if (ligne1 && ligne2 && feuille1 && feuille2) {
    feuille1.getRange(index1 + 1, 1, 1, ligne2.length).setValues([ligne2]);
    feuille2.getRange(index2 + 1, 1, 1, ligne1.length).setValues([ligne1]);
    SpreadsheetApp.flush();
    V2_calculerStatistiquesFinales();
  }
}

function toggleColonneID() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const colA = feuille.getRange("A:A");
  const estCachee = feuille.isColumnHiddenByUser(1);
  colA.setHidden(!estCachee);
}

function swapManuelDepuisConsole(id1, id2) {
  echangerEleves(id1, id2);
}

function deplacementManuelDepuisConsole(id, classeDest) {
  deplacerEleve(id, classeDest);
}
