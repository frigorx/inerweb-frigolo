/**
 * inerWeb FRIGOLO — Réception des résultats du module Conversion Pouce & Cintrage
 *
 * INSTALLATION :
 * 1. Créer un Google Sheet (ou utiliser un existant)
 * 2. Extensions > Apps Script
 * 3. Coller ce code dans Code.gs
 * 4. Déployer > Nouveau déploiement > Application Web
 *    - Exécuter en tant que : Moi
 *    - Accès : Tout le monde
 * 5. Copier l'URL du déploiement
 * 6. Coller l'URL dans le fichier HTML (variable APPS_SCRIPT_URL)
 */

const SHEET_NAME = 'Résultats';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Créer la feuille avec en-têtes si elle n'existe pas
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Horodatage',
        'Nom',
        'Prénom',
        'Classe',
        'Phase 1 (%)',
        'Phase 2 (%)',
        'Phase 3 (%)',
        'Phase 4 (%)',
        'Phase 5 (%)',
        'Phase 6 (%)',
        'Évaluation (%)',
        'Détail',
        'Temps (s)',
        'Appréciation'
      ]);

      // Mise en forme en-têtes
      const headerRange = sheet.getRange(1, 1, 1, 14);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#0f2d52');
      headerRange.setFontColor('#ffffff');
      sheet.setFrozenRows(1);

      // Largeur colonnes
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidths(2, 3, 120);
      sheet.setColumnWidths(5, 7, 90);
      sheet.setColumnWidth(12, 100);
      sheet.setColumnWidth(14, 160);
    }

    // Appréciation automatique
    const score = data.scoreEval || 0;
    let appreciation = 'Non acquis';
    if (score >= 85) appreciation = 'Maîtrisé';
    else if (score >= 70) appreciation = 'Acquis';
    else if (score >= 50) appreciation = "En cours d'acquisition";

    // Formatage temps
    const tempsMin = Math.floor((data.tempsSecondes || 0) / 60);
    const tempsSec = (data.tempsSecondes || 0) % 60;

    // Ajout ligne
    sheet.appendRow([
      new Date(data.timestamp),
      data.nom || '',
      data.prenom || '',
      data.classe || '',
      data.scorePhase1 || '',
      data.scorePhase2 || '',
      data.scorePhase3 || '',
      data.scorePhase4 || '',
      data.scorePhase5 || '',
      data.scorePhase6 || '',
      data.scoreEval || 0,
      data.detail || '',
      data.tempsSecondes || 0,
      appreciation
    ]);

    // Coloration conditionnelle de la dernière ligne
    const lastRow = sheet.getLastRow();
    const evalCell = sheet.getRange(lastRow, 11);
    if (score >= 85) evalCell.setBackground('#d5f5e3');
    else if (score >= 70) evalCell.setBackground('#d4efdf');
    else if (score >= 50) evalCell.setBackground('#fef9e7');
    else evalCell.setBackground('#fdedec');

    return ContentService.createTextOutput(JSON.stringify({
      status: 'ok',
      message: 'Résultat enregistré'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    module: 'inerWeb FRIGOLO — Conversion Pouce & Cintrage',
    message: 'Le service fonctionne. Utilisez POST pour envoyer des résultats.'
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Menu personnalisé pour stats rapides
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('inerWeb')
    .addItem('Statistiques', 'showStats')
    .addItem('Réinitialiser la feuille', 'resetSheet')
    .addToUi();
}

function showStats() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Aucun résultat enregistré.');
    return;
  }

  const data = sheet.getRange(2, 11, sheet.getLastRow() - 1, 1).getValues().flat().filter(v => v !== '');
  const n = data.length;
  const avg = data.reduce((a, b) => a + b, 0) / n;
  const max = Math.max(...data);
  const min = Math.min(...data);
  const acquired = data.filter(v => v >= 70).length;

  SpreadsheetApp.getUi().alert(
    'Statistiques Module Conversion & Cintrage\n\n' +
    'Nombre de passages : ' + n + '\n' +
    'Moyenne évaluation : ' + avg.toFixed(1) + '%\n' +
    'Meilleur score : ' + max + '%\n' +
    'Score le plus bas : ' + min + '%\n' +
    'Acquis (≥70%) : ' + acquired + '/' + n + ' (' + Math.round(acquired/n*100) + '%)'
  );
}

function resetSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Attention', 'Supprimer tous les résultats ?', ui.ButtonSet.YES_NO);
  if (response === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (sheet) ss.deleteSheet(sheet);
    ui.alert('Feuille supprimée. Elle sera recréée au prochain envoi.');
  }
}
