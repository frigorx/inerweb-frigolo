/**
 * inerWeb — Réception des résultats du module Conversion Pouce & Cintrage
 * Note /20 + Compétences TNE (NI / ECA / A / M)
 */

const SHEET_NAME = 'Résultats';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = [
        'Horodatage',
        'Nom',
        'Prénom',
        'Classe',
        'Note / 20',
        'Score (%)',
        'Détail',
        'Temps',
        'C1 Conversion',
        'C2 Rayon',
        'C3 Développé',
        'C4 Synthèse',
        'Appréciation'
      ];
      sheet.appendRow(headers);

      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#0f2d52');
      headerRange.setFontColor('#ffffff');
      headerRange.setHorizontalAlignment('center');
      sheet.setFrozenRows(1);

      sheet.setColumnWidth(1, 150);  // Horodatage
      sheet.setColumnWidth(2, 130);  // Nom
      sheet.setColumnWidth(3, 130);  // Prénom
      sheet.setColumnWidth(4, 100);  // Classe
      sheet.setColumnWidth(5, 80);   // Note /20
      sheet.setColumnWidth(6, 80);   // Score %
      sheet.setColumnWidth(7, 80);   // Détail
      sheet.setColumnWidth(8, 80);   // Temps
      sheet.setColumnWidth(9, 110);  // C1
      sheet.setColumnWidth(10, 110); // C2
      sheet.setColumnWidth(11, 110); // C3
      sheet.setColumnWidth(12, 110); // C4
      sheet.setColumnWidth(13, 160); // Appréciation
    }

    // Appréciation globale
    const score = data.scoreEval || 0;
    let appreciation = 'Non acquis';
    if (score >= 85) appreciation = 'Maîtrisé';
    else if (score >= 70) appreciation = 'Acquis';
    else if (score >= 50) appreciation = "En cours d'acquisition";

    // Formatage temps
    const totalSec = data.tempsSecondes || 0;
    const tempsStr = Math.floor(totalSec / 60) + ' min ' + String(totalSec % 60).padStart(2, '0') + ' s';

    // Note /20
    const note20 = data.note20 || 0;

    // Compétences (codes : M, A, ECA, NI)
    function compLabel(code) {
      switch (code) {
        case 'M': return 'Maîtrisé';
        case 'A': return 'Acquis';
        case 'ECA': return "En cours";
        case 'NI': return 'Non acquis';
        default: return code || '';
      }
    }

    sheet.appendRow([
      new Date(data.timestamp),
      data.nom || '',
      data.prenom || '',
      data.classe || '',
      note20,
      score,
      data.detail || '',
      tempsStr,
      compLabel(data.C1_conversion),
      compLabel(data.C2_rayon),
      compLabel(data['C3_développé']),
      compLabel(data.C4_synthese),
      appreciation
    ]);

    // Coloration de la dernière ligne
    const lastRow = sheet.getLastRow();

    // Note /20
    const noteCell = sheet.getRange(lastRow, 5);
    noteCell.setNumberFormat('0.0');
    noteCell.setFontWeight('bold');
    if (note20 >= 14) noteCell.setBackground('#d5f5e3');
    else if (note20 >= 10) noteCell.setBackground('#d4efdf');
    else if (note20 >= 8) noteCell.setBackground('#fef9e7');
    else noteCell.setBackground('#fdedec');

    // Compétences C1-C4 (colonnes 9-12)
    const compColors = {
      'Maîtrisé': '#d5f5e3',
      'Acquis': '#d4efdf',
      'En cours': '#fef9e7',
      'Non acquis': '#fdedec'
    };

    for (let col = 9; col <= 12; col++) {
      const cell = sheet.getRange(lastRow, col);
      const val = cell.getValue();
      if (compColors[val]) cell.setBackground(compColors[val]);
    }

    // Appréciation
    const appCell = sheet.getRange(lastRow, 13);
    if (compColors[appreciation]) appCell.setBackground(compColors[appreciation]);

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
    module: 'inerWeb — Conversion Pouce & Cintrage',
    message: 'Le service fonctionne.'
  })).setMimeType(ContentService.MimeType.JSON);
}

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

  const notes = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getValues().flat().filter(v => v !== '');
  const n = notes.length;
  const avg = notes.reduce((a, b) => a + b, 0) / n;
  const max = Math.max(...notes);
  const min = Math.min(...notes);
  const above10 = notes.filter(v => v >= 10).length;

  // Compter les compétences
  const c1 = sheet.getRange(2, 9, sheet.getLastRow() - 1, 1).getValues().flat();
  const c2 = sheet.getRange(2, 10, sheet.getLastRow() - 1, 1).getValues().flat();
  const c3 = sheet.getRange(2, 11, sheet.getLastRow() - 1, 1).getValues().flat();
  const c4 = sheet.getRange(2, 12, sheet.getLastRow() - 1, 1).getValues().flat();

  function countAcq(arr) { return arr.filter(v => v === 'Acquis' || v === 'Maîtrisé').length; }

  SpreadsheetApp.getUi().alert(
    'Statistiques Module Conversion & Cintrage\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'Élèves évalués : ' + n + '\n\n' +
    'NOTE /20\n' +
    '  Moyenne : ' + avg.toFixed(1) + ' / 20\n' +
    '  Meilleure : ' + max.toFixed(1) + ' / 20\n' +
    '  Plus basse : ' + min.toFixed(1) + ' / 20\n' +
    '  ≥ 10/20 : ' + above10 + '/' + n + ' (' + Math.round(above10/n*100) + '%)\n\n' +
    'COMPÉTENCES (Acquis + Maîtrisé)\n' +
    '  C1 Conversion : ' + countAcq(c1) + '/' + n + '\n' +
    '  C2 Rayon : ' + countAcq(c2) + '/' + n + '\n' +
    '  C3 Développé : ' + countAcq(c3) + '/' + n + '\n' +
    '  C4 Synthèse : ' + countAcq(c4) + '/' + n
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
