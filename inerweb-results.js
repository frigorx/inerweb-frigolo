/**
 * inerWeb — Envoi universel des résultats élèves
 * Inclure ce fichier dans n'importe quel module HTML inerWeb
 *
 * USAGE :
 *   <script src="../inerweb-results.js"></script>
 *   <script>
 *     inerwebSend({
 *       module: 'Nom du module',
 *       nom: 'Dupont',
 *       prenom: 'Jean',
 *       classe: '2 TNE',
 *       note20: 15.5,
 *       score: 78,
 *       detail: '12/15',
 *       temps: 480,
 *       competences: {
 *         C1: 'Maîtrisé',    // ou 'M'
 *         C2: 'Acquis',       // ou 'A'
 *         C3: 'En cours',     // ou 'ECA'
 *         C4: 'Non acquis'    // ou 'NI'
 *       }
 *     });
 *   </script>
 */

(function(window) {
  'use strict';

  // ════════════════════════════════════════════
  // URL DU COLLECTEUR UNIVERSEL
  // ════════════════════════════════════════════
  var SHEET_URL = 'https://script.google.com/macros/s/AKfycbz5Bkn1tacs98bJezjnnYt38Yuy6QiHh7qWuEk1KRxS4UMIjl0yFOA0FVakLwCAJhZ5/exec';

  // Clé localStorage
  var STORAGE_KEY = 'inerweb-results';

  /**
   * Envoyer un résultat
   * @param {Object} data - Les données du résultat
   * @returns {Promise}
   */
  function inerwebSend(data) {
    // Compléter les données
    data.timestamp = data.timestamp || new Date().toISOString();

    // Calculer note /20 si pas fournie
    if (data.note20 == null && data.score != null) {
      data.note20 = Math.round(data.score / 100 * 20 * 2) / 2;
    }

    // 1. Toujours sauvegarder en local
    inerwebSaveLocal(data);

    // 2. Envoyer au Sheet
    return fetch(SHEET_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    }).then(function() {
      console.log('[inerWeb] Résultat envoyé + sauvegardé localement');
      return { status: 'sent' };
    }).catch(function(err) {
      console.warn('[inerWeb] Envoi échoué, sauvegardé localement', err);
      return { status: 'local_only' };
    });
  }

  /**
   * Sauvegarder en local
   */
  function inerwebSaveLocal(data) {
    try {
      var results = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
      results.push(data);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(results));
    } catch(e) {
      console.warn('[inerWeb] Impossible de sauvegarder localement', e);
    }
  }

  /**
   * Lire les résultats locaux
   * @returns {Array}
   */
  function inerwebGetLocal() {
    try {
      return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    } catch(e) {
      return [];
    }
  }

  /**
   * Renvoyer tous les résultats locaux vers le Sheet
   * @returns {Promise}
   */
  function inerwebRetryAll() {
    var results = inerwebGetLocal();
    if (results.length === 0) return Promise.resolve({ sent: 0 });

    var promises = results.map(function(data, i) {
      return new Promise(function(resolve) {
        setTimeout(function() {
          fetch(SHEET_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
          }).then(function() { resolve(true); }).catch(function() { resolve(false); });
        }, i * 400);
      });
    });

    return Promise.all(promises).then(function(results) {
      var sent = results.filter(function(r) { return r; }).length;
      console.log('[inerWeb] ' + sent + '/' + results.length + ' renvoyé(s)');
      return { sent: sent, total: results.length };
    });
  }

  /**
   * Exporter en CSV
   */
  function inerwebExportCSV(filename) {
    var results = inerwebGetLocal();
    if (results.length === 0) { alert('Aucun résultat.'); return; }

    var headers = ['Date','Module','Nom','Prénom','Classe','Note /20','Score %','Détail','Temps','C1','C2','C3','C4','C5','C6'];
    var rows = results.map(function(r) {
      var d = new Date(r.timestamp);
      var comps = r.competences || {};
      return [
        d.toLocaleDateString('fr-FR') + ' ' + d.toLocaleTimeString('fr-FR',{hour:'2-digit',minute:'2-digit'}),
        r.module || '',
        r.nom || '',
        r.prenom || '',
        r.classe || '',
        r.note20 || '',
        r.score || '',
        r.detail || '',
        r.temps || '',
        comps.C1||'', comps.C2||'', comps.C3||'', comps.C4||'', comps.C5||'', comps.C6||''
      ].map(function(v) { return '"' + String(v).replace(/"/g,'""') + '"'; }).join(';');
    });

    var csv = '\uFEFF' + headers.join(';') + '\n' + rows.join('\n');
    var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = (filename || 'inerweb-resultats') + '.csv';
    a.click();
    URL.revokeObjectURL(url);
  }

  // Exposer l'API
  window.inerwebSend = inerwebSend;
  window.inerwebGetLocal = inerwebGetLocal;
  window.inerwebRetryAll = inerwebRetryAll;
  window.inerwebExportCSV = inerwebExportCSV;

  // Aussi en objet pour ceux qui préfèrent
  window.inerWeb = {
    send: inerwebSend,
    getLocal: inerwebGetLocal,
    retryAll: inerwebRetryAll,
    exportCSV: inerwebExportCSV,
    SHEET_URL: SHEET_URL
  };

})(window);
