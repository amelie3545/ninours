/**
 * Ninours — classement des scores dans Google Sheets.
 *
 * Deux onglets (créés automatiquement) :
 *   - « Scores »    : l'historique de toutes les parties (Pseudo, Pièces, Date)
 *   - « Meilleurs » : une ligne par joueur avec son record, toujours trié
 * Le jeu lit seulement « Meilleurs », donc le chargement reste rapide même avec beaucoup de parties.
 *
 * À coller dans : Extensions > Apps Script (depuis ton Google Sheet),
 * puis Déployer > Gérer les déploiements > crayon > Version : Nouvelle version > Déployer.
 * (Premier déploiement : Déployer > Nouveau déploiement > Application Web,
 *  Exécuter en tant que : Moi, Qui a accès : Tout le monde.)
 *
 * Si tu supprimes des lignes dans « Scores » (faux scores…), lance la fonction
 * reconstruireMeilleurs depuis l'éditeur (menu déroulant en haut, puis Exécuter)
 * pour recalculer l'onglet « Meilleurs ». Tu peux aussi supprimer la ligne directement dans « Meilleurs ».
 */

const SCORES_SHEET = 'Scores';     // historique de toutes les parties
const BEST_SHEET = 'Meilleurs';    // record de chaque joueur
const TOP_SIZE = 20;               // joueurs renvoyés au jeu (le jeu peut demander ?n=, max 50)
const MAX_COINS = 5000;            // au-delà, le score est refusé (anti-triche basique)

function getScoresSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SCORES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SCORES_SHEET);
    sh.appendRow(['Pseudo', 'Pièces', 'Date']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getBestSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(BEST_SHEET);
  if (!sh) {
    sh = ss.insertSheet(BEST_SHEET);
    sh.appendRow(['Pseudo', 'Record (pièces)', 'Date du record']);
    sh.setFrozenRows(1);
    fillBestFromScores_(sh);   // première fois : on part de l'historique existant
  }
  return sh;
}

// Recalcule « Meilleurs » à partir de tout l'historique « Scores ».
function fillBestFromScores_(bestSheet) {
  const rows = getScoresSheet_().getDataRange().getValues().slice(1);
  const best = {};
  rows.forEach(function (row) {
    const name = String(row[0] || '').trim();
    const coins = Number(row[1]) || 0;
    if (!name) return;
    if (!(name in best) || coins > best[name].coins) best[name] = { coins: coins, date: row[2] };
  });
  const out = Object.keys(best)
    .map(function (name) { return [name, best[name].coins, best[name].date]; })
    .sort(function (a, b) { return b[1] - a[1]; });
  if (bestSheet.getLastRow() > 1) bestSheet.getRange(2, 1, bestSheet.getLastRow() - 1, 3).clearContent();
  if (out.length) bestSheet.getRange(2, 1, out.length, 3).setValues(out);
}

// À lancer à la main depuis l'éditeur si tu as nettoyé l'onglet « Scores ».
function reconstruireMeilleurs() {
  fillBestFromScores_(getBestSheet_());
}

function sortBest_(sh) {
  const last = sh.getLastRow();
  if (last > 2) sh.getRange(2, 1, last - 1, 3).sort({ column: 2, ascending: false });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Le jeu lit le classement : l'onglet « Meilleurs » est déjà trié.
function doGet(e) {
  const n = Math.min(50, Math.max(1, Number(e && e.parameter && e.parameter.n) || TOP_SIZE));
  const sh = getBestSheet_();
  const last = sh.getLastRow();
  if (last < 2) return json_({ top: [] });
  const rows = sh.getRange(2, 1, Math.min(n, last - 1), 2).getValues();
  const top = rows
    .filter(function (r) { return String(r[0] || '').trim(); })
    .map(function (r) { return { pseudo: String(r[0]).trim(), coins: Number(r[1]) || 0 }; });
  return json_({ top: top });
}

// Le jeu envoie un score en fin de partie : { pseudo, coins }.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const name = String(data.pseudo || '')
      .replace(/[<>]/g, '')
      .replace(/^[=+\-@\s]+/, '')   // évite qu'un pseudo soit lu comme une formule
      .trim()
      .slice(0, 14);
    const coins = Math.floor(Number(data.coins));
    if (!name || !(coins > 0) || coins > MAX_COINS) return json_({ ok: false });

    const now = new Date();
    getScoresSheet_().appendRow([name, coins, now]);   // historique

    // record du joueur
    const best = getBestSheet_();
    const last = best.getLastRow();
    let found = null;
    if (last > 1) {
      found = best.getRange(2, 1, last - 1, 1)
        .createTextFinder(name).matchEntireCell(true).matchCase(true).findNext();
    }
    if (!found) {
      best.appendRow([name, coins, now]);
      sortBest_(best);
    } else if (coins > (Number(best.getRange(found.getRow(), 2).getValue()) || 0)) {
      best.getRange(found.getRow(), 2, 1, 2).setValues([[coins, now]]);
      sortBest_(best);
    }
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
