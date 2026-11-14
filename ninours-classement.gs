/**
 * Ninours — classement des scores dans Google Sheets.
 *
 * À coller dans : Extensions > Apps Script (depuis ton Google Sheet),
 * puis Déployer > Nouveau déploiement > Application Web,
 *   - Exécuter en tant que : Moi
 *   - Qui a accès : Tout le monde
 * L'adresse obtenue (qui se termine par /exec) est à coller dans le jeu (LEADERBOARD_URL).
 */

const SHEET_NAME = 'Scores';   // nom de l'onglet (créé automatiquement)
const TOP_SIZE = 10;           // nombre de joueurs renvoyés au jeu
const MAX_COINS = 5000;        // au-delà, le score est refusé (anti-triche basique)

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Pseudo', 'Pièces', 'Date']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Le jeu lit le classement : meilleur score de chaque pseudo, trié.
function doGet() {
  const rows = getSheet_().getDataRange().getValues().slice(1);
  const best = {};
  rows.forEach(function (row) {
    const name = String(row[0] || '').trim();
    const coins = Number(row[1]) || 0;
    if (!name) return;
    if (!(name in best) || coins > best[name]) best[name] = coins;
  });
  const top = Object.keys(best)
    .map(function (name) { return { pseudo: name, coins: best[name] }; })
    .sort(function (a, b) { return b.coins - a.coins; })
    .slice(0, TOP_SIZE);
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
    getSheet_().appendRow([name, coins, new Date()]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
