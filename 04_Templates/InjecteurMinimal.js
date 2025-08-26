// Ajoute une ligne de test “V2” (toutes les questions = 3), email = le tien
function injectLigneV2_Median() {
  const cfg = getTestConfiguration();
  const sh  = _getReponsesSheet_(cfg, {}); // pointe vers [CONFIG] automatiquement
  const lc  = sh.getLastColumn();
  const lr  = sh.getLastRow();
  const headers = sh.getRange(1,1,1,lc).getValues()[0];
  const row = new Array(lc).fill('');

  for (let c = 0; c < lc; c++) {
    const h = String(headers[c] || '');
    if (h.indexOf(':') !== -1) {
      // Colonne question "Qxxx: ...": on met 3 par défaut
      row[c] = 3;
    } else if (/mail|e-?mail/i.test(h)) {
      row[c] = Session.getActiveUser().getEmail();
    } else if (/nom/i.test(h)) {
      row[c] = 'Test V2 (injecteur médian)';
    } else if (/horodatage|timestamp/i.test(h)) {
      row[c] = new Date();
    }
  }

  sh.appendRow(row);
  const newRow = lr + 1;
  Logger.log('✅ Ligne V2 de test ajoutée en ligne ' + newRow + ' dans "' + sh.getName() + '"');
}
