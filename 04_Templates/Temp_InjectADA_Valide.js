// ============================================================================
// TEMPORAIRE — à supprimer après tests
// Injecte une ligne r&K_Adaptabilite avec des réponses *texte* valides
// dans le fichier seed pointé par RESPONSES_SSID_FORCE.
// ============================================================================

function Temp_injectADA_valide() {
  const sp = PropertiesService.getScriptProperties();
  const ssid = sp.getProperty('RESPONSES_SSID_FORCE');
  const sheetName = sp.getProperty('RESPONSES_SHEET_NAME') || 'Feuille 1';
  if (!ssid) throw new Error('FORCE non défini. Exécute d’abord run_pointToTest_Adap().');

  // 1) Feuille de réponses cible
  const ss = SpreadsheetApp.openById(ssid);
  const sh = ss.getSheetByName(sheetName) || ss.getSheets()[0];
  const lc = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];

  // 2) Charger les questions pour récupérer les libellés d’options
  const ids = getSystemIds();
  const bdd = SpreadsheetApp.openById(ids.ID_BDD);
  const qsh = bdd.getSheetByName('Questions_r&K_Adaptabilite_FR');
  if (!qsh) throw new Error('Questions_r&K_Adaptabilite_FR introuvable dans la BDD.');

  const qdata = qsh.getDataRange().getValues();
  const qHeaders = qdata.shift();
  const colID = qHeaders.indexOf('ID');
  const colParams = qHeaders.indexOf('Paramètres (JSON)');
  const qMap = {};
  qdata.forEach(r => {
    const id = r[colID];
    if (!id) return;
    try { qMap[id] = JSON.parse(r[colParams] || '{}'); } catch (e) {}
  });

  // 3) Construire la ligne (texte)
  const row = new Array(lc).fill('');
  headers.forEach((h, i) => {
    const H = String(h || '');
    if (/horodatage/i.test(H)) row[i] = new Date();
    else if (/mail|e-?mail/i.test(H)) row[i] = Session.getActiveUser().getEmail();
    else if (/nom/i.test(H)) row[i] = 'DEV – ADA (valide)';
    else if (/langue/i.test(H)) row[i] = 'Français';
    else if (/^RK\d{2}\s*:/.test(H)) {
      const id = H.split(':')[0].trim();   // ex: "RK01"
      const p = qMap[id];
      if (p && p.options && p.options.length) {
        // on choisit la 1re option par défaut (valide pour QCU_CAT)
        row[i] = p.options[0].libelle;
      }
    }
  });

  sh.appendRow(row);
  Logger.log('✅ Ligne ADA (texte) injectée dans "%s" (row %s).', sh.getName(), sh.getLastRow());
}

