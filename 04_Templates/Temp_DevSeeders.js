// ============================================================================
// FICHIER : Temp_DevSeeders.gs  (TEMPORAIRE / DEBUG)
// VERSION : 1.2
// RÔLE    : Créer/relier des feuilles de réponses DEV et poser les mappings.
// ============================================================================

/** Lie (ou crée) une feuille de réponses DEV et enregistre le mapping. */
function dev_linkSeedSheet(typeTest, langue) {
  const sp = PropertiesService.getScriptProperties();
  const sheetName = sp.getProperty('RESPONSES_SHEET_NAME') || 'Réponses au formulaire 1';

  // 1) mapping existant ?
  let map = {};
  try { map = JSON.parse(sp.getProperty('RESPONSES_SSID_BY_TEST') || '{}'); } catch (_) {}
  if (map[typeTest]) {
    Logger.log('Trouvé : [DEV] %s – Réponses → %s', typeTest, map[typeTest]);
    dev_setResponseMapping(typeTest, map[typeTest], sheetName);
    return;
  }

  // 2) fichier DEV déjà présent dans Drive ?
  const found = dev_findDevResponseFile(typeTest);
  if (found) {
    Logger.log('Trouvé via Drive : [DEV] %s – Réponses → %s', typeTest, found.getId());
    dev_setResponseMapping(typeTest, found.getId(), sheetName);
    return;
  }

  // 3) sinon : on crée et on “seed” les entêtes
  const ss = dev_seedResponseSheet(typeTest, langue);
  const ssid = ss ? ss.getId() : '';
  const firstName = (ss && ss.getSheets()[0]) ? ss.getSheets()[0].getName() : sheetName;

  Logger.log('Créé via dev_seedResponseSheet : [DEV] %s – Réponses → %s', typeTest, ssid || '(undefined)');
  dev_setResponseMapping(typeTest, ssid, firstName);
}

/** Crée un Spreadsheet [DEV] …, pose l’onglet et les en-têtes à partir de Questions_* */
function dev_seedResponseSheet(typeTest, langue) {
  try {
    const name = `[DEV] ${typeTest} – Réponses`;
    const existing = dev_findDevResponseFile(typeTest);
    let ss = existing ? SpreadsheetApp.openById(existing.getId()) : SpreadsheetApp.create(name);

    // Onglet cible (on garde le 1er onglet existant)
    const sh = ss.getSheets()[0];

    // Construire les en-têtes depuis BDD: Questions_{type}_{langue}
    let headers = ['Horodateur', '[Type Inconnu: TEXTE_EMAIL] Votre adresse e-mail', 'Votre_nom_et_prenom', 'Langue___Language'];
    try {
      const systemIds = getSystemIds();
      const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
      const qSheet = bdd.getSheetByName(`Questions_${typeTest}_${langue}`);
      if (!qSheet) throw new Error(`Feuille introuvable: Questions_${typeTest}_${langue}`);
      const data = qSheet.getDataRange().getValues();
      const head = data.shift();
      const colID = head.indexOf('ID');
      const colTitre = head.indexOf('Titre');
      if (colID === -1 || colTitre === -1) throw new Error("Colonnes 'ID' ou 'Titre' manquantes.");
      data.forEach(row => {
        const id = row[colID];
        const titre = row[colTitre];
        if (id && titre) headers.push(`${id}: ${titre}`);
      });
    } catch (e) {
      Logger.log('⚠️ Impossible de lire Questions_%s_%s : %s', typeTest, langue, e.message);
    }

    // Écrit les en-têtes si la 1ère ligne est vide
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    Logger.log('Créé et mappé : %s → %s (onglet "%s")', typeTest, ss.getId(), sh.getName());
    return ss;
  } catch (e) {
    Logger.log('dev_seedResponseSheet ERREUR: ' + e.message);
    return null;
  }
}

/** Enregistre le mapping dans Script Properties (+ FORCE immédiat) */
function dev_setResponseMapping(typeTest, ssid, sheetName) {
  if (!ssid) {
    Logger.log('⚠️ dev_setResponseMapping: SSID vide pour ' + typeTest + ' → mapping ignoré.');
    return;
  }
  const sp = PropertiesService.getScriptProperties();
  let map = {};
  try { map = JSON.parse(sp.getProperty('RESPONSES_SSID_BY_TEST') || '{}'); } catch (_) {}
  map[typeTest] = ssid;
  sp.setProperty('RESPONSES_SSID_BY_TEST', JSON.stringify(map));
  if (sheetName) sp.setProperty('RESPONSES_SHEET_NAME', String(sheetName));
  sp.setProperty('RESPONSES_SSID_FORCE', ssid); // pour que le dry-run prenne effet tout de suite
  Logger.log('Mappé : %s → %s (onglet="%s")', typeTest, ssid, sheetName || '(par défaut)');
}

/** Cherche un fichier Drive “[DEV] {type} – Réponses” */
function dev_findDevResponseFile(typeTest) {
  const name = `[DEV] ${typeTest} – Réponses`;
  const q = 'mimeType = "application/vnd.google-apps.spreadsheet" and trashed = false and title contains "' +
            name.replace(/"/g, '\\"') + '"';
  const it = DriveApp.searchFiles(q);
  return it.hasNext() ? it.next() : null;
}
