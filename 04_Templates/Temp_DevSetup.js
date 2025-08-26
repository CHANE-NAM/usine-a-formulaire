// === FICHIER TEMPORAIRE : Temp_DevSetup.gs ===
// But : créer des feuilles de réponses "DEV" et injecter 1 ligne de test
// Statut : TEMPORAIRE (à supprimer une fois les Forms/feuilles réelles en place)

function dev_seedResponseSheet(typeTest, langue) {
  const sys = getSystemIds();
  const bdd = SpreadsheetApp.openById(sys.ID_BDD);
  const shQ = bdd.getSheetByName(`Questions_${typeTest}_${langue}`);
  if (!shQ) throw new Error(`Feuille introuvable: Questions_${typeTest}_${langue}`);

  const data = shQ.getDataRange().getValues();
  const headers = data.shift();
  const idxID = headers.indexOf('ID');
  const idxTitre = headers.indexOf('Titre');
  if (idxID === -1 || idxTitre === -1) throw new Error("Colonnes 'ID' ou 'Titre' manquantes.");

  // Crée la feuille de réponses [DEV]
  const respSS = SpreadsheetApp.create(`[DEV] ${typeTest} – Réponses`);
  const rs = respSS.getSheets()[0];

  // Entêtes standard + entêtes questions "ID: Titre"
  const std = ['Horodateur','Votre adresse e-mail','Votre nom et prénom','Langue / Language'];
  const qHeaders = data.map(r => `${r[idxID]}: ${r[idxTitre]}`);
  rs.getRange(1,1,1,std.length+qHeaders.length).setValues([std.concat(qHeaders)]);

  // Map Script Property pour le résolveur
  PropertiesService.getScriptProperties()
    .setProperty(`RESPONSES_SSID_${typeTest}`, respSS.getId());

  Logger.log(`Créé et mappé : ${typeTest} → ${respSS.getId()} (onglet "${rs.getName()}")`);
}

function dev_injectMedianRow(typeTest) {
  const ssid = PropertiesService.getScriptProperties().getProperty(`RESPONSES_SSID_${typeTest}`);
  if (!ssid) throw new Error(`RESPONSES_SSID_${typeTest} absent (lance d'abord dev_seedResponseSheet).`);
  const sh = SpreadsheetApp.openById(ssid).getSheets()[0];

  const lc = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,lc).getValues()[0];
  const row = new Array(lc).fill('');

  // Champs standard
  const idxH = headers.indexOf('Horodateur');
  if (idxH !== -1) row[idxH] = new Date();
  const idxM = headers.indexOf('Votre adresse e-mail');
  if (idxM !== -1) row[idxM] = Session.getActiveUser().getEmail();
  const idxN = headers.indexOf('Votre nom et prénom');
  if (idxN !== -1) row[idxN] = 'DEV – Injecté (médian)';
  const idxL = headers.indexOf('Langue / Language');
  if (idxL !== -1) row[idxL] = 'Français';

  // Questions : met "3" par défaut (médian simple)
  for (let c=0;c<lc;c++){
    if (/^[A-Z]{3}\d{3}: /.test(String(headers[c]||''))) row[c] = 3;
  }

  sh.appendRow(row);
  Logger.log(`Ligne injectée dans "${sh.getName()}" (row ${sh.getLastRow()}).`);
}

// === WRAPPERS SANS PARAMÈTRES (TEMP) ===
// Statut : TEMPORAIRE (préfixe Temp) — on supprimera après les tests.

// 1) Création des feuilles de réponses "DEV"
function Temp_seed_ADA_FR()  { dev_seedResponseSheet('r&K_Adaptabilite', 'FR'); }
function Temp_seed_RESI_FR() { dev_seedResponseSheet('r&K_Resilience',   'FR'); }
function Temp_seed_CREA_FR() { dev_seedResponseSheet('r&K_Creativite',   'FR'); }

// 2) Injection d’une ligne médiane (valeur 3 partout)
function Temp_inject_ADA()   { dev_injectMedianRow('r&K_Adaptabilite'); }
function Temp_inject_RESI()  { dev_injectMedianRow('r&K_Resilience'); }
function Temp_inject_CREA()  { dev_injectMedianRow('r&K_Creativite'); }

// ====== TEMP PATCH : injecteur sécurisé vers la feuille de réponses mappée ======
// Statut: TEMP (on supprimera après la mise en préprod)

// Wrapper sans paramètres pour SEED Adaptabilité FR
function Temp_seed_ADA_FR() { dev_seedResponseSheet('r&K_Adaptabilite', 'FR'); }

// Résout la feuille de réponses via Script Properties (clé typée, puis générique)
function _dev_getResponsesSheetFor_(typeTest) {
  const sp = PropertiesService.getScriptProperties();
  const keys = ['RESPONSES_SSID_' + typeTest, 'RESPONSES_SSID']; // compat rétro
  for (var i=0; i<keys.length; i++) {
    var id = sp.getProperty(keys[i]);
    if (id) {
      var ss = SpreadsheetApp.openById(id);
      var sh = ss.getSheetByName('Réponses au formulaire 1') || ss.getSheets()[0];
      return sh;
    }
  }
  throw new Error('Aucun SSID de réponses trouvé pour : ' + typeTest + ' (ni clé typée, ni clé générique).');
}

// Injecte une ligne médiane dans la feuille MAPPÉE (et pas la feuille active)
function Temp_inject_ADA_safe() {
  var sh = _dev_getResponsesSheetFor_('r&K_Adaptabilite');
  var lc = sh.getLastColumn();
  var headers = sh.getRange(1,1,1,lc).getValues()[0];
  var row = new Array(lc).fill('');

  for (var c=0; c<lc; c++) {
    var h = String(headers[c] || '');
    if (h.indexOf(':') !== -1) {
      row[c] = 3; // valeur médiane par défaut
    } else if (/mail|e-?mail/i.test(h)) {
      row[c] = Session.getActiveUser().getEmail();
    } else if (/nom/i.test(h)) {
      row[c] = 'DEV – V2 médian (ADA)';
    } else if (/horodatage|timestamp/i.test(h)) {
      row[c] = new Date();
    }
  }

  sh.appendRow(row);
  Logger.log('✅ Ligne injectée dans "' + sh.getName() + '" du fichier [' + sh.getParent().getName() + '].');
}

