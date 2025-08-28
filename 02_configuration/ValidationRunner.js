/** ValidationRunner.gs — Runner de validation des en-têtes (CONFIG → BDD → TEMPLATE)
 * Ajoute un menu "Validation" dans le classeur CONFIG pour vérifier les onglets requis
 * et les en-têtes attendues, en se basant STRICTEMENT sur les noms d’en-têtes (jamais d’indices).
 * Rapport affiché en sidebar (HTML).
 *
 * PRÉREQUIS :
 * - Avoir les IDs des 3 classeurs :
 *    - ID_CONFIG : celui du classeur courant (détecté automatiquement)
 *    - ID_BDD, ID_TEMPLATE : soit lus d’un onglet "sys_ID_Fichiers" (si présent),
 *      soit saisis en dur ci-dessous dans FALLBACK_IDS.
 */

/** ================ CONFIGURATION MINIMALE ================ **/
const FALLBACK_IDS = {
  ID_BDD:      '1m2MGBd0nyiAl3qw032B6Nfj7zQL27bRSBexiOPaRZd8', // ← remplace si besoin
  ID_TEMPLATE: '1XwyTt9hcFLd-_IrCYuKY4_E6Dw9aUrls-AGQp65dzDU'  // ← remplace si besoin
};
// Variantes de noms acceptées pour certains onglets (tolérance orthographe/accents)
const SHEET_NAME_VARIANTS = {
  'Paramètres Généraux': ['Paramètres Généraux','Parametres Generaux','Parameters','Paramètres Generaux','Parametres Généraux']
};

/** ================ MENU ================== **/
function addValidationMenu_(ui) {
  ui.createMenu('Validation')
    .addItem('Vérifier les en-têtes (CONFIG, BDD, TEMPLATE)', 'validateAllHeaders')
    .addToUi();
}
/** ================ HELPERS ================== **/
function normalizeHeader_(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // retire accents
    .replace(/\s+/g, ' ') // espaces multiples → un espace
    .trim()
    .toLowerCase();
}

function getHeaderRow_(sheet) {
  if (!sheet) return [];
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1,1,1,lastCol).getValues()[0] || [];
}

function findSheetByVariants_(ss, canonicalName) {
  const variants = SHEET_NAME_VARIANTS[canonicalName] || [canonicalName];
  for (const name of variants) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  return null;
}

function assertHeaders_(sheet, requiredNames, report, context) {
  const headers = getHeaderRow_(sheet);
  const normalized = headers.map(normalizeHeader_);
  const missing = requiredNames
    .map(normalizeHeader_)
    .filter(req => !normalized.includes(req));

  if (missing.length) {
    report.push({
      classeur: context.classeur,
      onglet: sheet ? sheet.getName() : context.ongletAttendu,
      type: 'En-têtes manquantes',
      details: 'Manquantes: ' + missing.join(', '),
      headersTrouves: headers.join(' | ')
    });
  }
}

function getSystemIdsFromConfig_() {
  // Essaie de lire dans un onglet "sys_ID_Fichiers" (2 colonnes : Clé, Valeur)
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('sys_ID_Fichiers');
  const ids = { ID_CONFIG: ss.getId(), ID_BDD: FALLBACK_IDS.ID_BDD, ID_TEMPLATE: FALLBACK_IDS.ID_TEMPLATE };

  if (!sh) return ids;

  const values = sh.getDataRange().getValues();
  const map = {};
  values.forEach(row => {
    const k = String(row[0] || '').trim();
    const v = String(row[1] || '').trim();
    if (k && v) map[k] = v;
  });

  if (map.ID_CONFIG)   ids.ID_CONFIG = map.ID_CONFIG;
  if (map.ID_BDD)      ids.ID_BDD = map.ID_BDD;
  if (map.ID_TEMPLATE) ids.ID_TEMPLATE = map.ID_TEMPLATE;

  return ids;
}

function htmlReport_(rows) {
  const esc = s => String(s||'').replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
  const head = `
    <style>
      body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;padding:12px}
      h2{margin:0 0 10px 0}
      table{border-collapse:collapse;width:100%}
      th,td{border:1px solid #ddd;padding:6px;vertical-align:top}
      th{background:#fafafa;text-align:left}
      tr:nth-child(even){background:#fcfcfc}
      .ok{color:#2e7d32}
      .err{color:#b71c1c}
    </style>`;
  if (!rows.length) {
    return HtmlService.createHtmlOutput(head + `<h2>Validation des en-têtes</h2><p class="ok">Aucune anomalie détectée 👍</p>`)
      .setTitle('Validation en-têtes');
  }
  const rowsHtml = rows.map(r => `
    <tr>
      <td>${esc(r.classeur)}</td>
      <td>${esc(r.onglet)}</td>
      <td class="err">${esc(r.type)}</td>
      <td>${esc(r.details)}</td>
      <td>${esc(r.headersTrouves || '')}</td>
    </tr>`).join('');
  const html = `
    ${head}
    <h2>Validation des en-têtes</h2>
    <table>
      <thead><tr>
        <th>Classeur</th><th>Onglet</th><th>Type</th><th>Détails</th><th>En-têtes trouvées</th>
      </tr></thead>
      <tbody>${rowsHtml}</tbody>
    </table>`;
  return HtmlService.createHtmlOutput(html).setTitle('Validation en-têtes');
}

/** ================ RUNNER PRINCIPAL ================== **/
function validateAllHeaders() {
  const report = [];
  const { ID_CONFIG, ID_BDD, ID_TEMPLATE } = getSystemIdsFromConfig_();

  // --- CONFIG ---
  try {
    const ssCfg = SpreadsheetApp.openById(ID_CONFIG);
    const shParams = findSheetByVariants_(ssCfg, 'Paramètres Généraux');
    if (!shParams) {
      report.push({ classeur:'CONFIG', onglet:'Paramètres Généraux', type:'Onglet manquant', details:'Aucune variante trouvée (Paramètres Généraux/Parametres Generaux/Parameters)' });
    } else {
      assertHeaders_(shParams, [
        'Type_Test','Repondant_Quand','Repondant_Contenu','Patron_Quand','Patron_Contenu','Formateur_Quand','Formateur_Contenu'
      ], report, { classeur:'CONFIG', ongletAttendu:'Paramètres Généraux' });
    }
  } catch (e) {
    report.push({ classeur:'CONFIG', onglet:'*', type:'Erreur', details:String(e) });
  }

  // --- BDD ---
  try {
    const ssBdd = SpreadsheetApp.openById(ID_BDD);
    // sys_Composition_Emails
    const shCompo = ssBdd.getSheetByName('sys_Composition_Emails');
    if (!shCompo) {
      report.push({ classeur:'BDD', onglet:'sys_Composition_Emails', type:'Onglet manquant', details:'Non trouvé' });
    } else {
      assertHeaders_(shCompo, [
        'Type_Test','Code_Langue','Code_Niveau_Email','Code_Profil','Element','Ordre','Contenu / ID_Document'
      ], report, { classeur:'BDD', ongletAttendu:'sys_Composition_Emails' });
    }
    // sys_PiecesJointes
    const shPJ = ssBdd.getSheetByName('sys_PiecesJointes');
    if (!shPJ) {
      report.push({ classeur:'BDD', onglet:'sys_PiecesJointes', type:'Onglet manquant', details:'Non trouvé' });
    } else {
      assertHeaders_(shPJ, [
        'Type_Test','Code_Langue','Code_Niveau_Email','Code_Profil','ID_Document','Nom_Fichier'
      ], report, { classeur:'BDD', ongletAttendu:'sys_PiecesJointes' });
    }
    // Questions_META_FR (optionnel mais fréquent)
    const shMeta = ssBdd.getSheetByName('Questions_META_FR');
    if (shMeta) {
      assertHeaders_(shMeta, [
        'ID_Question','Libelle','Type','Obligatoire','Bloc'
      ], report, { classeur:'BDD', ongletAttendu:'Questions_META_FR' });
    }
  } catch (e) {
    report.push({ classeur:'BDD', onglet:'*', type:'Erreur', details:String(e) });
  }

  // --- TEMPLATE ---
  try {
    const ssTpl = SpreadsheetApp.openById(ID_TEMPLATE);
    // Onglet de config attendu côté TEMPLATE (à adapter si besoin)
    const shTplCfg = ssTpl.getSheetByName('sys_Template_Config');
    if (shTplCfg) {
      assertHeaders_(shTplCfg, [
        'Cle','Valeur'
      ], report, { classeur:'TEMPLATE', ongletAttendu:'sys_Template_Config' });
    } else {
      // pas bloquant : beaucoup de logique template est dans la BDD (compo emails / PJ)
      // On signale juste l'absence d'onglet de config local si on s'y attendait :
      // report.push({ classeur:'TEMPLATE', onglet:'sys_Template_Config', type:'Onglet manquant', details:'Optionnel mais recommandé' });
    }
  } catch (e) {
    report.push({ classeur:'TEMPLATE', onglet:'*', type:'Erreur', details:String(e) });
  }

  // Affiche le rapport
  const out = htmlReport_(report);
  SpreadsheetApp.getUi().showSidebar(out);
}

