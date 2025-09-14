/**
 * =================================================================================
 * == FICHIER : TEMPLATE_T_Data.gs (POUR LA BIBLIOTHÈQUE)
 * == VERSION : 2.1 - Ajout d'une fonction de vérification de version.
 * == RÔLE    : Gère la lecture et la préparation des données en utilisant le contexte du kit.
 * =================================================================================
 */

// ============================================================================
// SECTION - Fonctions de diagnostic
// ============================================================================

/**
 * Retourne un numéro de version codé en dur pour vérifier quelle version
 * du script est actuellement déployée et active.
 */
function getVersion() {
  return "VERSION_FINALE_15_SEPT";
}

var __DBG = true;

function DBG() {
  if (!__DBG) return;
  const parts = [].slice.call(arguments).map(x => (typeof x === 'object' ? JSON.stringify(x) : String(x)));
  Logger.log('[DBG] ' + parts.join(' '));
}

function _spyDumpRow_(sheet, rowIndex) {
  try {
    const lastCol = sheet.getLastColumn();
    if (!lastCol) return null;
    const H = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const V = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const subset = {};
    for (let i = 0; i < Math.min(H.length, 25); i++) subset[H[i]] = V[i];
    DBG('DUMP row', rowIndex, 'subset=', subset);
    return { headers: H, values: V };
  } catch (e) {
    DBG('spyDumpRow ERROR', e.message);
  }
  return null;
}

function _spyFindNomEmail_(reponse) {
  const keys = Object.keys(reponse || {});
  const norm = k => _nettoyerEnTete(k).toLowerCase();
  const allowedName = new Set(['votre_nom_et_prenom', 'nom_et_prenom', 'nom_prenom', 'nomprenom']);
  const allowedEmail = new Set(['votre_adresse_e_mail', 'votre_adresse_email', 'adresse_e_mail', 'email', 'email_repondant', 'email_du_repondant']);
  let nom = '', email = '';
  for (const k of keys) {
    const n = norm(k);
    if (!nom && allowedName.has(n)) nom = reponse[k];
    if (!email && allowedEmail.has(n)) email = reponse[k];
  }
  return { nom, email };
}

// ============================================================================
// SECTION - Fonctions de lecture et de préparation des données
// ============================================================================

function _nettoyerEnTete(enTete) {
  if (!enTete) return "";
  const accents = 'ÀÁÂÃÄÅàáâãäåÒÓÔÕÖØòóôõöøÈÉÊËèéêëÇçÌÍÎÏìíîïÙÚÛÜùúûüÿÑñ';
  const sansAccents = 'AAAAAAaaaaaaOOOOOOooooooEEEEeeeeCcIIIIiiiiUUUUuuuuyNn';
  return enTete.toString().split('').map((char) => {
    const i = accents.indexOf(char);
    return i !== -1 ? sansAccents[i] : char;
  }).join('').replace(/[^a-zA-Z0-9_]/g, '_');
}

function _sheetLooksLikeResponses_(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    if (!lastCol) return false;
    const rawHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
    const norm = h => _nettoyerEnTete(h).toLowerCase();
    const Hn = rawHeaders.map(norm);
    const hasName = Hn.includes('votre_nom_et_prenom') || Hn.includes('nom_et_prenom');
    const hasEmail = Hn.includes('votre_adresse_e_mail') || Hn.includes('votre_adresse_email') || Hn.includes('adresse_e_mail') || Hn.includes('email');
    const hasQuestionId = rawHeaders.some(h => /(^|\s)Q\d+\s*:/.test(h) || /^ENV\s*\d{3}/i.test(h) || /^[A-Z]{2,4}\d{2,3}\s*:/.test(h));
    return (hasName && hasEmail) || hasQuestionId;
  } catch (e) {
    return false;
  }
}

function _pickSheetByNameOrHeuristic_(ss, nameMaybe) {
  if (nameMaybe) {
    const sh = ss.getSheetByName(nameMaybe);
    if (sh) return sh;
  }
  const rx = /^(réponses?\s+au\s+formulaire.*|form\s+responses?.*|responses?)$/i;
  const sheets = ss.getSheets();
  for (const sh of sheets) {
    if (rx.test(sh.getName())) return sh;
  }
  const candidates = sheets.filter(sh => _sheetLooksLikeResponses_(sh));
  return candidates.length > 0 ? candidates[0] : sheets[0];
}

function _getReponsesSheet_(config, kitSpreadsheet) {
  const props = PropertiesService.getScriptProperties();
  const ssidProp = props.getProperty('RESPONSES_SSID');
  let ss = null;
  let used = '';
  if (ssidProp) {
    try {
      ss = SpreadsheetApp.openById(ssidProp);
      used = `ScriptProp(${ssidProp})`;
    } catch (e) {
      DBG('ID de réponse (ScriptProp) invalide:', ssidProp);
    }
  }
  if (!ss && kitSpreadsheet) {
    ss = kitSpreadsheet;
    used = `KitActif(${kitSpreadsheet.getId()})`;
  }
  if (!ss) {
    throw new Error("Impossible d’ouvrir le classeur de réponses. Configurez-le via le menu : Usine à Tests → 'Configurer la feuille de réponses…'.");
  }
  const sheet = _pickSheetByNameOrHeuristic_(ss, config.Nom_Onglet_Reponses);
  if (!sheet || !_sheetLooksLikeResponses_(sheet)) {
    throw new Error(`Classeur ouvert (“${ss.getName()}” via ${used}), mais aucune feuille ne ressemble à une feuille de réponses.`);
  }
  Logger.log(`Source réponses → ${ss.getName()} [${used}] :: onglet "${sheet.getName()}"`);
  DBG('ReponsesSheet -> classeur:', ss.getName(), '| onglet:', sheet.getName(), '| lastRow=', sheet.getLastRow(), '| lastCol=', sheet.getLastColumn());
  return sheet;
}

function _creerObjetReponse(rowIndex, kitSpreadsheet) {
  const config = getTestConfiguration(kitSpreadsheet);
  const sheet = _getReponsesSheet_(config, kitSpreadsheet);
  _spyDumpRow_(sheet, Math.max(2, rowIndex || sheet.getLastRow()));
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (!rowIndex || rowIndex < 2 || rowIndex > lastRow) {
    if (lastRow < 2) {
      throw new Error("Aucune donnée dans la feuille de réponses (seulement l’en-tête).");
    }
    rowIndex = lastRow;
  }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  const reponse = {};
  headers.forEach((header, i) => {
    let cle = header;
    if (header && !String(header).includes(':')) {
      cle = _nettoyerEnTete(header);
    }
    if (cle) reponse[cle] = rowValues[i];
  });
  const spy = _spyFindNomEmail_(reponse);
  reponse.nomRepondant = spy.nom;
  reponse.emailRepondant = spy.email;
  DBG('_creerObjetReponse row=', rowIndex, 'keys=', Object.keys(reponse).slice(0, 12), '| nom=', spy.nom, '| email=', spy.email);
  return reponse;
}

function getDonneesPourRetraitement(rowIndex) {
  try {
    const kitSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const config = getTestConfiguration(kitSpreadsheet);
    const reponse = _creerObjetReponse(rowIndex, kitSpreadsheet);
    const langueOrigine = getOriginalLanguage(reponse);
    return {
      nomRepondant: reponse.nomRepondant,
      emailRepondant: reponse.emailRepondant,
      langueOrigine: langueOrigine,
      repondantActif: config.Repondant_Email_Actif === 'Oui',
      formateurActif: config.Formateur_Email_Actif === 'Oui',
      patronActif: config.Patron_Email_Mode === 'Oui',
      formateurEmail: config.Formateur_Email || '',
      patronEmail: config.Patron_Email || '',
      emailAlias: config.Email_Alias || ''
    };
  } catch (e) {
    Logger.log("ERREUR getDonneesPourRetraitement: " + e.message + "\n" + e.stack);
    throw new Error("Impossible de charger les données: " + e.message);
  }
}
// ... (le reste de vos fonctions de diagnostic reste ici)
function diagnostic_SourceReponses(kitId) {
  try {
    Logger.log('--- DIAGNOSTIC SOURCE RÉPONSES ---');
    if (!kitId) {
      Logger.log("Veuillez fournir un ID de kit pour lancer le diagnostic.");
      return;
    }
    const kitSpreadsheet = SpreadsheetApp.openById(kitId);
    const cfg = getTestConfiguration(kitSpreadsheet);
    const sheet = _getReponsesSheet_(cfg, kitSpreadsheet);
    Logger.log(`✅ Succès : La feuille de réponses a été trouvée et validée.`);
    Logger.log(`   -> Classeur: "${sheet.getParent().getName()}" (ID: ${sheet.getParent().getId()})`);
    Logger.log(`   -> Onglet: "${sheet.getName()}"`);
  } catch (e) {
    Logger.log('❌ ERREUR lors du diagnostic de la source des réponses :');
    Logger.log(e.message);
  }
}

function diagnostic_CompoEmails_v20_1(kitId) {
  try {
    if (!kitId) {
      Logger.log("Veuillez fournir un ID de kit pour lancer le diagnostic.");
      return;
    }
    const kitSpreadsheet = SpreadsheetApp.openById(kitId);
    const cfg = getTestConfiguration(kitSpreadsheet);
    const typeTest = cfg.Type_Test;
    const niveau = (String(cfg.ID_Gabarit_Email_Repondant || '').replace('RESULTATS_', '').trim() || 'N1');
    Logger.log(`--- DIAGNOSTIC COMPOSITION E-MAILS (type=${typeTest}, niveau=${niveau}) ---`);

    assemblerEtEnvoyerEmailUniversel(
      cfg,
      { Votre_adresse_e_mail: 'test@example.com', Votre_nom_et_prenom: 'Testeur' },
      { profilFinal: 'PROFIL_TEST', scoresData: { A: 10, B: 20 }, mapCodeToName: { A: 'Profil A', B: 'Profil B' } },
      'FR',
      { dryRun: true },
      kitSpreadsheet
    );
  } catch (e) {
    Logger.log('ERREUR diagnostic compo: ' + e.message);
  }
}