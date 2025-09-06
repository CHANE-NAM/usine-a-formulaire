/**
 * T_Data.gs
 * @fileoverview Gère la lecture et la préparation des données depuis les feuilles de calcul.
 * Inclut les fonctions de récupération des réponses, les utilitaires de débogage (spy)
 * et les diagnostics.
 * @version 1.0
 */

// ============================================================================
// SECTION - Fonctions de débogage / espions
// ============================================================================
var __DBG = true; // ← mets false pour couper les logs

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
  } catch (e) { DBG('spyDumpRow ERROR', e.message);
  }
  return null;
}
function _spyFindNomEmail_(reponse) {
  const keys = Object.keys(reponse || {});
  const norm = k => _nettoyerEnTete(k).toLowerCase();
  const allowedName = new Set(['votre_nom_et_prenom','nom_et_prenom','nom_prenom','nomprenom']);
  const allowedEmail = new Set(['votre_adresse_e_mail','votre_adresse_email','adresse_e_mail','email','email_repondant','email_du_repondant']);
  let nom = '', email = '';
  for (const k of keys) {
    const n = norm(k);
    if (!nom && allowedName.has(n))  nom = reponse[k];
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
    const hasName  = Hn.includes('votre_nom_et_prenom') || Hn.includes('nom_et_prenom');
    const hasEmail = Hn.includes('votre_adresse_e_mail') || Hn.includes('votre_adresse_email') || Hn.includes('adresse_e_mail') || Hn.includes('email');
    const hasQuestionId = rawHeaders.some(h => /(^|\s)Q\d+\s*:/.test(h) || /^ENV\s*\d{3}/i.test(h) || /^[A-Z]{2,4}\d{2,3}\s*:/.test(h));
    const ok = (hasName && hasEmail) || hasQuestionId;
    if (!ok) { DBG('sheetLooksLikeResponses=FALSE name=', sheet.getName(), 'headersSample=', rawHeaders.slice(0, 15));
    }
    return ok;
  } catch (e) { return false;
  }
}

function _pickSheetByNameOrHeuristic_(ss, nameMaybe) {
  if (nameMaybe) {
    const sh = ss.getSheetByName(nameMaybe);
    if (sh) return sh;
  }
  const rx = /^(réponses?\s+au\s+formulaire.*|form\s+responses?.*|responses?)$/i;
  const sheets = ss.getSheets();
  for (const sh of sheets) { if (rx.test(sh.getName())) return sh;
  }
  return sheets[0];
}

function _getReponsesSheet_(config, options) {
  options = options || {};
  const sys = (typeof getSystemIds === 'function') ? getSystemIds() : {};
  const props = PropertiesService.getScriptProperties();
  const ssidProp = props.getProperty('RESPONSES_SSID');
  let ss = null, used = '';
  function tryOpenById(id, tag) {
    if (!id) return null;
    try { return { ss: SpreadsheetApp.openById(id), used: `${tag}(${id})` }; } catch(_){ DBG('tryOpenById FAIL', tag, id); return null;
    }
  }
  let pick = (options.reponsesSpreadsheetId && tryOpenById(options.reponsesSpreadsheetId, 'ById(options)')) || (ssidProp && tryOpenById(ssidProp, 'ScriptProp')) ||
    ( (config?.ID_Sheet_Reponses || config?.ID_SHEET_REPONSES || config?.ID_REPONSES_SPREADSHEET) && tryOpenById(config.ID_Sheet_Reponses || config.ID_SHEET_REPONSES || config.ID_REPONSES_SPREADSHEET, 'CONFIG') ) ||
    ( (sys?.ID_Sheet_Reponses || sys?.ID_SHEET_REPONSES || sys?.ID_REPONSES || sys?.ID_REPONSES_SHEET) && tryOpenById(sys.ID_Sheet_Reponses || sys.ID_SHEET_REPONSES || sys.ID_REPONSES || sys.ID_REPONSES_SHEET, 'SYS') );
  if (pick) { ss = pick.ss; used = pick.used; }
  if (!ss) { try { ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) used = 'ActiveSpreadsheet'; } catch (_) {} }
  if (!ss) throw new Error("Impossible d’ouvrir le classeur de réponses. Configure-le via le menu “Configurer la feuille de réponses…” (RESPONSES_SSID).");
  let sheet = _pickSheetByNameOrHeuristic_(ss, options.reponsesSheetName);
  if (!sheet || !_sheetLooksLikeResponses_(sheet)) {
    const candidates = ss.getSheets().filter(sh => _sheetLooksLikeResponses_(sh));
    if (candidates.length) { sheet = candidates[0]; DBG('Heuristic sheet rejected → picked candidate', sheet.getName());
    }
  }
  if (!sheet || !_sheetLooksLikeResponses_(sheet)) { throw new Error( "Classeur ouvert (“" + ss.getName() + "” via " + used + "), mais aucune feuille ne ressemble à une feuille de réponses de test.\n" + "→ Renseigne l’ID du classeur de réponses (Google Sheet lié au Form) via le menu : Usine à Tests → « Configurer la feuille de réponses… »." );
  }
  Logger.log(`Source réponses → ${ss.getName()} [${used}] :: onglet "${sheet.getName()}"`);
  DBG('ReponsesSheet -> classeur:', ss.getName(), '| onglet:', sheet.getName(), '| lastRow=', sheet.getLastRow(), '| lastCol=', sheet.getLastColumn());
  return sheet;
}

function _creerObjetReponse(rowIndex, options) {
  const config = (typeof getTestConfiguration === 'function') ? getTestConfiguration() : {};
  const sheet = _getReponsesSheet_(config, options);
  _spyDumpRow_(sheet, Math.max(2, rowIndex || sheet.getLastRow()));
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (!rowIndex || rowIndex < 2 || rowIndex > lastRow) { if (lastRow < 2) { throw new Error("Aucune donnée dans la feuille de réponses (seulement l’en-tête).");
    } rowIndex = lastRow; }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  const reponse = {};
  headers.forEach((header, i) => { let cle = header; if (header && !String(header).includes(':')) cle = _nettoyerEnTete(header); if (cle) reponse[cle] = rowValues[i]; });
  if (reponse.Votre_adresse_e_mail && !reponse.Votre_adresse_email) reponse.Votre_adresse_email = reponse.Votre_adresse_e_mail;
  if (reponse.Votre_adresse_email && !reponse.Votre_adresse_e_mail) reponse.Votre_adresse_e_mail = reponse.Votre_adresse_email;
  if (reponse.Votre_nom_et_prenom && !reponse.Nom_et_prenom) reponse.Nom_et_prenom = reponse.Votre_nom_et_prenom;
  const spy = _spyFindNomEmail_(reponse);
  DBG('_creerObjetReponse row=', rowIndex, 'keys=', Object.keys(reponse).slice(0, 12), '| nom=', spy.nom, '| email=', spy.email);
  return reponse;
}

function getDonneesPourRetraitement(rowIndex) {
  try {
    const config = getTestConfiguration();
    const reponse = _creerObjetReponse(rowIndex);
    const langueOrigine = getOriginalLanguage(reponse);
    const id = _spyFindNomEmail_(reponse);
    return {
      nomRepondant: id.nom,
      emailRepondant: id.email,
      langueOrigine: langueOrigine,
      repondantActif: config.Repondant_Email_Actif === 'Oui',
      formateurActif: config.Formateur_Email_Actif === 'Oui',
      patronActif: config.Patron_Email_Mode === 'Oui',
      formateurEmail: config.Formateur_Email || '',
      patronEmail: config.Patron_Email || '',
      emailAlias: config.Email_Alias || ''
    };
  } catch (e) {
    Logger.log("ERREUR getDonneesPourRetraitement: " + e.message);
    throw new Error("Impossible de charger les données: " + e.message);
  }
}


// ============================================================================
// SECTION - Fonctions de diagnostic
// ============================================================================

function diagnostic_SourceReponses() {
  try {
    const cfg = getTestConfiguration();
    Logger.log('--- DIAGNOSTIC SOURCE RÉPONSES ---');
    const sheet = _getReponsesSheet_(cfg, {});
    Logger.log(`✅ Succès : La feuille de réponses a été trouvée et validée.`);
    Logger.log(`   -> Classeur: "${sheet.getParent().getName()}" (ID: ${sheet.getParent().getId()})`);
    Logger.log(`   -> Onglet: "${sheet.getName()}"`);
  } catch (e) {
    Logger.log('❌ ERREUR lors du diagnostic de la source des réponses :');
    Logger.log(e.message);
  }
}

function diagnostic_CompoEmails_v20_1() {
  try {
    const cfg = getTestConfiguration();
    const typeTest = cfg.Type_Test;
    const niveau = (String(cfg.ID_Gabarit_Email_Repondant||'').replace('RESULTATS_','').trim() || 'N1');
    Logger.log(`--- DIAGNOSTIC COMPOSITION E-MAILS (type=${typeTest}, niveau=${niveau}) ---`);
    assemblerEtEnvoyerEmailUniversel(
      cfg,
      { Votre_adresse_e_mail: 'test@example.com', Votre_nom_et_prenom: 'Testeur' },
      { profilFinal: 'PROFIL_TEST', scoresData: { A:10, B:20 }, mapCodeToName: { A:'Profil A', B:'Profil B' } },
      'FR',
      { dryRun: true }
    );
  } catch(e){
    Logger.log('ERREUR diagnostic compo: ' + e.message);
  }
}