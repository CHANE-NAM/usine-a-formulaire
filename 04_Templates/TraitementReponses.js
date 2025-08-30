// =================================================================================
// == FICHIER : TraitementReponses.gs
// == VERSION : 20.7 - L'e-mail de confirmation n'est plus envoyé lors d'un retraitement
// ==           ou d'un envoi de résultats immédiat.
// =================================================================================

// ====== DEBUG / ESPIONS ======
var __DBG = true; // ← mets false pour couper les logs

function DBG() {
  if (!__DBG) return;
  const parts = [].slice.call(arguments).map(x => (typeof x === 'object' ? JSON.stringify(x) : String(x)));
  Logger.log('[DBG] ' + parts.join(' '));
}

// ... (le reste du fichier TraitementReponses.js reste identique jusqu'à la fonction traiterLigne)
// ... (pour la lisibilité, je ne colle ici que la fonction modifiée)

function _spyDumpRow_(sheet, rowIndex) {
  try {
    const lastCol = sheet.getLastColumn();
    if (!lastCol) return null;
    const H = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const V = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const subset = {};
    for (let i = 0; i < Math.min(H.length, 25); i++) subset[H[i]] = V[i]; // 25 1res colonnes
    DBG('DUMP row', rowIndex, 'subset=', subset);
    return { headers: H, values: V };
  } catch (e) { DBG('spyDumpRow ERROR', e.message); }
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
    if (!ok) { DBG('sheetLooksLikeResponses=FALSE name=', sheet.getName(), 'headersSample=', rawHeaders.slice(0, 15)); }
    return ok;
  } catch (e) { return false; }
}
function _pickSheetByNameOrHeuristic_(ss, nameMaybe) {
  if (nameMaybe) {
    const sh = ss.getSheetByName(nameMaybe);
    if (sh) return sh;
  }
  const rx = /^(réponses?\s+au\s+formulaire.*|form\s+responses?.*|responses?)$/i;
  const sheets = ss.getSheets();
  for (const sh of sheets) { if (rx.test(sh.getName())) return sh; }
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
    try {
      const ssp = SpreadsheetApp.openById(id);
      DBG('tryOpenById OK', tag, id);
      return { ss: ssp, used: `${tag}(${id})` };
    } catch(_){ DBG('tryOpenById FAIL', tag, id); return null; }
  }
  let pick = (options.reponsesSpreadsheetId && tryOpenById(options.reponsesSpreadsheetId, 'ById(options)')) || (ssidProp && tryOpenById(ssidProp, 'ScriptProp')) || ( (config?.ID_Sheet_Reponses || config?.ID_SHEET_REPONSES || config?.ID_REPONSES_SPREADSHEET) && tryOpenById(config.ID_Sheet_Reponses || config.ID_SHEET_REPONSES || config.ID_REPONSES_SPREADSHEET, 'CONFIG') ) || ( (sys?.ID_Sheet_Reponses || sys?.ID_SHEET_REPONSES || sys?.ID_REPONSES || sys?.ID_REPONSES_SHEET) && tryOpenById(sys.ID_Sheet_Reponses || sys.ID_SHEET_REPONSES || sys.ID_REPONSES || sys.ID_REPONSES_SHEET, 'SYS') );
  if (pick) { ss = pick.ss; used = pick.used; }
  if (!ss) { try { ss = SpreadsheetApp.getActiveSpreadsheet(); if (ss) used = 'ActiveSpreadsheet'; } catch (_) {} }
  if (!ss) throw new Error("Impossible d’ouvrir le classeur de réponses. Configure-le via le menu “Configurer la feuille de réponses…” (RESPONSES_SSID).");
  let sheet = _pickSheetByNameOrHeuristic_(ss, options.reponsesSheetName);
  if (!sheet || !_sheetLooksLikeResponses_(sheet)) {
    const candidates = ss.getSheets().filter(sh => _sheetLooksLikeResponses_(sh));
    if (candidates.length) { sheet = candidates[0]; DBG('Heuristic sheet rejected → picked candidate', sheet.getName()); }
  }
  if (!sheet || !_sheetLooksLikeResponses_(sheet)) { throw new Error( "Classeur ouvert (“" + ss.getName() + "” via " + used + "), mais aucune feuille ne ressemble à une feuille de réponses de test.\n" + "→ Renseigne l’ID du classeur de réponses (Google Sheet lié au Form) via le menu : Usine à Tests → « Configurer la feuille de réponses… »." ); }
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
  if (!rowIndex || rowIndex < 2 || rowIndex > lastRow) { if (lastRow < 2) { throw new Error("Aucune donnée dans la feuille de réponses (seulement l’en-tête)."); } rowIndex = lastRow; }
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
function genererPdfDepuisModele(templateId, variables, nomFichier) {
  if (!templateId) throw new Error("ID du modèle manquant.");
  const templateFile = DriveApp.getFileById(templateId);
  const tempCopy = templateFile.makeCopy((nomFichier || ("Rapport_" + new Date().toISOString().slice(0,10))) + " (temp)");
  const doc = DocumentApp.openById(tempCopy.getId());
  const body = doc.getBody();
  for (const key in variables) {
    const placeholder = "{{" + key + "}}";
    body.replaceText(placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), String(variables[key] ?? ""));
  }
  doc.saveAndClose();
  const pdfBlob = tempCopy.getAs(MimeType.PDF);
  tempCopy.setTrashed(true);
  return pdfBlob;
}
function normalizeAndDedupeCompositionEmailsRows_(rows, idx) {
  const seen = new Set();
  return (rows || []).map(r => { r[idx.element] = (r[idx.element] || '').toString().trim(); return r; }).filter(r => { const key = [(r[idx.typeTest] || '').toString().trim(), (r[idx.langue] || '').toString().trim(), (r[idx.niveau] || '').toString().trim(), (r[idx.profil] || '').toString().trim(), (r[idx.element] || '').toString().trim(), (r[idx.ordre] || '').toString().trim()].join('|'); if (seen.has(key)) return false; seen.add(key); return true; });
}
function _enrichirDonneesPourEmail_(reponse, resultats) {
  const R = reponse || {};
  const donnees = { ...R, ...(resultats || {}) };
  const email = R.Votre_adresse_e_mail || R.Votre_adresse_email || R.Adresse_e_mail || R.emailRepondant || '';
  if (email) { if (!donnees.Votre_adresse_e_mail) donnees.Votre_adresse_e_mail = email; if (!donnees.Votre_adresse_email) donnees.Votre_adresse_email = email; }
  if (R.Votre_nom_et_prenom && !donnees.Votre_nom_et_prenom) { donnees.Votre_nom_et_prenom = R.Votre_nom_et_prenom; } else if (R.Nom_et_prenom && !donnees.Votre_nom_et_prenom) { donnees.Votre_nom_et_prenom = R.Nom_et_prenom; }
  if (donnees.titreProfil && !donnees.Titre_Profil) donnees.Titre_Profil = donnees.titreProfil;
  if (donnees.descriptionProfil && !donnees.Description_Profil) donnees.Description_Profil = donnees.descriptionProfil;
  return donnees;
}
function onFormSubmit(e) {
  try {
    traiterLigne(e.range.getRow(), {});
  } catch (err) {
    Logger.log(`Erreur critique onFormSubmit: ${err}\n${err.stack}`);
  }
}
function _envoyerEmailDeConfirmation(config, reponse, langueCible) {
  try {
    let idGabaritConfirmation = config[`ID_Gabarit_Email_Confirmation_${langueCible}`];
    if (!idGabaritConfirmation || String(idGabaritConfirmation).trim() === '') {
      idGabaritConfirmation = getSystemIds()[`ID_GABARIT_CONFIRMATION_${langueCible}`];
      Logger.log(`Utilisation du gabarit de confirmation PAR DÉFAUT pour ${langueCible}.`);
    } else {
      Logger.log(`Utilisation du gabarit de confirmation SPÉCIFIQUE pour ${langueCible}.`);
    }
    const emailRepondant = reponse.Votre_adresse_e_mail || reponse.Votre_adresse_email || reponse.Adresse_e_mail || reponse.emailRepondant;
    if (!idGabaritConfirmation || !emailRepondant) return;
    const doc = DocumentApp.openById(idGabaritConfirmation);
    let sujet = doc.getName();
    const url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + idGabaritConfirmation + "&exportFormat=html";
    const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } });
    let corpsHtml = response.getContentText();
    const donneesPourEmail = _enrichirDonneesPourEmail_(reponse, null);
    for (const key in donneesPourEmail) {
      const regex = new RegExp(`{{${key}}}`, 'g');
      sujet = sujet.replace(regex, donneesPourEmail[key] || '');
      corpsHtml = corpsHtml.replace(regex, donneesPourEmail[key] || '');
    }
    const mailOptions = { to: emailRepondant, subject: sujet, htmlBody: corpsHtml };
    if (config.Email_Alias && config.Email_Alias.trim() !== '') mailOptions.from = config.Email_Alias;
    GmailApp.sendEmail(mailOptions.to, mailOptions.subject, "", mailOptions);
    Logger.log(`E-mail de confirmation [${langueCible}] envoyé à ${emailRepondant}.`);
  } catch (e) {
    Logger.log(`ERREUR e-mail de confirmation : ${e}\n${e.stack}`);
  }
}

function traiterLigne(rowIndex, optionsSurcharge = {}) {
  try {
    const config = getTestConfiguration();
    const reponse = _creerObjetReponse(rowIndex, optionsSurcharge);
    const langueOrigine = getOriginalLanguage(reponse);
    const langueCible = optionsSurcharge.langue || langueOrigine;

    // --- MODIFICATION V20.7 START ---
    // On n'envoie l'email de confirmation QUE si l'envoi des résultats est différé
    // ET qu'il ne s'agit pas d'un retraitement manuel.
    if (!optionsSurcharge.isRetraitement && config.Repondant_Quand !== 'Immediat') {
      _envoyerEmailDeConfirmation(config, reponse, langueCible);
    }
    // --- MODIFICATION V20.7 END ---

    const resultats = calculerResultats(reponse, langueCible, config, langueOrigine);

    if (optionsSurcharge.isRetraitement || config.Repondant_Quand === 'Immediat') {
      if (config.Moteur_Calcul === 'Universel') {
        Logger.log("Moteur UNIVERSEL → envoi immédiat.");
        assemblerEtEnvoyerEmailUniversel(config, reponse, resultats, langueCible, optionsSurcharge);
      }
    } else {
      Logger.log(`Envoi différé (“${config.Repondant_Quand}”) → programmation.`);
      programmerEnvoiResultats(rowIndex, langueCible, config.Repondant_Quand);
    }
  } catch (err) {
    Logger.log("ERREUR FATALE traiterLigne: " + err + "\n" + err.stack);
  }
}

// ... (le reste du fichier reste identique)
function assemblerEtEnvoyerEmailUniversel(config, reponse, resultats, langueCible, optionsSurcharge = {}){
  const typeTest = (config.Type_Test || '').toString().trim();
  let codeNiveauEmail = (config.ID_Gabarit_Email_Repondant || '').toString().replace('RESULTATS_', '').trim();
  if (optionsSurcharge && optionsSurcharge.niveau && optionsSurcharge.niveau !== '') codeNiveauEmail = optionsSurcharge.niveau;
  const profilFinal = (resultats.profilFinal || '').toString().trim();
  const systemIds = getSystemIds();
  const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
  const compoSheet = bdd.getSheetByName("sys_Composition_Emails");
  const compoData = compoSheet.getDataRange().getValues();
  const compoHeaders = compoData.shift();
  const idx = { typeTest: compoHeaders.indexOf('Type_Test'), langue: compoHeaders.indexOf('Code_Langue'), niveau: compoHeaders.indexOf('Code_Niveau_Email'), profil: compoHeaders.indexOf('Code_Profil'), element: compoHeaders.indexOf('Element'), ordre: compoHeaders.indexOf('Ordre'), contenu: compoHeaders.indexOf('Contenu / ID_Document') };
  const compoRows = normalizeAndDedupeCompositionEmailsRows_(compoData, idx);
  let briquesDeContenu = compoRows.filter(row => { const typeLigne = (row[idx.typeTest] || '').toString().trim(); const typeMatch = (typeLigne === typeTest || typeLigne === ''); const langMatch = ((row[idx.langue] || '').toString().trim() === (langueCible || '').toString().trim()); const levelValue = (row[idx.niveau] || '').toString(); const levelList = levelValue.split(',').map(s => s.trim()).filter(Boolean); const levelMatch = levelList.length > 0 ? levelList.includes(codeNiveauEmail) : levelValue.includes(codeNiveauEmail); const profilLigne = (row[idx.profil] || '').toString().trim(); const profileMatch= (profilLigne === profilFinal || profilLigne === ''); return typeMatch && langMatch && levelMatch && profileMatch; });
  briquesDeContenu.sort((a, b) => (Number(a[idx.ordre]) || 0) - (Number(b[idx.ordre]) || 0));
  let contenuInfoCopie = null;
  const indexInfoCopie = briquesDeContenu.findIndex(b => (b[idx.element] || '').toString().trim() === 'Info_Copie');
  if (indexInfoCopie > -1) { contenuInfoCopie = briquesDeContenu[indexInfoCopie][idx.contenu]; briquesDeContenu.splice(indexInfoCopie, 1); }
  let sujet = `Résultats de votre test ${typeTest}`;
  let corpsHtml = "";
  const piecesJointesIds = new Set();
  for (const brique of briquesDeContenu) {
    const elementType = (brique[idx.element] || '').toString().trim();
    const contenu = brique[idx.contenu];
    switch (elementType) {
      case 'Sujet_Email': sujet = contenu; break;
      case 'Introduction': case 'Corps_Texte': corpsHtml += (contenu || "") + "<br>"; break;
      case 'Document': if (contenu && String(contenu).trim()) piecesJointesIds.add(String(contenu).trim()); break;
      case 'Ligne_Score': Object.entries(resultats.scoresData).sort((a, b) => b[1] - a[1]).forEach(([code, score]) => { let ligneScore = (contenu || "").replace(/{{nom_profil}}/g, resultats.mapCodeToName[code] || code).replace(/{{score}}/g, score); corpsHtml += ligneScore + "<br>"; }); break;
    }
  }
  const donneesPourEmail = _enrichirDonneesPourEmail_(reponse, resultats);
  for (const key in donneesPourEmail) { const placeholder = `{{${key}}}`; const valeur = donneesPourEmail[key] || ''; const regex = new RegExp(placeholder.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g'); sujet = sujet.replace(regex, valeur); corpsHtml = corpsHtml.replace(regex, valeur); if (contenuInfoCopie) contenuInfoCopie = contenuInfoCopie.replace(regex, valeur); }
  const variablesFusion = { ...donneesPourEmail };
  const piecesJointes = [];
  for (const contenuDoc of Array.from(piecesJointesIds)) { let candidate = contenuDoc; if (candidate.startsWith("{{") && candidate.endsWith("}}")) { const cle = candidate.slice(2, -2); candidate = variablesFusion[cle] || ""; } if (/^[a-zA-Z0-9_-]{20,}$/.test(candidate)) { try { const nomRapport = (resultats.titreProfil || resultats.profilFinal || config.Type_Test || "Rapport"); const pdf = genererPdfDepuisModele(candidate, variablesFusion, nomRapport); piecesJointes.push(pdf); } catch(e) { Logger.log("Fusion Doc->PDF échouée pour " + candidate + " : " + e.message); try { piecesJointes.push(DriveApp.getFileById(candidate).getBlob()); } catch(_) {} } } else { Logger.log("Ignoré (Document) : valeur non reconnue " + candidate); } }
  const T = loadTraductions(langueCible);
  const emailRepondantPrincipal = reponse.Votre_adresse_e_mail || reponse.Votre_adresse_email || reponse.Adresse_e_mail || reponse.emailRepondant;
  const override = optionsSurcharge.overrideRecipients === true;
  const ignoreDev = optionsSurcharge.ignoreDeveloppeurEmail === true;
  const dryRun = optionsSurcharge.dryRun === true;
  const destS = optionsSurcharge.destinataires || {};
  const adressesUniques = new Set();
  if (override) { if (destS.repondant && emailRepondantPrincipal) adressesUniques.add(emailRepondantPrincipal); if (destS.formateur && destS.formateurEmail) adressesUniques.add(destS.formateurEmail); if (destS.patron && destS.patronEmail) adressesUniques.add(destS.patronEmail); if (destS.test && destS.test.trim() !== '') { destS.test.split(',').map(e => e.trim()).forEach(email => adressesUniques.add(email)); } } else { if (Object.keys(destS).length > 0) { if (destS.repondant && emailRepondantPrincipal) adressesUniques.add(emailRepondantPrincipal); if (destS.formateur && destS.formateurEmail) adressesUniques.add(destS.formateurEmail); if (destS.patron && destS.patronEmail) adressesUniques.add(destS.patronEmail); if (destS.test && destS.test.trim() !== '') { destS.test.split(',').map(e => e.trim()).forEach(email => adressesUniques.add(email)); } } else { if (config.Repondant_Email_Actif === 'Oui' && emailRepondantPrincipal) adressesUniques.add(emailRepondantPrincipal); if (config.Patron_Email_Mode === 'Oui' && config.Patron_Email) adressesUniques.add(config.Patron_Email); if (config.Formateur_Email_Actif === 'Oui' && config.Formateur_Email) adressesUniques.add(config.Formateur_Email); } if (config.Developpeur_Email && !ignoreDev) adressesUniques.add(config.Developpeur_Email); }
  if (dryRun) { Logger.log('— DRY-RUN — AUCUN EMAIL ENVOYÉ —'); Logger.log('Destinataires simulés : ' + Array.from(adressesUniques).join(', ')); Logger.log('Sujet (après remplacements) : ' + sujet); Logger.log('Corps (aperçu 400c) : ' + (corpsHtml || '').slice(0, 400)); Logger.log('Pièces jointes (nb) : ' + piecesJointes.length + (piecesJointesIds.size ? ' | Modèles: ' + Array.from(piecesJointesIds).join(', ') : '')); return; }
  adressesUniques.forEach(adresse => { try { let sujetFinal = sujet; let corpsHtmlFinal = corpsHtml; if (adresse.toLowerCase() !== (emailRepondantPrincipal || "").toLowerCase()) { sujetFinal = (T.PREFIXE_COPIE_EMAIL || "Copie : ") + sujet; if (contenuInfoCopie) corpsHtmlFinal = contenuInfoCopie + corpsHtml; } const mailOptions = { to: adresse, subject: sujetFinal, htmlBody: corpsHtmlFinal, attachments: piecesJointes }; const aliasExpediteur = optionsSurcharge.alias || config.Email_Alias; if (aliasExpediteur && aliasExpediteur.trim() !== '') mailOptions.from = aliasExpediteur; GmailApp.sendEmail(mailOptions.to, mailOptions.subject, "", mailOptions); Logger.log(`E-mail de RÉSULTATS [${langueCible}] envoyé à ${adresse}.`); } catch (e) { Logger.log(`Echec de l'envoi des résultats à ${adresse}. Erreur: ${e.message}`); } });
}
function getDonneesPourRetraitement(rowIndex) {
  try {
    const config = getTestConfiguration();
    const reponse = _creerObjetReponse(rowIndex, {});
    return { nomRepondant: reponse.Votre_nom_et_prenom || reponse.Nom_et_prenom || '', emailRepondant: reponse.Votre_adresse_e_mail || reponse.Votre_adresse_email || reponse.Adresse_e_mail || '', langueOrigine: getOriginalLanguage(reponse), repondantActif: config.Repondant_Email_Actif === 'Oui', formateurActif: config.Formateur_Email_Actif === 'Oui', patronActif: config.Patron_Email_Mode === 'Oui', emailAlias: config.Email_Alias || '' };
  } catch (e) {
    Logger.log(`ERREUR getDonneesPourRetraitement(${rowIndex}): ${e}`);
    throw new Error("Impossible de récupérer les données pour la ligne " + rowIndex + ". " + e.message);
  }
}
function lancerRetraitementDepuisUI(options) {
  if (!options || !options.rowIndex) throw new Error("Options de retraitement invalides.");
  options.isRetraitement = true;
  traiterLigne(options.rowIndex, options);
  return "Retraitement pour la ligne " + options.rowIndex + " lancé avec succès !";
}
function retraitementTestSansEnvoi(rowIndex, options) {
  options = options || {};
  options.isRetraitement = true;
  options.dryRun = true;
  options.overrideRecipients = true;
  options.ignoreDeveloppeurEmail = true;
  if (!rowIndex) { const config = (typeof getTestConfiguration === 'function') ? getTestConfiguration() : {}; const sh = _getReponsesSheet_(config, options); const lr = sh.getLastRow(); if (lr < 2) throw new Error("Aucune donnée dans la feuille de réponses (seulement l’en-tête)."); rowIndex = lr; }
  traiterLigne(rowIndex, options);
}
function diagnostic_SourceReponses() {
  const sh = _getReponsesSheet_((typeof getTestConfiguration === 'function') ? getTestConfiguration() : {}, {});
  Logger.log(`Diagnostic → classeur: "${sh.getParent().getName()}" | onglet: "${sh.getName()}" | lignes: ${sh.getLastRow()} | colonnes: ${sh.getLastColumn()}`);
}
function diagnostic_CompoEmails_v20_1() {
  const systemIds = getSystemIds();
  const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
  const sh = bdd.getSheetByName('sys_Composition_Emails');
  const data = sh.getDataRange().getValues();
  const headers = data.shift();
  const idx = { typeTest: headers.indexOf('Type_Test'), langue: headers.indexOf('Code_Langue'), niveau: headers.indexOf('Code_Niveau_Email'), profil: headers.indexOf('Code_Profil'), element: headers.indexOf('Element'), ordre: headers.indexOf('Ordre'), contenu: headers.indexOf('Contenu / ID_Document') };
  const before = data.length;
  const trailingSpaces = data.filter(r => /\s$/.test(String(r[idx.element] || ''))).length;
  const afterRows = normalizeAndDedupeCompositionEmailsRows_(data, idx);
  const after = afterRows.length;
  Logger.log(`v20.1 ► sys_Composition_Emails : ${before} → ${after} (doublons retirés = ${before - after})`);
  Logger.log(`v20.1 ► 'Element' avec espace final détectés avant normalisation : ${trailingSpaces}`);
}
