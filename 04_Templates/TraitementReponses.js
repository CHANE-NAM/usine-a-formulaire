/**
 * =================================================================================
 * == FICHIER : TraitementReponses.gs
 * == VERSION : 23.1 - Version de débogage pour analyser le filtrage des e-mails.
 * ==           (Précédent: 23.0 - Intégration graphique et Champ_Profil)
 * =================================================================================
 */

// ====== DEBUG / ESPIONS ======
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
    try { return { ss: SpreadsheetApp.openById(id), used: `${tag}(${id})` }; } catch(_){ DBG('tryOpenById FAIL', tag, id); return null; }
  }
  let pick = (options.reponsesSpreadsheetId && tryOpenById(options.reponsesSpreadsheetId, 'ById(options)')) || (ssidProp && tryOpenById(ssidProp, 'ScriptProp')) ||
    ( (config?.ID_Sheet_Reponses || config?.ID_SHEET_REPONSES || config?.ID_REPONSES_SPREADSHEET) && tryOpenById(config.ID_Sheet_Reponses || config.ID_SHEET_REPONSES || config.ID_REPONSES_SPREADSHEET, 'CONFIG') ) ||
    ( (sys?.ID_Sheet_Reponses || sys?.ID_SHEET_REPONSES || sys?.ID_REPONSES || sys?.ID_REPONSES_SHEET) && tryOpenById(sys.ID_Sheet_Reponses || sys.ID_SHEET_REPONSES || sys.ID_REPONSES || sys.ID_REPONSES_SHEET, 'SYS') );
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
  // ... (contenu de la fonction identique à la version précédente)
}

function normalizeAndDedupeCompositionEmailsRows_(rows, idx) {
  // ... (contenu de la fonction identique)
}
function _enrichirDonneesPourEmail_(reponse, resultats) {
  // ... (contenu de la fonction identique)
}
function onFormSubmit(e) {
  // ... (contenu de la fonction identique)
}
function _envoyerEmailDeConfirmation(config, reponse, langueCible) {
  // ... (contenu de la fonction identique)
}

function traiterLigne(rowIndex, optionsSurcharge = {}) {
  // ... (contenu de la fonction identique)
}

// ==================== DÉBUT DE LA MODIFICATION (VERSION DEBOGAGE) ====================
function assemblerEtEnvoyerEmailUniversel(config, reponse, resultats, langueCible, optionsSurcharge = {}){
  const typeTest = (config.Type_Test || '').toString().trim();
  let codeNiveauEmail = (config.ID_Gabarit_Email_Repondant || '').toString().replace('RESULTATS_', '').trim();
  if (optionsSurcharge && optionsSurcharge.niveau && optionsSurcharge.niveau !== '') codeNiveauEmail = optionsSurcharge.niveau;
  const profilFinal = (resultats.profilFinal || '').toString().trim();
  
  Logger.log("--- DÉBUT DÉBOGAGE FILTRE E-MAIL ---");
  Logger.log(`Paramètres de filtrage : typeTest="${typeTest}", langueCible="${langueCible}", codeNiveauEmail="${codeNiveauEmail}", profilFinal="${profilFinal}"`);
  const systemIds = getSystemIds();
  const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
  const compoSheet = bdd.getSheetByName("sys_Composition_Emails");
  const compoData = compoSheet.getDataRange().getValues();
  const compoHeaders = compoData.shift();
  const idx = { typeTest: compoHeaders.indexOf('Type_Test'), langue: compoHeaders.indexOf('Code_Langue'), niveau: compoHeaders.indexOf('Code_Niveau_Email'), profil: compoHeaders.indexOf('Code_Profil'), element: compoHeaders.indexOf('Element'), ordre: compoHeaders.indexOf('Ordre'), contenu: compoHeaders.indexOf('Contenu / ID_Document') };
  const compoRows = normalizeAndDedupeCompositionEmailsRows_(compoData, idx);
  
  Logger.log(`Analyse de ${compoRows.length} lignes de sys_Composition_Emails...`);
  let briquesDeContenu = compoRows.filter((row, rowIndex) => {
    const typeLigne = (row[idx.typeTest] || '').toString().trim();
    const langLigne = (row[idx.langue] || '').toString().trim();
    const levelValue = (row[idx.niveau] || '').toString();
    const profilLigne = (row[idx.profil] || '').toString().trim();
    
    const typeMatch = (typeLigne === typeTest || typeLigne === '');
    const langMatch = (langLigne === langueCible || langLigne === '');
    const levelList = levelValue.split(',').map(s => s.trim()).filter(Boolean);
    const levelMatch = levelList.length > 0 ? levelList.includes(codeNiveauEmail) : levelValue.includes(codeNiveauEmail);
    const profileMatch = (profilLigne === profilFinal || profilLigne === '');
    
    const decision = typeMatch && langMatch && levelMatch && profileMatch;

    // Log uniquement pour les lignes potentiellement pertinentes
    if (typeLigne === typeTest) {
       Logger.log(`Ligne ${rowIndex + 2}: [type="${typeLigne}", lang="${langLigne}", niveau="${levelValue}", profil="${profilLigne}"] -> typeMatch=${typeMatch}, langMatch=${langMatch}, levelMatch=${levelMatch}, profileMatch=${profileMatch} => ${decision ? "INCLUS" : "REJETÉ"}`);
    }
    
    return decision;
  });
  Logger.log(`--- FIN DÉBOGAGE --- Total de briques trouvées : ${briquesDeContenu.length}`);
  if (briquesDeContenu.length === 0) {
    Logger.log("ERREUR ❌: Aucune brique de contenu n'a été trouvée. L'e-mail sera vide. Vérifiez les conditions de filtrage ci-dessus.");
  }
  
  // Le reste de la fonction est identique...
  briquesDeContenu.sort((a, b) => (Number(a[idx.ordre]) || 0) - (Number(b[idx.ordre]) || 0));
  const donneesPourEmail = _enrichirDonneesPourEmail_(reponse, resultats);
  let contenuInfoCopie = null;
  const indexInfoCopie = briquesDeContenu.findIndex(b => (b[idx.element] || '').toString().trim() === 'Info_Copie');
  if (indexInfoCopie > -1) {
    contenuInfoCopie = briquesDeContenu[indexInfoCopie][idx.contenu];
    briquesDeContenu.splice(indexInfoCopie, 1);
  }
  let sujet = `Résultats de votre test ${typeTest}`;
  let corpsHtml = "";
  const piecesJointesIds = new Set();
  const testsAvecScoreEntier = ['ANCRES', 'COULEURS', 'MBTI'];
  for (const brique of briquesDeContenu) {
    const elementType = (brique[idx.element] || '').toString().trim();
    const contenu = brique[idx.contenu];
    switch (elementType) {
      case 'Sujet_Email': sujet = contenu; break;
      case 'Introduction': case 'Corps_Texte': corpsHtml += (contenu || "") + "<br>"; break;
      case 'Champ_Profil': if (contenu && donneesPourEmail[contenu]) { corpsHtml += donneesPourEmail[contenu] + "<br>"; } break;
      case 'Document': if (contenu && String(contenu).trim()) piecesJointesIds.add(String(contenu).trim()); break;
      case 'Ligne_Score':
        const scoresAAfficher = resultats.scoresData;
        if (scoresAAfficher) {
          Object.entries(scoresAAfficher)
          .sort((a, b) => b[1] - a[1])
          .forEach(([code, score]) => {
            let scoreArrondi;
            if (testsAvecScoreEntier.some(test => typeTest.toUpperCase().includes(test))) {
                scoreArrondi = (typeof score === 'number') ? Math.round(score) : score;
      
            } else {
                scoreArrondi = (typeof score === 'number') ? score.toFixed(1) : score;
            }
            const nomProfil = resultats.mapCodeToName[code] || code;
            const totalPossible = resultats.scoresMaxPossible ? (resultats.scoresMaxPossible[code] || 'N/A') : 'N/A';
            let ligneScore = (contenu || `- {{nom_profil}} : {{score}} points`)
              .replace(/{{nom_profil}}/g, nomProfil).replace(/{{score}}/g, scoreArrondi).replace(/{{total_possible}}/g, totalPossible);
            corpsHtml += ligneScore + "<br>";
          });
        }
        break;
    }
  }
  for (const key in donneesPourEmail) {
    const placeholder = `{{${key}}}`;
    const valeur = donneesPourEmail[key] || '';
    const regex = new RegExp(placeholder.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
    sujet = sujet.replace(regex, valeur);
    corpsHtml = corpsHtml.replace(regex, valeur);
    if (contenuInfoCopie) contenuInfoCopie = contenuInfoCopie.replace(regex, valeur);
  }
  const variablesFusion = { ...donneesPourEmail };
  const piecesJointes = [];
  for (const contenuDoc of Array.from(piecesJointesIds)) {
    let candidate = contenuDoc;
    if (candidate.startsWith("{{") && candidate.endsWith("}}")) {
      const cle = candidate.slice(2, -2);
      candidate = variablesFusion[cle] || "";
    }
    if (/^[a-zA-Z0-9_-]{20,}$/.test(candidate)) {
      try {
        const nomRapport = (resultats.titreProfil || resultats.profilFinal || config.Type_Test || "Rapport");
        const pdf = genererPdfDepuisModele(candidate, variablesFusion, nomRapport);
        if (pdf) { piecesJointes.push(pdf); }
      } catch(e) {
        Logger.log("Fusion Doc->PDF échouée pour " + candidate + " : " + e.message);
        try { piecesJointes.push(DriveApp.getFileById(candidate).getBlob()); } catch(_) {}
      }
    } else {
      Logger.log("Ignoré (Document) : valeur non reconnue " + candidate);
    }
  }
  const T = loadTraductions(langueCible);
  const emailRepondantPrincipal = reponse.Votre_adresse_e_mail || reponse.Votre_adresse_email || reponse.Adresse_e_mail || reponse.emailRepondant;
  const override = optionsSurcharge.overrideRecipients === true;
  const ignoreDev = optionsSurcharge.ignoreDeveloppeurEmail === true;
  const dryRun = optionsSurcharge.dryRun === true;
  const destS = optionsSurcharge.destinataires || {};
  const adressesUniques = new Set();
  if (override) {
    if (destS.repondant && emailRepondantPrincipal) adressesUniques.add(emailRepondantPrincipal);
    if (destS.formateur && destS.formateurEmail) adressesUniques.add(destS.formateurEmail);
    if (destS.patron && destS.patronEmail) adressesUniques.add(destS.patronEmail);
    if (destS.test && destS.test.trim() !== '') { destS.test.split(',').map(e => e.trim()).forEach(email => adressesUniques.add(email)); }
  } else {
    if (Object.keys(destS).length > 0) {
      if (destS.repondant && emailRepondantPrincipal) adressesUniques.add(emailRepondantPrincipal);
      if (destS.formateur && destS.formateurEmail) adressesUniques.add(destS.formateurEmail);
      if (destS.patron && destS.patronEmail) adressesUniques.add(destS.patronEmail);
      if (destS.test && destS.test.trim() !== '') { destS.test.split(',').map(e => e.trim()).forEach(email => adressesUniques.add(email)); }
    } else {
      if (config.Repondant_Email_Actif === 'Oui' && emailRepondantPrincipal) adressesUniques.add(emailRepondantPrincipal);
      if (config.Patron_Email_Mode === 'Oui' && config.Patron_Email) adressesUniques.add(config.Patron_Email);
      if (config.Formateur_Email_Actif === 'Oui' && config.Formateur_Email) adressesUniques.add(config.Formateur_Email);
    }
    if (config.Developpeur_Email && !ignoreDev) adressesUniques.add(config.Developpeur_Email);
  }
  if (dryRun) {
    Logger.log('— DRY-RUN — AUCUN EMAIL ENVOYÉ —');
    Logger.log('Destinataires simulés : ' + Array.from(adressesUniques).join(', '));
    Logger.log('Sujet (après remplacements) : ' + sujet);
    Logger.log('Corps (aperçu 400c) : ' + (corpsHtml || '').slice(0, 400));
    Logger.log('Pièces jointes (nb) : ' + piecesJointes.length + (piecesJointesIds.size ? ' | Modèles: ' + Array.from(piecesJointesIds).join(', ') : ''));
    return;
  }
  adressesUniques.forEach(adresse => {
    try {
      let sujetFinal = sujet;
      let corpsHtmlFinal = corpsHtml;
      if (adresse.toLowerCase() !== (emailRepondantPrincipal || "").toLowerCase()) {
        sujetFinal = (T.PREFIXE_COPIE_EMAIL || "Copie : ") + sujet;
        if (contenuInfoCopie) corpsHtmlFinal = contenuInfoCopie + corpsHtml;
      }
      const mailOptions = { to: adresse, subject: sujetFinal, htmlBody: corpsHtmlFinal, attachments: piecesJointes };
      const aliasExpediteur = optionsSurcharge.alias || config.Email_Alias;
      if (aliasExpediteur && aliasExpediteur.trim() !== '') mailOptions.from = aliasExpediteur;
      GmailApp.sendEmail(mailOptions.to, mailOptions.subject, "", mailOptions);
      Logger.log(`E-mail de RÉSULTATS [${langueCible}] envoyé à ${adresse}.`);
    } catch (e) {
      Logger.log(`Echec de l'envoi des résultats à ${adresse}. Erreur: ${e.message}`);
    }
  });
}
// ==================== FIN DE LA MODIFICATION (VERSION DEBOGAGE) ====================

/**
 * Récupère les données initiales d'une ligne pour pré-remplir la sidebar de retraitement.
 * Appelé depuis RetraitementUI.html.
 * @param {number} rowIndex Le numéro de la ligne à retraiter.
 * @returns {object} Un objet avec les informations nécessaires pour l'interface.
 */
function getDonneesPourRetraitement(rowIndex) {
  try {
    const config = getTestConfiguration();
    const reponse = _creerObjetReponse(rowIndex);
    const langueOrigine = getOriginalLanguage(reponse);

    // Utilise la fonction robuste pour trouver le nom et l'e-mail
    const id = _spyFindNomEmail_(reponse);

    return {
      nomRepondant: id.nom,
      emailRepondant: id.email,
      langueOrigine: langueOrigine,
      // Pré-coche les cases en fonction de la configuration
      repondantActif: config.Repondant_Email_Actif === 'Oui',
      formateurActif: config.Formateur_Email_Actif === 'Oui',
      patronActif: config.Patron_Email_Mode === 'Oui',
      // Pré-remplit les champs e-mail depuis la configuration
      formateurEmail: config.Formateur_Email || '',
      patronEmail: config.Patron_Email || '',
      emailAlias: config.Email_Alias || ''
    };
  } catch (e) {
    Logger.log("ERREUR getDonneesPourRetraitement: " + e.message);
    throw new Error("Impossible de charger les données: " + e.message);
  }
}

/**
 * Lance le retraitement complet depuis l'interface utilisateur (sidebar).
 * Appelé par le bouton "Lancer le Retraitement".
 * @param {object} options Les options sélectionnées dans la sidebar.
 * @returns {string} Un message de succès.
 */
function lancerRetraitementDepuisUI(options) {
  try {
    const destinatairesSurcharge = options.destinataires || {};
    // On force l'override pour n'envoyer qu'aux destinataires cochés
    destinatairesSurcharge.overrideRecipients = true; 

    traiterLigne(options.rowIndex, {
      isRetraitement: true,
      dryRun: false, // C'est un envoi réel
      ignoreDeveloppeurEmail: true, // On n'inclut pas le dev, sauf si c'est une adresse de test
      langue: options.langue,
      niveau: options.niveau,
      alias: options.alias,
      destinataires: destinatairesSurcharge
    });

    Logger.log(`Retraitement manuel lancé pour la ligne ${options.rowIndex} avec succès.`);
    return `Retraitement pour la ligne ${options.rowIndex} terminé avec succès !`;
  } catch (e) {
    Logger.log(`ERREUR lors du retraitement depuis UI pour la ligne ${options.rowIndex}: ${e.toString()}`);
    throw new Error(`Échec du retraitement : ${e.message}`);
  }
}

/**
 * Simule un traitement complet (calculs + assemblage e-mail) sans aucun envoi.
 * Les logs affichent le contenu de l'e-mail qui aurait été envoyé.
 * @param {number} rowIndex Le numéro de la ligne.
 * @param {object} options Options de langue, niveau et destinataires pour le test.
 */
function retraitementTestSansEnvoi(rowIndex, options) {
  try {
    traiterLigne(rowIndex, {
      isRetraitement: true,
      dryRun: true, // Active le mode simulation
      ignoreDeveloppeurEmail: true,
      langue: options.langue,
      niveau: options.niveau,
      destinataires: options.destinataires,
      overrideRecipients: true
    });
  } catch(e) {
    Logger.log(`ERREUR lors du dry-run pour la ligne ${rowIndex}: ${e.toString()}`);
    throw new Error(e.message);
  }
}

/**
 * Outil de diagnostic qui vérifie la configuration de la source des réponses.
 */
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

/**
 * Outil de diagnostic pour la composition des e-mails (version 20.1+).
 */
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