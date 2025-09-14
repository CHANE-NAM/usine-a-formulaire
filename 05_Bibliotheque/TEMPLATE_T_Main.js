/**
 * =================================================================================
 * == FICHIER : TEMPLATE_T_Main.gs (POUR LA BIBLIOTHÈQUE)
 * == RÔLE    : Orchestrateur principal pour le traitement des réponses.
 * == VERSION : 3.2 - Ajout de l'alias 'main' pour compatibilité V4.
 * =================================================================================
 */

// ============================================================================
// SECTION - ALIAS AJOUTÉ POUR COMPATIBILITÉ AVEC LES KITS V4
// ============================================================================

/**
 * ALIAS pour onFormSubmit. C'est le point d'entrée appelé par le déclencheur
 * 'onFormSubmit' des nouveaux kits de test V4.
 * @param {object} e L'objet événement de la soumission du formulaire.
 * @param {string} kitId L'ID du Google Sheet du kit.
 */
function main(e, kitId) {
  // Appelle simplement la fonction onFormSubmit existante dans ce même fichier.
  onFormSubmit(e, kitId);
}

// ============================================================================
// SECTION - Fonctions appelées par le menu "Usine à Tests" via le connecteur
// ============================================================================

/**
 * Ouvre le classeur du kit via son ID et récupère la ligne sélectionnée.
 * @param {string} kitId L'ID du Google Sheet du kit.
 * @returns {number} Le numéro de la ligne.
 */
function _getRowFromSelectionOrAsk_(kitId) {
  const ss = SpreadsheetApp.openById(kitId);
  const sh = ss.getActiveSheet();
  const r = sh.getActiveRange();
  if (r && r.getRow() >= 2) return r.getRow();

  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Retraitement', 'Numéro de ligne (≥ 2) ?', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) throw new Error('Annulé par l\'utilisateur.');

  const n = parseInt(resp.getResponseText(), 10);
  if (!n || n < 2) throw new Error('Numéro de ligne invalide.');
  return n;
}

/** Dry-run sur la dernière ligne (reçoit l'ID du kit). */
function ui_DryRunDerniereLigne(kitId) {
  try {
    const kitSpreadsheet = SpreadsheetApp.openById(kitId);
    const cfg = getTestConfiguration(kitSpreadsheet);
    const sh = _getReponsesSheet_(cfg, kitSpreadsheet);
    const lr = sh.getLastRow();
    if (lr < 2) throw new Error('Feuille vide.');

    const reponse = _creerObjetReponse(lr, kitSpreadsheet);
    const langue = getOriginalLanguage(reponse) || 'FR';
    const niveau = (String(cfg.ID_Gabarit_Email_Repondant || '').replace('RESULTATS_', '').trim() || 'N1');

    retraitementTestSansEnvoi(lr, kitSpreadsheet, {
      langue: langue,
      niveau: niveau,
      destinataires: { test: Session.getActiveUser().getEmail() }
    });
    SpreadsheetApp.getUi().alert('Dry-run lancé sur la dernière ligne (' + lr + '). Voir les journaux de la bibliothèque.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erreur Dry-run (dernière ligne) : ' + e.message);
  }
}

/** Dry-run sur la ligne sélectionnée (reçoit l'ID du kit). */
function ui_DryRunLigneSelection(kitId) {
  try {
    const row = _getRowFromSelectionOrAsk_(kitId);
    const kitSpreadsheet = SpreadsheetApp.openById(kitId);
    const cfg = getTestConfiguration(kitSpreadsheet);
    const reponse = _creerObjetReponse(row, kitSpreadsheet);
    const langue = getOriginalLanguage(reponse) || 'FR';
    const niveau = (String(cfg.ID_Gabarit_Email_Repondant || '').replace('RESULTATS_', '').trim() || 'N1');

    retraitementTestSansEnvoi(row, kitSpreadsheet, {
      langue: langue,
      niveau: niveau,
      destinataires: { test: Session.getActiveUser().getEmail() }
    });
    SpreadsheetApp.getUi().alert('Dry-run lancé sur la ligne ' + row + '. Voir les journaux de la bibliothèque.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erreur Dry-run (ligne sélectionnée) : ' + e.message);
  }
}

/** Envoi réel sur la ligne sélectionnée (reçoit l'ID du kit). */
function ui_EnvoiReelLigneSelection(kitId) {
  try {
    const row = _getRowFromSelectionOrAsk_(kitId);
    const kitSpreadsheet = SpreadsheetApp.openById(kitId);
    traiterLigne(row, kitSpreadsheet, { isRetraitement: true, dryRun: false, ignoreDeveloppeurEmail: false });
    SpreadsheetApp.getUi().alert('Envoi RÉEL lancé sur la ligne ' + row + '. Voir les journaux de la bibliothèque.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erreur Envoi réel : ' + e.message);
  }
}

/** Persiste l'ID du classeur de réponses (ne change pas, car indépendant du kit). */
function ui_ConfigResponsesSheet() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperty('RESPONSES_SSID') || '';
  const msg = 'Colle ici l’ID du *classeur de réponses* lié au Google Form.';
  const resp = ui.prompt('Configurer la feuille de réponses', msg + (current ? '\n\nActuel : ' + current : ''), ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const val = (resp.getResponseText() || '').trim();
  if (!val) { ui.alert('ID vide — aucune modification.'); return; }

  props.setProperty('RESPONSES_SSID', val);
  ui.alert('✅ Feuille de réponses configurée.\nID = ' + val);
}


// ============================================================================
// SECTION - Points d'entrée et Orchestration
// ============================================================================

/**
 * Point d'entrée principal (reçoit l'ID du kit depuis le connecteur).
 */
function onFormSubmit(e, kitId) {
  try {
    Logger.log("Nouvelle réponse reçue, traitement pour le kit ID: " + kitId);
    const kitSpreadsheet = SpreadsheetApp.openById(kitId);
    traiterLigne(e.range.getRow(), kitSpreadsheet);
  } catch (err) {
    Logger.log("ERREUR FATALE dans onFormSubmit: " + err.toString() + "\n" + err.stack);
  }
}

/**
 * Orchestre le traitement complet pour une ligne, en utilisant le bon classeur.
 */
function traiterLigne(rowIndex, kitSpreadsheet, optionsSurcharge = {}) {
  try {
    const config = getTestConfiguration(kitSpreadsheet);
    Logger.log('CONFIG CHARGÉE: ' + JSON.stringify(config)); 
    const reponse = _creerObjetReponse(rowIndex, kitSpreadsheet);
    const langueOrigine = getOriginalLanguage(reponse);
    const langueCible = optionsSurcharge.langue || langueOrigine || 'FR';
    const resultats = calculerResultats(reponse, langueCible, config, langueOrigine);

    if (typeof creerGraphiqueRadar === 'function' && resultats.Score_Echec) {
      const axesData = {
        Echec: { score: resultats.Score_Echec },
        Changement: { score: resultats.Score_Changement },
        Ressources: { score: resultats.Score_Ressources },
        Crise: { score: resultats.Score_Crise },
        Objectifs: { score: resultats.Score_Objectifs }
      };
      const chartImage = creerGraphiqueRadar(axesData);
      if (chartImage) {
        resultats.Graphique_Radar_Blob = chartImage;
      }
    }

    assemblerEtEnvoyerEmailUniversel(config, reponse, resultats, langueCible, optionsSurcharge, kitSpreadsheet);
  } catch (e) {
    Logger.log("!!!! ERREUR FATALE dans traiterLigne !!!! pour le kit " + kitSpreadsheet.getName());
    Logger.log("Message : " + e.message);
    Logger.log("Stack Trace : " + e.stack);
    try {
      const sheet = _getReponsesSheet_(getTestConfiguration(kitSpreadsheet), kitSpreadsheet);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const statutCol = headers.indexOf('Statut_Traitement');
      if (statutCol !== -1) {
        sheet.getRange(rowIndex, statutCol + 1).setValue('ERREUR');
      }
    } catch (e2) {
      Logger.log("Impossible de mettre à jour le statut d'erreur : " + e2.message);
    }
    throw e;
  }
}

/**
 * Lance le retraitement (reçoit les options incluant l'ID du kit).
 */
function lancerRetraitementDepuisUI(options) {
  try {
    const kitSpreadsheet = SpreadsheetApp.openById(options.kitId);
    traiterLigne(options.rowIndex, kitSpreadsheet, {
      isRetraitement: true,
      dryRun: false,
      ignoreDeveloppeurEmail: true,
      langue: options.langue,
      niveau: options.niveau,
      alias: options.alias,
      destinataires: options.destinataires || {},
      overrideRecipients: true
    });
    Logger.log(`Retraitement manuel lancé pour la ligne ${options.rowIndex} avec succès.`);
    return `Retraitement pour la ligne ${options.rowIndex} terminé avec succès !`;
  } catch (e) {
    Logger.log(`ERREUR lors du retraitement depuis UI pour la ligne ${options.rowIndex}: ${e.toString()}`);
    throw new Error(`Échec du retraitement : ${e.message}`);
  }
}

/**
 * Exécute un test de retraitement sans envoyer d'email.
 */
function retraitementTestSansEnvoi(rowIndex, kitSpreadsheet, options) {
  try {
    traiterLigne(rowIndex, kitSpreadsheet, {
      isRetraitement: true,
      dryRun: true,
      ignoreDeveloppeurEmail: true,
      langue: options.langue,
      niveau: options.niveau,
      destinataires: options.destinataires,
      overrideRecipients: true
    });
  } catch (e) {
    Logger.log(`ERREUR lors du dry-run pour la ligne ${rowIndex}: ${e.toString()}`);
    throw new Error(e.message);
  }
}