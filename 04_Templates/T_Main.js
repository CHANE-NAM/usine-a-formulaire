/**
 * @fileoverview Orchestrateur principal pour le traitement des réponses.
 * Contient les points d'entrée (onFormSubmit), la logique centrale (traiterLigne)
 * et les fonctions liées à l'interface utilisateur (menus de retraitement).
 * @version 1.0
 */

// ============================================================================
// SECTION - Fonctions appelées par le menu "Usine à Tests"
// ============================================================================

/** Récupère une ligne depuis la sélection, ou demande à l'utilisateur. */
function _getRowFromSelectionOrAsk_() {
  const sh = SpreadsheetApp.getActiveSheet();
  const r = sh.getActiveRange();
  if (r && r.getRow() >= 2) return r.getRow();

  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Retraitement', 'Numéro de ligne (≥ 2) ?', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) throw new Error('Annulé');

  const n = parseInt(resp.getResponseText(), 10);
  if (!n || n < 2) throw new Error('Numéro de ligne invalide.');
  return n;
}

/** Dry-run sur la dernière ligne de la feuille de réponses (aucun e-mail envoyé). */
function ui_DryRunDerniereLigne() {
  try {
    if (typeof getTestConfiguration !== 'function' || typeof _getReponsesSheet_ !== 'function') {
      SpreadsheetApp.getUi().alert('⚠️ Fonctions manquantes (getTestConfiguration/_getReponsesSheet_). Vérifie que le projet contient TraitementReponses.gs v20.4+');
      return;
    }
    const cfg = getTestConfiguration();
    const sh  = _getReponsesSheet_(cfg, {});
    const lr  = sh.getLastRow();
    if (lr < 2) throw new Error('Feuille vide (seulement l’en-tête).');
    const langue = (typeof getOriginalLanguage === 'function' && typeof _creerObjetReponse === 'function')
      ? (getOriginalLanguage(_creerObjetReponse(lr, {})) || 'FR')
      : 'FR';
    const niveau = (String(cfg.ID_Gabarit_Email_Repondant || '').replace('RESULTATS_', '').trim() || 'N1');
    if (typeof retraitementTestSansEnvoi !== 'function') {
      SpreadsheetApp.getUi().alert('⚠️ Fonction manquante: retraitementTestSansEnvoi(). Vérifie TraitementReponses.gs v20.4+');
      return;
    }

    retraitementTestSansEnvoi(lr, {
      langue: langue,
      niveau: niveau,
      destinataires: { test: Session.getActiveUser().getEmail() }
    });
    SpreadsheetApp.getUi().alert('Dry-run lancé sur la dernière ligne (' + lr + '). Voir Journaux.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erreur Dry-run (dernière ligne) : ' + e.message);
  }
}

/** Dry-run sur la ligne sélectionnée (aucun e-mail envoyé). */
function ui_DryRunLigneSelection() {
  try {
    const row   = _getRowFromSelectionOrAsk_();
    const cfg   = (typeof getTestConfiguration === 'function') ? getTestConfiguration() : {};
    const langue = (typeof getOriginalLanguage === 'function' && typeof _creerObjetReponse === 'function')
      ? (getOriginalLanguage(_creerObjetReponse(row, {})) || 'FR')
      : 'FR';
    const niveau = (String(cfg.ID_Gabarit_Email_Repondant || '').replace('RESULTATS_', '').trim() || 'N1');
    if (typeof retraitementTestSansEnvoi !== 'function') {
      SpreadsheetApp.getUi().alert('⚠️ Fonction manquante: retraitementTestSansEnvoi(). Vérifie TraitementReponses.gs v20.4+');
      return;
    }

    retraitementTestSansEnvoi(row, {
      langue: langue,
      niveau: niveau,
      destinataires: { test: Session.getActiveUser().getEmail() }
    });
    SpreadsheetApp.getUi().alert('Dry-run lancé sur la ligne ' + row + '. Voir Journaux.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erreur Dry-run (ligne sélectionnée) : ' + e.message);
  }
}

/** Envoi réel sur la ligne sélectionnée (envoie les e-mails selon CONFIG). */
function ui_EnvoiReelLigneSelection() {
  try {
    const row = _getRowFromSelectionOrAsk_();
    if (typeof traiterLigne !== 'function') {
      SpreadsheetApp.getUi().alert('⚠️ Fonction manquante: traiterLigne(). Vérifie TraitementReponses.gs v20.4+');
      return;
    }
    // Envoi réel (pas de dryRun, destinataires selon CONFIG)
    traiterLigne(row, { isRetraitement: true, dryRun: false, ignoreDeveloppeurEmail: false });
    SpreadsheetApp.getUi().alert('Envoi RÉEL lancé sur la ligne ' + row + '. Voir Journaux.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erreur Envoi réel : ' + e.message);
  }
}

/** Persiste l'ID du vrai classeur de réponses (lié au Google Form). */
function ui_ConfigResponsesSheet() {
  const ui   = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperty('RESPONSES_SSID') || '';
  const msg = 'Colle ici l’ID du *classeur de réponses* lié au Google Form (celui avec les colonnes "Qxxx: ...").\n' +
              'Astuce : Formulaire → onglet "Réponses" → icône Google Sheets (verte) → ouvre le classeur → copie l’ID dans l’URL.';
  const resp = ui.prompt('Configurer la feuille de réponses', msg + (current ? '\n\nActuel : ' + current : ''), ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const val = (resp.getResponseText() || '').trim();
  if (!val) { ui.alert('ID vide — aucune modification.'); return; }

  props.setProperty('RESPONSES_SSID', val);
  ui.alert('✅ Feuille de réponses configurée.\nID = ' + val + '\nRelance un dry-run.');
}

// ============================================================================
// SECTION - Points d'entrée et Orchestration
// ============================================================================

/**
 * Point d'entrée principal déclenché par la soumission d'un formulaire.
 * @param {Object} e L'objet événement de soumission de formulaire.
 */
function onFormSubmit(e) {
  try {
    Logger.log("Nouvelle réponse reçue, lancement du traitement...");
    traiterLigne(e.range.getRow());
  } catch (err) {
    Logger.log("ERREUR FATALE dans onFormSubmit: " + err.toString() + "\n" + err.stack);
  }
}

/**
 * Orchestre le traitement complet pour une ligne de réponse donnée.
 * @param {number} rowIndex L'index de la ligne à traiter.
 * @param {Object} optionsSurcharge Options pour surcharger le comportement par défaut.
 */
function traiterLigne(rowIndex, optionsSurcharge = {}) {
  try {
    const config = getTestConfiguration();
    const reponse = _creerObjetReponse(rowIndex, optionsSurcharge);
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
    
    assemblerEtEnvoyerEmailUniversel(config, reponse, resultats, langueCible, optionsSurcharge);
  } catch (e) {
    Logger.log("!!!! ERREUR FATALE dans traiterLigne !!!!");
    Logger.log("Message : " + e.message);
    Logger.log("Stack Trace : " + e.stack);
    try {
      const sheet = _getReponsesSheet_({}, {});
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const statutCol = headers.indexOf('Statut_Traitement');
      if (statutCol !== -1) {
        sheet.getRange(rowIndex, statutCol + 1).setValue('ERREUR');
      }
    } catch (e2) {
      Logger.log("Impossible de mettre à jour le statut d'erreur dans la feuille de réponses : " + e2.message);
    }
    throw e;
  }
}

/**
 * Lance le retraitement d'une ligne à partir de l'interface utilisateur.
 * @param {Object} options Les options de retraitement fournies par l'UI.
 * @returns {string} Un message de succès.
 */
function lancerRetraitementDepuisUI(options) {
  try {
    const destinatairesSurcharge = options.destinataires || {};
    destinatairesSurcharge.overrideRecipients = true;
    traiterLigne(options.rowIndex, {
      isRetraitement: true,
      dryRun: false,
      ignoreDeveloppeurEmail: true,
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
 * Exécute un test de retraitement sans envoyer d'email (dry run).
 * @param {number} rowIndex L'index de la ligne à tester.
 * @param {Object} options Les options pour le test.
 */
function retraitementTestSansEnvoi(rowIndex, options) {
  try {
    traiterLigne(rowIndex, {
      isRetraitement: true,
      dryRun: true,
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