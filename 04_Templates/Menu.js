// =================================================================================
// == FICHIER : Menu.gs
// == VERSION : 2.6  (onInstall + onOpen unique, titre harmonisé)
// == RÔLE    : Menus : déclencheur, retraitement, dry-run, config classeur de réponses
// == NOTE    : Ce fichier doit être l'UNIQUE endroit où une fonction onOpen() existe.
// =================================================================================

/**
 * Construit le menu du Kit dans la feuille Réponses.
 * ⚠️ Assure-toi qu'aucun autre fichier du projet ne déclare onOpen().
 */
function onOpen(e) {
  try {
    const ui = SpreadsheetApp.getUi();

    // Menu principal (titre harmonisé pour tous les fichiers Réponses)
    const main = ui.createMenu('⚙️ Actions du Kit')
      .addItem('Autoriser le traitement auto', 'activerTraitementAutomatique')
      .addSeparator()
      .addItem('Retraiter une réponse...', 'retraiterReponse_UI');

    // Sous-menu Injecteur (optionnel, seulement si les fonctions existent)
    if (typeof injectScenarioStableLent === 'function') {
      const inj = ui.createMenu('Injecteur')
        .addItem('Stable & Lent',        'injectScenarioStableLent')
        .addItem('Turbulent & Rapide',   'injectScenarioTurbulentRapide')
        .addItem('Mixte',                'injectScenarioMixte')
        .addSeparator()
        .addItem('Stable & Rapide',      'injectScenarioStableRapide')
        .addItem('Instable & Lent',      'injectScenarioInstableLent')
        .addItem('Très K (stable fort)', 'injectScenarioKFort')
        .addItem('Très r (turbulent)',   'injectScenarioRFort')
        .addItem('Alterné',              'injectScenarioAlterne')
        .addItem('Médian',               'injectScenarioMedian')
        .addItem('Stress test x3',       'injectScenarioStressTest');
      main.addSubMenu(inj);
    }

    // Sous-menu Usine à Tests (dry-run, envoi réel, config feuille de réponses)
    const usine = ui.createMenu('Usine à Tests')
      .addItem('Dry-run (dernière ligne)',        'ui_DryRunDerniereLigne')
      .addItem('Dry-run (ligne sélectionnée)',    'ui_DryRunLigneSelection')
      .addSeparator()
      .addItem('ENVOI RÉEL (ligne sélectionnée)', 'ui_EnvoiReelLigneSelection')
      .addSeparator()
      .addItem('Configurer la feuille de réponses…', 'ui_ConfigResponsesSheet');

    main.addSubMenu(usine);
    main.addToUi();

  } catch (err) {
    Logger.log('onOpen() a échoué : ' + err);
  }
}

/**
 * Assure l’apparition du menu à l’installation/copie du projet dans un nouveau classeur.
 * (Apps Script appelle onInstall lors de l’ajout initial du projet au fichier)
 */
function onInstall(e) {
  onOpen(e);
}

/** Ouvre le dialogue de saisie manuelle du numéro de ligne. */
function retraiterReponse_UI() {
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = HtmlService.createHtmlOutputFromFile('DialogueLigne.html')
    .setWidth(350)
    .setHeight(160);
  ui.showModalDialog(htmlOutput, 'Retraitement de Réponse');
}

/** Ouvre la sidebar de retraitement pour une ligne donnée (appelée depuis HTML). */
function ouvrirSidebarPourLigne(rowIndex) {
  const ui = SpreadsheetApp.getUi();
  const template = HtmlService.createTemplateFromFile('RetraitementUI');
  template.ligneActive = rowIndex;
  const htmlOutput = template.evaluate()
    .setTitle('Retraitement - Ligne ' + rowIndex)
    .setWidth(350);
  ui.showSidebar(htmlOutput);
}

/** Crée le déclencheur onFormSubmit pour le traitement automatique. */
function activerTraitementAutomatique() {
  const ss = SpreadsheetApp.getActive();

  // Nettoyage pour éviter les doublons
  ScriptApp.getUserTriggers(ss).forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Création
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  SpreadsheetApp.getUi().alert('✅ Déclencheur activé ! Le traitement automatique des réponses est maintenant opérationnel.');
}

/* ============================================================================
 * SOUS-MENU "Usine à Tests" : helpers & actions
 * Nécessite : getTestConfiguration(), getOriginalLanguage(),
 * _creerObjetReponse(), _getReponsesSheet_(), retraitementTestSansEnvoi(), traiterLigne()
 * ============================================================================ */

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
    if (lr < 2) throw new Error('Feuille vide (seulement l’entête).');

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
