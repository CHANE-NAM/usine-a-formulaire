/**
 * =================================================================================
 * == FICHIER : TEMPLATE_Menu.gs (DANS LA BIBLIOTHÈQUE)
 * == VERSION : 2.7 - Ajout d'un point d'entrée pour les kits V4.
 * == RÔLE    : Contient TOUTE la logique des menus et des actions associées.
 * =================================================================================
 */

// =================================================================================
// == SECTION AJOUTÉE POUR COMPATIBILITÉ AVEC LES KITS V4
// =================================================================================

/**
/**
 * Point d'entrée pour les kits V4. Affiche la sidebar de retraitement
 * pour une ligne spécifique.
 * @param {number} rowIndex Le numéro de la ligne, fourni par le kit.
 */
function showAdvancedReprocessingDialog(rowIndex) {
  // Appelle votre fonction existante qui ouvre la sidebar
  ouvrirSidebarPourLigne(rowIndex);
}

// =================================================================================
// == LOGIQUE DE MENU EXISTANTE
// =================================================================================

/**
 * Construit le menu du Kit. Cette fonction est appelée par le relais onOpen() du kit.
 */
function onOpen(e) {
  try {
    const ui = SpreadsheetApp.getUi();

    // Menu principal
    const main = ui.createMenu('⚙️ Actions du Kit')
      .addItem('Autoriser Accès à Google', 'activerTraitementAutomatique')
      .addSeparator()
      .addItem('Retraiter une réponse...', 'retraiterReponse_UI');

    // Sous-menu Injecteur (logique inchangée)
    if (typeof injectScenarioStableLent === 'function') {
      const inj = ui.createMenu('Injecteur')
        .addItem('Stable & Lent', 'injectScenarioStableLent')
        .addItem('Turbulent & Rapide', 'injectScenarioTurbulentRapide')
        .addItem('Stress test x3', 'injectScenarioStressTest');
        main.addSubMenu(inj);
    }

    // Sous-menu Usine à Tests
    const usine = ui.createMenu('Usine à Tests')
      .addItem('Dry-run (dernière ligne)', 'ui_DryRunDerniereLigne')
      .addItem('Dry-run (ligne sélectionnée)', 'ui_DryRunLigneSelection')
      .addSeparator()
      .addItem('ENVOI RÉEL (ligne sélectionnée)', 'ui_EnvoiReelLigneSelection')
      .addSeparator()
      .addItem('Configurer la feuille de réponses…', 'ui_ConfigResponsesSheet');
      main.addSubMenu(usine);
    main.addToUi();

  } catch (err) {
    Logger.log('Erreur dans Bibliothèque - onOpen() : ' + err);
  }
}

/**
 * Assure l’apparition du menu à l’installation.
 */
function onInstall(e) {
  onOpen(e);
}

/** Ouvre le dialogue de saisie manuelle du numéro de ligne. */
function retraiterReponse_UI() {
  const ui = SpreadsheetApp.getUi();
  // Assurez-vous que le fichier 'DialogueLigne.html' est bien dans la bibliothèque
  const htmlOutput = HtmlService.createHtmlOutputFromFile('DialogueLigne.html')
    .setWidth(350)
    .setHeight(160);
    ui.showModalDialog(htmlOutput, 'Retraitement de Réponse');
}

/** Ouvre la sidebar de retraitement (appelée depuis le HTML 'DialogueLigne.html'). */
function ouvrirSidebarPourLigne(rowIndex) {
  const ui = SpreadsheetApp.getUi();
  // Assurez-vous que le fichier 'RetraitementUI.html' est bien dans la bibliothèque
  const template = HtmlService.createTemplateFromFile('RetraitementUI');
    template.ligneActive = rowIndex;
  const htmlOutput = template.evaluate()
    .setTitle('Retraitement - Ligne ' + rowIndex)
    .setWidth(450);
  ui.showSidebar(htmlOutput);
}

/** Crée le déclencheur onFormSubmit. */
function activerTraitementAutomatique() {
  const ss = SpreadsheetApp.getActive();
    ScriptApp.getUserTriggers(ss).forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
    ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  SpreadsheetApp.getUi().alert('✅ Déclencheur activé ! Le traitement automatique est maintenant opérationnel.');
}

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

/** Dry-run sur la dernière ligne. */
function ui_DryRunDerniereLigne() {
  try {
    if (typeof getTestConfiguration !== 'function' || typeof _getReponsesSheet_ !== 'function') {
      SpreadsheetApp.getUi().alert('⚠️ Fonctions manquantes. Vérifiez que le projet contient TraitementReponses.gs');
        return;
    }
    const cfg = getTestConfiguration();
    const sh  = _getReponsesSheet_(cfg, {});
    const lr  = sh.getLastRow();
    if (lr < 2) throw new Error('Feuille vide.');
    const langue = (typeof getOriginalLanguage === 'function' && typeof _creerObjetReponse === 'function') ? (getOriginalLanguage(_creerObjetReponse(lr, {})) || 'FR') : 'FR';
    const niveau = (String(cfg.ID_Gabarit_Email_Repondant || '').replace('RESULTATS_', '').trim() || 'N1');
    if (typeof retraitementTestSansEnvoi !== 'function') {
      SpreadsheetApp.getUi().alert('⚠️ Fonction manquante: retraitementTestSansEnvoi().');
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

/** Dry-run sur la ligne sélectionnée. */
function ui_DryRunLigneSelection() {
  try {
    const row   = _getRowFromSelectionOrAsk_();
    const cfg   = (typeof getTestConfiguration === 'function') ? getTestConfiguration() : {};
    const langue = (typeof getOriginalLanguage === 'function' && typeof _creerObjetReponse === 'function') ? (getOriginalLanguage(_creerObjetReponse(row, {})) || 'FR') : 'FR';
    const niveau = (String(cfg.ID_Gabarit_Email_Repondant || '').replace('RESULTATS_', '').trim() || 'N1');
    if (typeof retraitementTestSansEnvoi !== 'function') {
      SpreadsheetApp.getUi().alert('⚠️ Fonction manquante: retraitementTestSansEnvoi().');
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

/** Envoi réel sur la ligne sélectionnée. */
function ui_EnvoiReelLigneSelection() {
  try {
    const row = _getRowFromSelectionOrAsk_();
    if (typeof traiterLigne !== 'function') {
      SpreadsheetApp.getUi().alert('⚠️ Fonction manquante: traiterLigne().');
      return;
    }
    traiterLigne(row, { isRetraitement: true, dryRun: false, ignoreDeveloppeurEmail: false });
    SpreadsheetApp.getUi().alert('Envoi RÉEL lancé sur la ligne ' + row + '. Voir Journaux.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erreur Envoi réel : ' + e.message);
  }
}

/** Persiste l'ID du vrai classeur de réponses. */
function ui_ConfigResponsesSheet() {
  const ui   = SpreadsheetApp.getUi();
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