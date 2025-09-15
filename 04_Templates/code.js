

// ===============================================================
// == DÉCLENCHEURS ET MENUS
// ===============================================================

/**
 * S'exécute à l'ouverture de la feuille de calcul pour créer le menu personnalisé.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Actions du Kit')
    .addItem("✅ Activer le traitement automatique", "installTrigger")
    .addSeparator()
    .addItem("Relancer le traitement d'une ligne...", "runReprocessing")
    .addToUi();
}

/**
 * Installe un déclencheur "installable" robuste.
 * C'est la fonction appelée par le menu.
 */
function installTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const functionToTrigger = 'handleFormSubmit';

  // 1. Nettoyer les anciens déclencheurs pour éviter tout conflit (les anciens et ceux avec l'ancien nom)
  const allTriggers = ScriptApp.getUserTriggers(ss);
  for (const trigger of allTriggers) {
    const handlerFunction = trigger.getHandlerFunction();
    if (handlerFunction === functionToTrigger || handlerFunction === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // 2. Créer le nouveau déclencheur installable qui pointe vers notre fonction renommée
  ScriptApp.newTrigger(functionToTrigger)
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  // 3. Informer l'utilisateur
  SpreadsheetApp.getUi().alert('✅ Succès ! Le déclencheur installable a été activé sur la fonction ' + functionToTrigger + '.');
}

// ===============================================================
// == FONCTION EXÉCUTÉE PAR LE DÉCLENCHEUR
// ===============================================================

/**
 * Fonction cible pour le déclencheur onFormSubmit.
 * Elle n'est PAS un déclencheur "simple" car elle n'est pas nommée "onFormSubmit".
 * C'est cette fonction qui sera maintenant exécutée avec les permissions complètes.
 * @param {Object} e L'objet événement fourni par Google.
 */
function handleFormSubmit(e) {
  Logger.log("DÉCLENCHEUR INSTALLABLE handleFormSubmit a démarré pour la ligne : " + e.range.getRow());
  TEMPLATE.main(e, SpreadsheetApp.getActiveSpreadsheet().getId());
}


// ===============================================================
// == FONCTIONS "RELAIS" POUR L'INTERFACE UTILISATEUR (HTML)
// ===============================================================

/**
 * Affiche une boîte de dialogue pour demander à l'utilisateur le numéro de la ligne à retraiter,
 * puis appelle la bibliothèque pour afficher l'interface de retraitement (sidebar).
 */
function runReprocessing() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Retraitement de ligne',
    'Veuillez entrer le numéro de la ligne à retraiter :',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const rowNum = parseInt(response.getResponseText());
    if (rowNum && rowNum > 1) {
      // Appelle la fonction de la bibliothèque qui ouvre la sidebar
      TEMPLATE.showAdvancedReprocessingDialog(rowNum);
    } else {
      ui.alert('Numéro de ligne invalide.');
    }
  }
}

/**
 * RELAIS #1 : Récupère les données nécessaires à l'affichage de la sidebar.
 * @param {number} rowIndex Le numéro de la ligne à retraiter.
 * @returns {Object} Les données de la ligne (nom, email, etc.).
 */
function getDonneesPourRetraitement(rowIndex) {
  const kitId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return TEMPLATE.getDonneesPourRetraitement(rowIndex, kitId);
}

/**
 * RELAIS #2 : Déclenche le retraitement complet depuis l'interface (sidebar).
 * @param {Object} options Les options sélectionnées dans l'interface.
 * @returns {string} Un message de succès ou d'erreur.
 */
function lancerRetraitementDepuisUI(options) {
  options = options || {};
  options.kitId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return TEMPLATE.lancerRetraitementDepuisUI(options);
}

// ===============================================================
// == FONCTIONS DE TEST D'AUTORISATION
// ===============================================================

/**
 * Fonction temporaire pour forcer la demande d'autorisation pour lire les alias Gmail.
 */
function testAuthGmail() {
  const aliases = GmailApp.getAliases();
  Logger.log('Alias disponibles : ' + aliases);
}

/**
 * Fonction temporaire pour forcer la demande d'autorisation pour Google Docs.
 */
function testAuthDocs() {
  try {
    // Tente d'ouvrir un document fictif pour déclencher la demande.
    DocumentApp.openById('12345_dummy_id_for_auth');
  } catch (e) {
    Logger.log('La demande d\'autorisation pour Google Docs a été déclenchée.');
  }
}