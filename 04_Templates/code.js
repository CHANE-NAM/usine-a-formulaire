/**
 * @OnlyCurrentDoc
 * Script connecteur pour un Kit de Test V4.
 * Ce script sert de pont entre cette feuille de calcul (le "Kit")
 * et la bibliothèque de code centralisée "TEMPLATE".
 */

// ===============================================================
// == DÉCLENCHEURS AUTOMATIQUES GOOGLE
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
 * S'exécute à chaque fois qu'une nouvelle réponse est soumise via le Google Form associé.
 * @param {Object} e L'objet événement fourni par Google.
 */
function onFormSubmit(e) {
  // Ajout d'un log pour vérifier le démarrage
  Logger.log("DÉCLENCHEUR onFormSubmit a démarré pour la ligne : " + e.range.getRow());
  
  // Appelle la fonction principale de traitement dans la bibliothèque
  TEMPLATE.main(e, SpreadsheetApp.getActiveSpreadsheet().getId());
}

// ===============================================================
// == FONCTIONS APPELÉES PAR LE MENU UTILISATEUR
// ===============================================================

/**
 * Installe un déclencheur "installable" robuste qui exécute onFormSubmit
 * à chaque soumission de formulaire.
 */
function installTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Nettoyer les anciens déclencheurs pour éviter les doublons
  const allTriggers = ScriptApp.getUserTriggers(ss);
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // 2. Créer le nouveau déclencheur installable
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
    
  // 3. Informer l'utilisateur
  SpreadsheetApp.getUi().alert('✅ Succès ! Le déclencheur de traitement automatique a été installé. Le système est maintenant pleinement opérationnel.');
}


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

// ===============================================================
// == FONCTIONS "RELAIS" POUR L'INTERFACE UTILISATEUR (HTML)
// ===============================================================

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