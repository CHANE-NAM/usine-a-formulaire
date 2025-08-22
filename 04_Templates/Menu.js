// =================================================================================
// == FICHIER : Menu.gs
// == VERSION : 2.3 (Retour à la saisie manuelle systématique du N° de ligne)
// == RÔLE : Crée le menu et gère l'ouverture de l'interface de retraitement.
// =================================================================================

/**
 * S'exécute à l'ouverture de la feuille de calcul pour créer le menu personnalisé.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Actions Usine')
    .addItem("Activer le traitement auto", "activerTraitementAutomatique")
    .addSeparator()
    // MODIFIÉ : Appelle directement la fonction qui ouvre la boîte de dialogue
    .addItem("Retraiter une réponse...", "retraiterReponse_UI")
    .addToUi();
}

/**
 * Ouvre une boîte de dialogue pour demander le numéro de la ligne à retraiter.
 * C'est maintenant la fonction par défaut appelée par le menu.
 */
function retraiterReponse_UI() {
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = HtmlService.createHtmlOutputFromFile('DialogueLigne.html')
    .setWidth(350)
    .setHeight(160);
  ui.showModalDialog(htmlOutput, 'Retraitement de Réponse');
}

/**
 * Ouvre la barre latérale de retraitement pour une ligne donnée.
 * Cette fonction est appelée par le code HTML de 'DialogueLigne.html'.
 */
function ouvrirSidebarPourLigne(rowIndex) {
  const ui = SpreadsheetApp.getUi();
  const template = HtmlService.createTemplateFromFile('RetraitementUI');
  template.ligneActive = rowIndex;
  const htmlOutput = template.evaluate()
    .setTitle("Retraitement - Ligne " + rowIndex)
    .setWidth(350);
  ui.showSidebar(htmlOutput);
}

/**
 * Crée le déclencheur "onFormSubmit" pour le traitement automatique.
 */
function activerTraitementAutomatique() {
  const ss = SpreadsheetApp.getActive();
  const triggers = ScriptApp.getUserTriggers(ss);
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  SpreadsheetApp.getUi().alert('✅ Déclencheur activé ! Le traitement automatique des réponses est maintenant opérationnel.');
}