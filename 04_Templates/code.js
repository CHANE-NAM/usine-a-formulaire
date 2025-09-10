/**
 * @OnlyCurrentDoc
 * Ce script connecteur sert de pont entre cette feuille de calcul et la bibliothèque de code centralisée.
 * Il identifie le kit actuel et relaie tous les appels vers la bibliothèque en transmettant son ID.
 */

// ===============================================================
// == FONCTIONS DÉCLENCHÉES PAR GOOGLE (TRIGGERS)
// ===============================================================

function onOpen() {
  TEMPLATE.onOpen();
}

function onInstall(e) {
  TEMPLATE.onInstall(e);
}

function onFormSubmit(e) {
  const kitId = SpreadsheetApp.getActiveSpreadsheet().getId();
  TEMPLATE.onFormSubmit(e, kitId);
}


// ===============================================================
// == FONCTIONS APPELÉES PAR LES MENUS
// ===============================================================

function activerTraitementAutomatique() {
  TEMPLATE.activerTraitementAutomatique();
}

function retraiterReponse_UI() {
  TEMPLATE.retraiterReponse_UI();
}

function ui_DryRunDerniereLigne() {
  const kitId = SpreadsheetApp.getActiveSpreadsheet().getId();
  TEMPLATE.ui_DryRunDerniereLigne(kitId);
}

function ui_DryRunLigneSelection() {
  const kitId = SpreadsheetApp.getActiveSpreadsheet().getId();
  TEMPLATE.ui_DryRunLigneSelection(kitId);
}

function ui_EnvoiReelLigneSelection() {
  const kitId = SpreadsheetApp.getActiveSpreadsheet().getId();
  TEMPLATE.ui_EnvoiReelLigneSelection(kitId);
}

function ui_ConfigResponsesSheet() {
  TEMPLATE.ui_ConfigResponsesSheet();
}

function injectScenarioStableLent() {
  const kitId = SpreadsheetApp.getActiveSpreadsheet().getId();
  TEMPLATE.injectScenarioStableLent(kitId);
}


// ===============================================================
// == FONCTIONS "RELAIS" POUR LES DIALOGUES HTML (google.script.run)
// ===============================================================

function ouvrirSidebarPourLigne(rowIndex) {
  TEMPLATE.ouvrirSidebarPourLigne(rowIndex);
}

function getDonneesPourRetraitement(rowIndex) {
  const kitId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return TEMPLATE.getDonneesPourRetraitement(rowIndex, kitId);
}

function lancerRetraitementDepuisUI(options) {
  options.kitId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return TEMPLATE.lancerRetraitementDepuisUI(options);
}