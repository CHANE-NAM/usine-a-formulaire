// ===================================================================
// FICHIER : InjecteurMenuHook.gs
// RÔLE   : Ajout non-intrusif du menu "Injecteur" (sans toucher onOpen())
// VERSION : 1.0
// ===================================================================

// Affiche le menu "Injecteur" sans impacter tes autres menus
function addInjectorMenu_() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Injecteur')
      .addItem('Stable & Lent', 'injectScenarioStableLent')
      .addItem('Turbulent & Rapide', 'injectScenarioTurbulentRapide')
      .addItem('Mixte', 'injectScenarioMixte')
      .addToUi();
  } catch (e) {
    Logger.log('addInjectorMenu_ error: ' + e);
  }
}

// Installe un trigger ON_OPEN qui appelle addInjectorMenu_
// => évite de modifier/écraser ton onOpen() existant
function installInjectorMenuTrigger() {
  const ssId = SpreadsheetApp.getActive().getId();

  // Nettoyage d'éventuels doublons
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'addInjectorMenu_' &&
                 t.getEventType() === ScriptApp.EventType.ON_OPEN)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('addInjectorMenu_')
    .forSpreadsheet(ssId)
    .onOpen()
    .create();

  Logger.log('Trigger du menu Injecteur installé pour le fichier : ' + ssId);
}

// (Optionnel) suppression du trigger si besoin
function uninstallInjectorMenuTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'addInjectorMenu_')
    .forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('Trigger du menu Injecteur supprimé.');
}

// Petit helper de vérification
function debugInjectorTarget_() {
  try {
    if (typeof INJECT_DEFAULT === 'object' && INJECT_DEFAULT.rowIndex) {
      Logger.log('L’injecteur utilisera la ligne CONFIG n° ' + INJECT_DEFAULT.rowIndex);
    } else {
      Logger.log('INJECT_DEFAULT.rowIndex introuvable. Ouvre InjecteurScenarios.gs et renseigne-le.');
    }
  } catch (e) {
    Logger.log('Impossible de lire INJECT_DEFAULT : ' + e);
  }
}
