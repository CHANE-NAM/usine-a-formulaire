// =================================================================================
// == PROJET [MOTEUR] - FICHIER INTERFACE UTILISATEUR
// == VERSION : 9.2 (Adaptation pour le déploiement asynchrone en 2 étapes)
// == RÔLE    : Gère l'interface utilisateur (menus et boîtes de dialogue).
// =================================================================================

// =================================================================================
// == INTERRUPTEUR DE DÉBOGAGE
// =================================================================================
// ==> METTEZ CETTE VARIABLE À 'false' POUR DÉSACTIVER LES LOGS DÉTAILLÉS.
const DEBUG_MODE_MENU = true;

/**
 * Journalise un message de débogage uniquement si le mode débogage est activé.
 * @param {string} message Le message à journaliser.
 */
function debugLogMenu(message) {
  if (DEBUG_MODE_MENU) {
    Logger.log(`[DEBUG MENU] ${message}`);
  }
}

/**
 * Crée le menu personnalisé à l'ouverture de la feuille de calcul.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏭 Usine à Tests')
    .addItem("🚀 Déployer un Test (Automatique)", "orchestrateurDeploiementComplet_UI")
    .addToUi();
}

/**
 * Orchestre le lancement du déploiement en 2 étapes depuis l'UI.
 * Affiche une boîte de dialogue à l'utilisateur et lance la première étape du processus.
 */
function orchestrateurDeploiementComplet_UI() {
  debugLogMenu("Ouverture de la boîte de dialogue de déploiement...");
  const ui = SpreadsheetApp.getUi();
  const response =
    ui.prompt(
      '🚀 Déploiement Asynchrone',
      'Entrez le numéro de la ligne à déployer entièrement :',
      ui.ButtonSet.OK_CANCEL
    );

  if (response.getSelectedButton() !== ui.Button.OK || response.getResponseText() === '') {
    debugLogMenu("Déploiement annulé par l'utilisateur.");
    return;
  }

  const rowIndex = parseInt(response.getResponseText(), 10);
  if (isNaN(rowIndex) || rowIndex <= 1) {
    ui.alert('Numéro de ligne invalide.');
    debugLogMenu(`Tentative de déploiement avec une ligne invalide : ${response.getResponseText()}`);
    return;
  }

  debugLogMenu(`Lancement de l'étape 1 pour la ligne ${rowIndex}...`);
  try {
    // On ne lance que l'étape 1. La fonction Etape 2 sera appelée automatiquement par un déclencheur.
    lancerDeploiementComplet_Etape1(rowIndex);

    // On affiche un message de confirmation simple pour informer l'utilisateur.
    // Il n'y a plus de liens à afficher ici, car ils seront générés lors de l'étape 2.
    ui.alert(
      '✅ Processus Lancé',
      'La création du kit a démarré en arrière-plan. La ligne ' + rowIndex + ' sera mise à jour automatiquement dans environ 1 minute.',
      ui.ButtonSet.OK
    );
    debugLogMenu("Message de confirmation affiché à l'utilisateur.");

  } catch (e) {
    Logger.log(`ERREUR Critique lors du lancement du déploiement (ligne ${rowIndex}) : ${e.toString()}`);
    ui.alert(`❌ ERREUR : Le lancement a échoué. Consultez les logs. Message : ${e.message}`);
  }
}