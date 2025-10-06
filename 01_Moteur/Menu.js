// =================================================================================
// == PROJET [MOTEUR] - FICHIER INTERFACE UTILISATEUR
// == VERSION : 10.2 (Ajout Étape 3 - Vérification automatique)
// == RÔLE    : Gère l'interface utilisateur (menus et boîtes de dialogue).
// =================================================================================

/**
 * Crée le menu personnalisé à l'ouverture de la feuille de calcul.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏭 Usine à Tests') // Émoji usine pour le titre principal
    .addItem("📦 Étape 1 : Créer les fichiers du Kit", "etape1_creerKit_UI")
    .addItem("⚙️ Étape 2 : Configurer le Kit sélectionné", "etape2_configurerKit_UI")
    .addItem("🔍 Étape 3 : Vérifier le Kit sélectionné", "etape3_verifierKit_UI")
    .addToUi();
}

/**
 * Interface pour lancer l'Étape 1 : Création des fichiers.
 */
function etape1_creerKit_UI() {
  const ui = SpreadsheetApp.getUi();
  try {
    const rowIndex = _getRowFromSelectionOrAsk_(
      "Lancement de l'Étape 1",
      "Entrez le numéro de la ligne à utiliser pour la CRÉATION :"
    );
    ui.alert(
      "Lancement de la création...",
      `Les fichiers vont être créés. La ligne ${rowIndex} sera mise à jour dans quelques instants.`,
      ui.ButtonSet.OK
    );

    etape1_creerKit(rowIndex);

    ui.alert(
      "✅ Étape 1 terminée",
      `Les fichiers ont été créés et les IDs ont été inscrits sur la ligne ${rowIndex}.\nVous pouvez maintenant passer à l'étape 2.`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log(`ERREUR lors du lancement de l'Étape 1 : ${e.toString()}`);
    SpreadsheetApp.getUi().alert(`❌ ERREUR (Étape 1) : ${e.message}`);
  }
}

/**
 * Interface pour lancer l'Étape 2 : Configuration des fichiers.
 */
function etape2_configurerKit_UI() {
  const ui = SpreadsheetApp.getUi();
  try {
    const rowIndex = _getRowFromSelectionOrAsk_(
      "Lancement de l'Étape 2",
      "Entrez le numéro de la ligne à CONFIGURER :"
    );
    ui.alert(
      "Lancement de la configuration...",
      `Le kit de la ligne ${rowIndex} va être configuré. Veuillez patienter.`,
      ui.ButtonSet.OK
    );

    etape2_configurerKit(rowIndex);

    ui.alert(
      "✅ Étape 2 terminée",
      `Le kit de la ligne ${rowIndex} a été configuré avec succès et est prêt à l'emploi.`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log(`ERREUR lors du lancement de l'Étape 2 : ${e.toString()}`);
    SpreadsheetApp.getUi().alert(`❌ ERREUR (Étape 2) : ${e.message}`);
  }
}

/**
 * Interface pour lancer l'Étape 3 : Vérification du kit.
 */
function etape3_verifierKit_UI() {
  const ui = SpreadsheetApp.getUi();
  try {
    const rowIndex = _getRowFromSelectionOrAsk_(
      "Vérification du Kit",
      "Entrez le numéro de la ligne à VÉRIFIER :"
    );
    ui.alert(
      "Lancement de la vérification...",
      `La ligne ${rowIndex} va être vérifiée.`,
      ui.ButtonSet.OK
    );

    etape3_verifierKit(rowIndex);

  } catch (e) {
    Logger.log(`ERREUR lors du lancement de l'Étape 3 : ${e.toString()}`);
    SpreadsheetApp.getUi().alert(`❌ ERREUR (Étape 3) : ${e.message}`);
  }
}

/**
 * Helper pour demander à l'utilisateur le numéro de la ligne à traiter.
 * VERSION CORRIGÉE : demande toujours le numéro de ligne et valide la saisie.
 */
function _getRowFromSelectionOrAsk_(title, promptMessage) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(title, promptMessage, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK || response.getResponseText() === '') {
    throw new Error("Opération annulée par l'utilisateur.");
  }

  const rowIndex = parseInt(response.getResponseText(), 10);
  if (isNaN(rowIndex) || rowIndex <= 1) {
    throw new Error("Numéro de ligne invalide (doit être > 1).");
  }

  return rowIndex;
}
