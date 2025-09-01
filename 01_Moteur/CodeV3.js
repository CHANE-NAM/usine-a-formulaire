// =================================================================================
// == PROJET [MOTEUR] - FICHIER PRINCIPAL (POINTS D'ENTRÉE)
// == VERSION : 8.0 - Architecture multi-fichiers stable
// == RÔLE    : Gère l'interface utilisateur (menus) et orchestre les appels
// ==           vers les scripts de logique métier.
// =================================================================================

/**
 * Crée le menu personnalisé dans l'interface utilisateur de Google Sheets à l'ouverture.
 * C'est le SEUL onOpen() du projet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏭 Usine à Tests')
    .addItem("🚀 Déployer un test de A à Z...", "orchestrateurDeploiementComplet_UI")
    .addToUi();
}

/**
 * Orchestre le déploiement complet d'un test depuis l'UI.
 * Appelle la fonction de logique métier `lancerDeploiementComplet` du fichier Moteur.gs.
 */
function orchestrateurDeploiementComplet_UI() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '🚀 Déploiement de A à Z',
    'Entrez le numéro de la ligne à déployer entièrement :',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK || response.getResponseText() === '') {
    return; // Annulation par l'utilisateur
  }

  const rowIndex = parseInt(response.getResponseText(), 10);
  if (isNaN(rowIndex) || rowIndex <= 1) {
    ui.alert('Numéro de ligne invalide. Veuillez entrer un nombre supérieur à 1.');
    return;
  }
  
  ui.alert('Lancement du déploiement complet... Cette opération peut prendre un moment.');

  try {
    // Appel à la fonction de logique métier
    const resultats = lancerDeploiementComplet(rowIndex);

    if (resultats && resultats.urlSheet && resultats.urlForm) {
      const htmlOutput = HtmlService.createHtmlOutput(
        `<h4>✅ Déploiement Réussi !</h4>` +
        `<p>Le kit "<b>${resultats.nomFichier}</b>" a été généré.</p><hr>` +
        `<p><b>1. Voici le lien public du formulaire à partager :</b></p>` +
        `<p style="margin-top:10px;"><a href="${resultats.urlForm}" target="_blank" style="background-color:#34A853; color:white; padding:8px 12px; text-decoration:none; border-radius:4px;">Copier ou ouvrir le lien du Formulaire</a></p><br>` +
        `<p><b>2. ACTION FINALE REQUISE (pour que le test fonctionne) :</b></p>` +
        `<p>Cliquez sur le lien ci-dessous, puis dans le menu :<br>` +
        `<b>&nbsp;&nbsp;&nbsp;⚙️ Actions du Kit -> Activer le traitement des réponses</b>.</p>` +
        `<p style="margin-top:10px;"><a href="${resultats.urlSheet}" target="_blank" style="background-color:#4285F4; color:white; padding:8px 12px; text-decoration:none; border-radius:4px;">Ouvrir le Kit pour l'activer</a></p>`
      )
      .setWidth(500)
      .setHeight(320);
      ui.showModalDialog(htmlOutput, "Déploiement Terminé");

    } else {
      ui.alert(`ℹ️ Le déploiement pour la ligne ${rowIndex} a été ignoré (le statut n'était probablement pas 'En construction').`);
    }

  } catch (e) {
    Logger.log(`ERREUR Critique lors du déploiement complet (ligne ${rowIndex}) : ${e.toString()}`);
    ui.alert(`❌ ERREUR : Le déploiement a échoué pour la ligne ${rowIndex}. Consultez les logs pour les détails. Message : ${e.message}`);
  }
}

