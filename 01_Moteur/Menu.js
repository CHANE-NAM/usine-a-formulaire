// =================================================================================
// == PROJET [MOTEUR] - FICHIER INTERFACE UTILISATEUR
// == VERSION : 9.1
// == RÔLE    : Gère l'interface utilisateur (menus et boîtes de dialogue).
// =================================================================================

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
 * Orchestre le déploiement complet et automatique d'un test depuis l'UI.
 */
function orchestrateurDeploiementComplet_UI() {
  const ui = SpreadsheetApp.getUi();
  const response =
    ui.prompt(
      '🚀 Déploiement de A à Z',
      'Entrez le numéro de la ligne à déployer entièrement :',
      ui.ButtonSet.OK_CANCEL
    );

  if (response.getSelectedButton() !== ui.Button.OK ||
    response.getResponseText() === '') {
    return;
  }

  const rowIndex = parseInt(response.getResponseText(), 10);
  if (isNaN(rowIndex) || rowIndex <= 1) {
    ui.alert('Numéro de ligne invalide.');
    return;
  }

  ui.alert('Lancement du déploiement (Plan B)... Cette opération peut prendre un moment.');
  try {
    const resultats = lancerDeploiementComplet(rowIndex);

    if (resultats && resultats.urlSheet && resultats.urlForm) {
      const htmlOutput = HtmlService.createHtmlOutput(
          `<h4>✅ Déploiement Réussi ! (Plan B)</h4>` +
          `<p>Le kit "<b>${resultats.nomFichier}</b>" a été généré.</p><hr>` +
          `<p><b>1. Voici le lien public du formulaire à partager (URL chiffrée) :</b></p>` +
          `<p style="margin-top:10px;"><a href="${resultats.urlForm}" target="_blank" style="background-color:#34A853; color:white; padding:8px 12px; text-decoration:none; border-radius:4px;">Ouvrir le lien du Formulaire d'identification</a></p><br>` +
          `<p><b>2. ACTION FINALE REQUISE (pour que le test fonctionne) :</b></p>` +
          `<p>Cliquez sur le lien ci-dessous pour ouvrir le nouveau Kit, puis dans le menu :<br>` +
          `<b>&nbsp;&nbsp;&nbsp;⚙️ Actions du Kit -> Activer le traitement automatique</b>.</p>` +
          `<p style="margin-top:10px;"><a href="${resultats.urlSheet}" target="_blank" style="background-color:#4285F4; color:white; padding:8px 12px; text-decoration:none; border-radius:4px;">Ouvrir le Kit pour l'activer</a></p>`
        )
        .setWidth(500)
        .setHeight(420);
      ui.showModalDialog(htmlOutput, "Déploiement Terminée");
    } else {
      ui.alert(`ℹ️ Le déploiement pour la ligne ${rowIndex} a été ignoré (statut non valide).`);
    }

  } catch (e) {
    Logger.log(`ERREUR Critique lors du déploiement (ligne ${rowIndex}) : ${e.toString()}`);
    ui.alert(`❌ ERREUR : Le déploiement a échoué. Consultez les logs. Message : ${e.message}`);
  }
}