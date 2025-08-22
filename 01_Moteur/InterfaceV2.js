// =================================================================================
// FICHIER : Interface V2.js
// RÔLE : Création du menu utilisateur et fonctions appelées par ce menu.
// VERSION : 4.2 - Ajout de l'outil de migration V1 -> V2 dans le menu
// =================================================================================

/**
 * Crée le menu personnalisé dans l'interface utilisateur de Google Sheets à l'ouverture.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏭 Usine à Tests')
    .addItem("🚀 Déployer un test de A à Z...", "orchestrateurDeploiementComplet_UI")
    .addSeparator()
    .addItem("Générer un test (choisir la ligne)", "lancerCreationDepuisPilote_UI")
    .addItem("Traiter TOUTES les nouvelles demandes", "orchestrateurCreationAutomatique_UI")
    // --- AJOUT ---
    .addSeparator()
    .addItem("🔧 Migrer les Questions (V1 -> V2)", "lancerMigrationV1versV2")
    .addToUi();
}

/**
 * Orchestre le déploiement complet d'un test depuis l'UI.
 * Gère la génération du kit ET guide l'utilisateur pour l'activation et le partage.
 */
function orchestrateurDeploiementComplet_UI() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '🚀 Déploiement de A à Z',
    'Entrez le numéro de la ligne à déployer entièrement :',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK || response.getResponseText() === '') {
    return;
  }

  const rowIndex = parseInt(response.getResponseText(), 10);
  if (isNaN(rowIndex) || rowIndex <= 1) {
    ui.alert('Numéro de ligne invalide. Veuillez entrer un nombre supérieur à 1.');
    return;
  }
  
  ui.alert('Lancement du déploiement complet... Cette opération peut prendre un moment.');

  try {
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

/**
 * Fonction UI appelée par le menu pour lancer la création manuelle (une seule ligne).
 */
function lancerCreationDepuisPilote_UI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Lancement de la création', 'Entrez le numéro de la ligne à utiliser :', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const rowIndex = parseInt(response.getResponseText(), 10);
  if (isNaN(rowIndex) || rowIndex <= 1) {
    ui.alert('Numéro de ligne invalide. Veuillez entrer un nombre supérieur à 1.');
    return;
  }

  try {
    ui.alert('Lancement de la création... Cette opération peut prendre un moment.');
    
    const resultats = lancerCreationSysteme(rowIndex);

    if (resultats) {
      const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName("Paramètres Généraux");
      
      const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
      const colIndex = {};
      headers.forEach((header, i) => { if (header) colIndex[header] = i; });
      
      const STATUT_COL = colIndex['Statut'];
      const ID_UNIQUE_COL = colIndex['Id_Unique'];
      const NOM_FICHIER_COL = colIndex['Nom_Fichier_Complet'];
      const ID_FORM_COL = colIndex['ID_Formulaire_Cible'];
      const ID_SHEET_COL = colIndex['ID_Sheet_Cible'];

      const idUnique = resultats.sheetFile.getId().slice(0, 8) + '-' + resultats.formFile.getId().slice(0, 8);
      configSheet.getRange(rowIndex, STATUT_COL + 1).setValue('Actif');
      configSheet.getRange(rowIndex, ID_UNIQUE_COL + 1).setValue(idUnique);
      configSheet.getRange(rowIndex, NOM_FICHIER_COL + 1).setValue(resultats.nomFichierComplet);
      if (ID_FORM_COL !== undefined) configSheet.getRange(rowIndex, ID_FORM_COL + 1).setValue(resultats.formFile.getId());
      if (ID_SHEET_COL !== undefined) configSheet.getRange(rowIndex, ID_SHEET_COL + 1).setValue(resultats.sheetFile.getId());
      
      SpreadsheetApp.flush();
      ui.alert(`✅ SUCCÈS : Le test '${resultats.nomFichierComplet}' a été créé et la ligne ${rowIndex} a été mise à jour.`);
      
    } else {
       ui.alert(`ℹ️ La création pour la ligne ${rowIndex} a été ignorée (le statut n'était probablement pas 'En construction').`);
    }

  } catch (e) {
    try {
        const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName("Paramètres Généraux");
        const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
        const statutColIndex = headers.indexOf('Statut');
        if (statutColIndex !== -1) {
            configSheet.getRange(rowIndex, statutColIndex + 1).setValue('ERREUR');
        }
    } catch (err) {
        Logger.log(`Impossible de mettre le statut à ERREUR pour la ligne ${rowIndex}. Erreur : ${err.message}`);
    }

    Logger.log(`ERREUR Critique lors de la création manuelle (ligne ${rowIndex}) : ${e.toString()}`);
    ui.alert(`❌ ERREUR : Une erreur critique est survenue pour la ligne ${rowIndex}. Le statut a été mis à 'ERREUR'. Consultez les logs pour les détails. Message : ${e.message}`);
  }
}

/**
 * Fonction UI pour lancer le traitement en masse de toutes les demandes "En construction".
 */
function orchestrateurCreationAutomatique_UI() {
  const ui = SpreadsheetApp.getUi();
  try {
    const lignesTraitees = orchestrateurCreationAutomatique();
    ui.alert(`Traitement terminé. ${lignesTraitees} nouvelle(s) demande(s) ont été traitée(s).`);
  } catch (e) {
    Logger.log(`ERREUR Critique dans l'orchestrateur : ${e.toString()}`);
    ui.alert(`Une erreur critique est survenue : ${e.message}`);
  }
}