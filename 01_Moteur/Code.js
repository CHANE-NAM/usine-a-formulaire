// =================================================================================
// == PROJET [MOTEUR] - FICHIER PRINCIPAL (POINTS D'ENTRÉE ET LOGIQUE)
// == VERSION : 9.0 - Plan B (Liaison par Déclencheur)
// == RÔLE    : Gère le menu et contient la logique de déploiement principale.
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
      ui.showModalDialog(htmlOutput, "Déploiement Terminé");
    } else {
      ui.alert(`ℹ️ Le déploiement pour la ligne ${rowIndex} a été ignoré (statut non valide).`);
    }

  } catch (e) {
    Logger.log(`ERREUR Critique lors du déploiement (ligne ${rowIndex}) : ${e.toString()}`);
    ui.alert(`❌ ERREUR : Le déploiement a échoué. Consultez les logs. Message : ${e.message}`);
  }
}

/**
 * Gère le déploiement complet en utilisant une liaison par déclencheur (Plan B Final).
 */
function lancerDeploiementComplet(rowIndex) {
  Logger.log(`Lancement du déploiement (Plan B - Déclencheur) pour la ligne ${rowIndex}...`);
  let formFile, sheetFile;
  
  try {
    const config = getConfigurationFromRow(rowIndex);
    if (config['Statut'].trim().toLowerCase() !== 'en construction') {
      Logger.log(`Déploiement ignoré pour la ligne ${rowIndex} (statut non valide).`);
      return null;
    }
    
    // --- Logique de nommage et de récupération des dossiers ---
    const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName("Paramètres Généraux");
    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    const colIndex = {};
    headers.forEach((header, i) => { if (header) colIndex[header.trim()] = i; });
    const nomFichierComplet = `${new Date().getFullYear()}${('0' + (new Date().getMonth() + 1)).slice(-2)}${('0' + new Date().getDate()).slice(-2)}_${config['Type_Test'] || 'TypeInconnu'}_${config['Statut_Deploiement'] || 'alpha'}_${config['nbQuestions'] || 'Nq'}q`;
    const systemIds = getSystemIds();
    if (!systemIds.ID_TEMPLATE_TRAITEMENT_V2) throw new Error("ID_TEMPLATE_TRAITEMENT_V2 introuvable.");
    const dossierCible = config['ID_Dossier_Cible'] ? DriveApp.getFolderById(config['ID_Dossier_Cible']) : DriveApp.getFolderById(systemIds.ID_DOSSIER_CIBLE_GEN);

    // --- ÉTAPE 1 : Création des fichiers séparément ---
    const form = FormApp.create(config['Titre_Formulaire_Utilisateur']);
    formFile = DriveApp.getFileById(form.getId());
    formFile.moveTo(dossierCible);

    const templateFile = DriveApp.getFileById(systemIds.ID_TEMPLATE_TRAITEMENT_V2);
    sheetFile = templateFile.makeCopy(nomFichierComplet, dossierCible);
    
    // --- ÉTAPE 2 : LIAISON PAR DÉCLENCHEUR ---
    ScriptApp.newTrigger('onFormSubmitTrigger')
        .forForm(form)
        .onFormSubmit()
        .create();

    Logger.log(`>>> SUCCÈS : Déclencheur 'onFormSubmitTrigger' installé sur le formulaire ${form.getId()}.`);
        
    // --- ÉTAPE 3 : Finalisation du déploiement ---
    form.setProgressBar(true);
    form.setDescription(config['Sous-Titre_Formulaire'] || "");
    
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const languesAInclure = _identifierLangues(bdd, config['Type_Test']);
    _construireQuestionsFormulaire(form, languesAInclure, config['nbQuestions']);
    
    const formUrl = form.getPublishedUrl();
    const editUrl = form.getEditUrl();
    const idUnique = sheetFile.getId().slice(0, 8) + '-' + formFile.getId().slice(0, 8);
    
    // Génération de l'URL chiffrée
    const urlHandler = systemIds.ID_WEBAPP_HANDLER;
    if (!urlHandler) throw new Error("ID_WEBAPP_HANDLER introuvable dans sys_ID_Fichiers.");
    const payload = { rowIndex: rowIndex };
    const encryptedCode = CryptoJS.AES.encrypt(JSON.stringify(payload), "FELIX QUI POTUIT RERUM COGNOCERE CAUSA VIC").toString();
    const urlFinaleChiffree = `${urlHandler}?code=${encodeURIComponent(encryptedCode)}`;

    // Mise à jour de la feuille CONFIG
    configSheet.getRange(rowIndex, colIndex['Statut'] + 1).setValue('Actif - Prêt');
    configSheet.getRange(rowIndex, colIndex['Id_Unique'] + 1).setValue(idUnique);
    configSheet.getRange(rowIndex, colIndex['Nom_Fichier_Complet'] + 1).setValue(nomFichierComplet);
    configSheet.getRange(rowIndex, colIndex['ID_Formulaire_Cible'] + 1).setValue(formFile.getId());
    configSheet.getRange(rowIndex, colIndex['ID_Sheet_Cible'] + 1).setValue(sheetFile.getId());
    configSheet.getRange(rowIndex, colIndex['Lien_Formulaire_Public'] + 1).setValue(formUrl);
    
    // --- LIGNE AJOUTÉE ---
    if (colIndex['Lien_Formulaire_Obfusqué'] !== undefined) {
      configSheet.getRange(rowIndex, colIndex['Lien_Formulaire_Obfusqué'] + 1).setValue(urlFinaleChiffree);
    }
    
    const colNameEditUrl = Object.keys(colIndex).find(k => k.toLowerCase().includes('accès direct formulaire'));
    if (colNameEditUrl) {
      configSheet.getRange(rowIndex, colIndex[colNameEditUrl] + 1).setFormula(`=HYPERLINK("${editUrl}"; "Ouvrir le formulaire")`);
    }
    SpreadsheetApp.flush();
    
    return { nomFichier: nomFichierComplet, urlSheet: sheetFile.getUrl(), urlForm: urlFinaleChiffree };
    
  } catch(e) {
    Logger.log(`ERREUR (Plan B, ligne ${rowIndex}) : ${e.toString()}\n${e.stack}`);
    if (formFile) { formFile.setTrashed(true); }
    if (sheetFile) { sheetFile.setTrashed(true); }
    throw e;
  }
}