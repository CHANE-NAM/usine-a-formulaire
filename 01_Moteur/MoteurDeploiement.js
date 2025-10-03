// =================================================================================
// == PROJET [MOTEUR] - FICHIER LOGIQUE MÉTIER
// == VERSION : 9.2 (Correction via SpreadsheetApp.copy())
// == RÔLE    : Contient la logique principale de déploiement d'un kit de test.
// =================================================================================

/**
 * Gère le déploiement complet en utilisant une liaison par déclencheur (Plan B Final).
 * VERSION AVEC CORRECTION ET ESPIONS DE DÉBOGAGE.
 */
function lancerDeploiementComplet(rowIndex) {
  Logger.log(`Lancement du déploiement (Plan B - Déclencheur) pour la ligne ${rowIndex}...`);
  let formFile, sheetFile; // sheetFile est conservé pour la gestion d'erreur (nettoyage)

  try {
    Logger.log("[ESPION 1] Avant getConfigurationFromRow...");
    const config = getConfigurationFromRow(rowIndex);
    Logger.log("[ESPION 1] Succès getConfigurationFromRow.");

    if (config['Statut'].trim().toLowerCase() !== 'en construction') {
      Logger.log(`Déploiement ignoré pour la ligne ${rowIndex} (statut non valide).`);
      return null;
    }

    Logger.log("[ESPION 2] Avant ouverture de la feuille de configuration...");
    const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName("Paramètres Généraux");
    Logger.log("[ESPION 2] Succès ouverture feuille de configuration.");

    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    const colIndex = {};
    headers.forEach((header, i) => { if (header) colIndex[header.trim()] = i; });
    const nomFichierComplet = `${new Date().getFullYear()}${('0' + (new Date().getMonth() + 1)).slice(-2)}${('0' + new Date().getDate()).slice(-2)}_${config['Type_Test'] || 'TypeInconnu'}_${config['Statut_Deploiement'] || 'alpha'}_${config['nbQuestions'] || 'Nq'}q`;
    
    Logger.log("[ESPION 3] Avant getSystemIds...");
    const systemIds = getSystemIds();
    Logger.log("[ESPION 3] Succès getSystemIds.");

    if (!systemIds.ID_TEMPLATE_TRAITEMENT_V2) throw new Error("ID_TEMPLATE_TRAITEMENT_V2 introuvable.");
    const dossierCible = config['ID_Dossier_Cible'] ? DriveApp.getFolderById(config['ID_Dossier_Cible']) : DriveApp.getFolderById(systemIds.ID_DOSSIER_CIBLE_GEN);

    Logger.log("[ESPION 4] Préparation terminée, début de la création des fichiers.");
    const form = FormApp.create(config['Titre_Formulaire_Utilisateur']);
    formFile = DriveApp.getFileById(form.getId());
    formFile.moveTo(dossierCible);

    // --- DEBUT DE LA MODIFICATION (SOLUTION 2) ---
    Logger.log("[SOLUTION 2] Ouverture du modèle de feuille de calcul via SpreadsheetApp...");
    const templateSheet = SpreadsheetApp.openById(systemIds.ID_TEMPLATE_TRAITEMENT_V2);
    
    Logger.log("[SOLUTION 2] Copie du modèle via .copy()...");
    const ssCopy = templateSheet.copy(nomFichierComplet); // Renvoie un objet Spreadsheet
    const sheetId = ssCopy.getId();
    
    // On récupère l'objet File pour la cohérence du reste du script (nettoyage en cas d'erreur)
    sheetFile = DriveApp.getFileById(sheetId);
    
    Logger.log(`[SOLUTION 2] Déplacement du nouveau classeur (ID: ${sheetId}) vers le dossier cible...`);
    sheetFile.moveTo(dossierCible); // On utilise l'objet File pour le déplacer
    
    const finalSheetUrl = ssCopy.getUrl(); // On obtient l'URL directement de l'objet Spreadsheet
    // --- FIN DE LA MODIFICATION (SOLUTION 2) ---

    Logger.log(`[ESPION 4] Fichiers créés. ID Formulaire: ${formFile.getId()}, ID Feuille: ${sheetId}`);
    Logger.log(`[ESPION 6] URL de la feuille obtenue avec succès via ssCopy.getUrl(): ${finalSheetUrl}`);
    
    ScriptApp.newTrigger('onFormSubmitTrigger')
        .forForm(form)
        .onFormSubmit()
        .create();
    Logger.log(`>>> SUCCÈS : Déclencheur 'onFormSubmitTrigger' installé sur le formulaire ${form.getId()}.`);

    form.setProgressBar(true);
    form.setDescription(config['Sous-Titre_Formulaire'] || "");

    Logger.log("[ESPION 5] Avant ouverture de la BDD...");
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    Logger.log("[ESPION 5] Succès ouverture de la BDD.");
    
    const languesAInclure = _identifierLangues(bdd, config['Type_Test']);
    _construireQuestionsFormulaire(form, languesAInclure, config['nbQuestions']);

    const formUrl = form.getPublishedUrl();
    const editUrl = form.getEditUrl();
    const idUnique = sheetId.slice(0, 8) + '-' + formFile.getId().slice(0, 8); // Utilisation de sheetId
    const urlHandler = systemIds.ID_WEBAPP_HANDLER;
    if (!urlHandler) throw new Error("ID_WEBAPP_HANDLER introuvable dans sys_ID_Fichiers.");
    const payload = { rowIndex: rowIndex };
    const encryptedCode = CryptoJS.AES.encrypt(JSON.stringify(payload), "FELIX QUI POTUIT RERUM COGNOCERE CAUSA VIC").toString();
    const urlFinaleChiffree = `${urlHandler}?code=${encodeURIComponent(encryptedCode)}`;

    // Mise à jour de la feuille CONFIG
    configSheet.getRange(rowIndex, colIndex['Statut'] + 1).setValue('Actif - Prêt');
    // ... (Mise à jour des autres colonnes)
    
    return { nomFichier: nomFichierComplet, urlSheet: finalSheetUrl, urlForm: urlFinaleChiffree };

  } catch(e) {
    Logger.log(`ERREUR (Plan B, ligne ${rowIndex}) : ${e.toString()}\n${e.stack}`);
    if (formFile) { try { formFile.setTrashed(true); } catch(errClean) { Logger.log("Erreur lors du nettoyage de formFile: " + errClean.message); } }
    if (sheetFile) { try { sheetFile.setTrashed(true); } catch(errClean) { Logger.log("Erreur lors du nettoyage de sheetFile: " + errClean.message); } }
    throw e;
  }
}
