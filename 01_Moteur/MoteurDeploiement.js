// =================================================================================
// == PROJET [MOTEUR] - FICHIER LOGIQUE MÉTIER
// == VERSION : 11.0 (Solution finale par dissociation avec déclencheur temporaire)
// == RÔLE    : Contient la logique de déploiement en deux étapes pour contourner les bugs de la plateforme.
// =================================================================================

/**
 * ÉTAPE 1 : Crée les fichiers et programme l'étape 2.
 * Cette fonction est appelée par l'utilisateur.
 */
function lancerDeploiementComplet_Etape1(rowIndex) {
  Logger.log(`Lancement du déploiement (ÉTAPE 1/2) pour la ligne ${rowIndex}...`);
  const properties = PropertiesService.getUserProperties();
  
  try {
    const config = getConfigurationFromRow(rowIndex);
    const nomFichierComplet = `${new Date().getFullYear()}${('0' + (new Date().getMonth() + 1)).slice(-2)}${('0' + new Date().getDate()).slice(-2)}_${config['Type_Test'] || 'TypeInconnu'}_${config['Statut_Deploiement'] || 'alpha'}_${config['nbQuestions'] || 'Nq'}q`;
    const systemIds = getSystemIds();
    const dossierCible = config['ID_Dossier_Cible'] ? DriveApp.getFolderById(config['ID_Dossier_Cible']) : DriveApp.getFolderById(systemIds.ID_DOSSIER_CIBLE_GEN);

    // Création du Sheet (rapide et stable)
    const templateSheet = SpreadsheetApp.openById(systemIds.ID_TEMPLATE_TRAITEMENT_V2);
    const ssCopy = templateSheet.copy(nomFichierComplet);
    const sheetId = ssCopy.getId();
    
    // Création du Form (rapide mais instable)
    const templateFormFile = DriveApp.getFileById(systemIds.ID_TEMPLATE_FORMULAIRE);
    const formFile = templateFormFile.makeCopy(nomFichierComplet, dossierCible); // Utilise le nom complet du sheet pour cohérence
    const formId = formFile.getId();

    Logger.log(`[ÉTAPE 1] Fichiers créés. ID Form: ${formId}, ID Sheet: ${sheetId}.`);

    // On stocke les IDs pour que l'étape 2 puisse les retrouver
    const state = { formId: formId, sheetId: sheetId, rowIndex: rowIndex, nomFichierComplet: nomFichierComplet };
    properties.setProperty('PENDING_KIT_SETUP', JSON.stringify(state));

    // On supprime les anciens déclencheurs pour éviter les doublons
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getHandlerFunction() === 'lancerDeploiementComplet_Etape2') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // On programme l'exécution de l'étape 2 dans 30 secondes
    ScriptApp.newTrigger('lancerDeploiementComplet_Etape2')
      .timeBased()
      .after(30 * 1000) // 30 secondes
      .create();
      
    Logger.log('[ÉTAPE 1] Terminé. Configuration finale programmée dans 30 secondes.');

  } catch (e) {
    Logger.log(`ERREUR Critique (ÉTAPE 1, ligne ${rowIndex}) : ${e.toString()}\n${e.stack}`);
    throw e;
  }
}


/**
 * ÉTAPE 2 : Configure les fichiers.
 * Cette fonction est appelée automatiquement par un déclencheur.
 */
function lancerDeploiementComplet_Etape2() {
  const properties = PropertiesService.getUserProperties();
  const stateString = properties.getProperty('PENDING_KIT_SETUP');
  
  // Nettoyage immédiat pour éviter de relancer cette config par erreur
  properties.deleteProperty('PENDING_KIT_SETUP');
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'lancerDeploiementComplet_Etape2') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`[ÉTAPE 2] Déclencheur temporaire supprimé.`);
    }
  });

  if (!stateString) {
    Logger.log('[ÉTAPE 2] ERREUR : Aucune configuration en attente trouvée.');
    return;
  }
  
  const state = JSON.parse(stateString);
  const { formId, sheetId, rowIndex, nomFichierComplet } = state;
  Logger.log(`Démarrage de la configuration (ÉTAPE 2/2) pour la ligne ${rowIndex}...`);
  let formFile, sheetFile;

  try {
    // On récupère les handles des fichiers pour le nettoyage en cas d'erreur
    formFile = DriveApp.getFileById(formId);
    sheetFile = DriveApp.getFileById(sheetId);
    
    // On ouvre les fichiers dans cette nouvelle session "propre"
    const form = FormApp.openById(formId);
    
    Logger.log('[ÉTAPE 2] Fichiers ouverts. Application de la configuration...');
    
    // Ces opérations devraient maintenant fonctionner sans erreur
    form.setRequireLogin(false);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, sheetId);
    form.setProgressBar(true);
    
    const config = getConfigurationFromRow(rowIndex); // On relit la config pour être à jour
    form.setTitle(config['Titre_Formulaire_Utilisateur']);
    form.setDescription(config['Sous-Titre_Formulaire'] || "");
    
    const bdd = SpreadsheetApp.openById(getSystemIds().ID_BDD); // On relit les systemIds
    const languesAInclure = _identifierLangues(bdd, config['Type_Test']);
    _construireQuestionsFormulaire(form, languesAInclure, config['nbQuestions']);
    
    Logger.log('[ÉTAPE 2] Configuration et remplissage des questions terminés.');

    // Mise à jour de la feuille CONFIG
    const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName("Paramètres Généraux");
    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    const colIndex = {};
    headers.forEach((header, i) => { if (header) colIndex[header.trim()] = i; });
    
    const urlHandler = getSystemIds().ID_WEBAPP_HANDLER;
    const payload = { rowIndex: rowIndex };
    const encryptedCode = CryptoJS.AES.encrypt(JSON.stringify(payload), "FELIX QUI POTUIT RERUM COGNOCERE CAUSA VIC").toString();
    const urlFinaleChiffree = `${urlHandler}?code=${encodeURIComponent(encryptedCode)}`;

    const idUnique = sheetId.slice(0, 8) + '-' + formId.slice(0, 8);
    configSheet.getRange(rowIndex, colIndex['Id_Unique'] + 1).setValue(idUnique);
    configSheet.getRange(rowIndex, colIndex['Nom_Fichier_Complet'] + 1).setValue(nomFichierComplet);
    configSheet.getRange(rowIndex, colIndex['Lien_Formulaire_Public'] + 1).setValue(urlFinaleChiffree);
    configSheet.getRange(rowIndex, colIndex['ID_Formulaire_Cible'] + 1).setValue(formId);
    configSheet.getRange(rowIndex, colIndex['ID_Sheet_Cible'] + 1).setValue(sheetId);
    configSheet.getRange(rowIndex, colIndex['Accès Direct Formulaire'] + 1).setValue(form.getEditUrl());
    configSheet.getRange(rowIndex, colIndex['Statut'] + 1).setValue('Actif - Prêt');

    Logger.log(`[ÉTAPE 2] Déploiement pour la ligne ${rowIndex} terminé avec SUCCÈS.`);

  } catch (e) {
    Logger.log(`ERREUR Critique (ÉTAPE 2, ligne ${rowIndex}) : ${e.toString()}\n${e.stack}`);
    // On tente de nettoyer les fichiers créés
    if (formFile) { try { formFile.setTrashed(true); } catch(errClean) {} }
    if (sheetFile) { try { sheetFile.setTrashed(true); } catch(errClean) {} }
  }
}