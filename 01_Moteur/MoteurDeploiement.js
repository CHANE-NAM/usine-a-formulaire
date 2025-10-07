// =================================================================================
// == PROJET [MOTEUR] - FICHIER LOGIQUE MÉTIER
// == VERSION : 12.9 (Correction finale avec FormApp.setRequireSignIn)
// == RÔLE    : Contient la logique de déploiement manuel en trois étapes.
// =================================================================================

/**
 * ÉTAPE 1 : Crée les fichiers (Form & Sheet) et met à jour la ligne de configuration avec les IDs.
 */
function etape1_creerKit(rowIndex) {
  Logger.log(`Lancement de l'Étape 1 (Création) pour la ligne ${rowIndex}...`);
  try {
    const config = getConfigurationFromRow(rowIndex);
    const now = new Date();
    const y = now.getFullYear();
    const m = ('0' + (now.getMonth() + 1)).slice(-2);
    const d = ('0' + now.getDate()).slice(-2);
    const nomFichierComplet = `${y}${m}${d}_${config['Type_Test'] || 'TypeInconnu'}_${config['Statut_Deploiement'] || 'alpha'}_${config['nbQuestions'] || 'Nq'}q`;

    const systemIds = getSystemIds();
    const dossierCible = config['ID_Dossier_Cible']
      ? DriveApp.getFolderById(config['ID_Dossier_Cible'])
      : DriveApp.getFolderById(systemIds.ID_DOSSIER_CIBLE_GEN);

    const templateSheet = SpreadsheetApp.openById(systemIds.ID_TEMPLATE_TRAITEMENT_V2);
    const ssCopy = templateSheet.copy(nomFichierComplet);
    const sheetId = ssCopy.getId();

    const templateFormFile = DriveApp.getFileById(systemIds.ID_TEMPLATE_FORMULAIRE);
    const formFile = templateFormFile.makeCopy(nomFichierComplet, dossierCible);
    const formId = formFile.getId();
    Logger.log(`[ÉTAPE 1] Fichiers créés. ID Form: ${formId}, ID Sheet: ${sheetId}.`);

    const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName('Paramètres Généraux');
    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    const colIndex = {};
    headers.forEach((header, i) => { if (header) colIndex[String(header).trim()] = i; });

    const idUnique = sheetId.slice(0, 8) + '-' + formId.slice(0, 8);
    configSheet.getRange(rowIndex, colIndex['Id_Unique'] + 1).setValue(idUnique);
    configSheet.getRange(rowIndex, colIndex['Nom_Fichier_Complet'] + 1).setValue(nomFichierComplet);
    configSheet.getRange(rowIndex, colIndex['ID_Formulaire_Cible'] + 1).setValue(formId);
    configSheet.getRange(rowIndex, colIndex['ID_Sheet_Cible'] + 1).setValue(sheetId);
    configSheet.getRange(rowIndex, colIndex['Statut'] + 1).setValue('Fichiers créés - En attente de configuration');

    SpreadsheetApp.flush();
    Logger.log(`[ÉTAPE 1] Terminé. La ligne ${rowIndex} a été mise à jour.`);

  } catch (e) {
    Logger.log(`ERREUR Critique (ÉTAPE 1, ligne ${rowIndex}) : ${e.toString()}\n${e.stack}`);
    throw e;
  }
}


/**
 * ÉTAPE 2 : Configure les fichiers en utilisant l'API REST Google Forms et finalise le statut.
 */
function etape2_configurerKit(rowIndex) {
  Logger.log(`Lancement de l'Étape 2 (Configuration via API) pour la ligne ${rowIndex}...`);
  try {
    const config = getConfigurationFromRow(rowIndex);
    const formId = config['ID_Formulaire_Cible'];
    const sheetId = config['ID_Sheet_Cible'];

    if (!formId || !sheetId) {
      throw new Error('ID de formulaire ou de feuille manquant sur la ligne ' + rowIndex + '. Lancez d\'abord l\'Étape 1.');
    }

    // --- Appel 1 : Mise à jour du titre et du mode quiz via UrlFetchApp ---
    const batchUpdateRequest = {
      requests: [
        {
          updateFormInfo: {
            info: {
              title: config['Titre_Formulaire_Utilisateur'] || 'Formulaire sans titre',
              description: config['Sous-Titre_Formulaire'] || ''
            },
            updateMask: 'title,description'
          }
        },
        {
          updateSettings: {
            settings: { quizSettings: { isQuiz: false } },
            updateMask: 'quizSettings.isQuiz'
          }
        }
      ]
    };
    const batchUpdateUrl = `https://forms.googleapis.com/v1/forms/${formId}:batchUpdate`;
    const batchUpdateOptions = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(batchUpdateRequest),
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
    const batchUpdateResponse = UrlFetchApp.fetch(batchUpdateUrl, batchUpdateOptions);
    Logger.log(`[ÉTAPE 2] Réponse API (Titre/Quiz): ${batchUpdateResponse.getResponseCode()} - ${batchUpdateResponse.getContentText()}`);
    if (batchUpdateResponse.getResponseCode() >= 300) {
      throw new Error('Erreur API Forms (Titre/Quiz) : ' + batchUpdateResponse.getContentText());
    }
    Logger.log(`[ÉTAPE 2] Titre et paramètres de base mis à jour avec succès.`);

    // --- Appel 2 (FINAL) : Publication via le service FormApp ---
    const form = FormApp.openById(formId);
    Logger.log(`[ÉTAPE 2] Publication du formulaire pour accès public via FormApp...`);
    try {

      form.setLimitOneResponsePerUser(false); // Correction : Le nom correct de la fonction.
        Logger.log(`[ÉTAPE 2] Formulaire rendu accessible publiquement.`);
    } catch(e) {
        Logger.log(`[ÉTAPE 2] AVERTISSEMENT : Échec de la publication automatique : ${e.message}`);
        SpreadsheetApp.getUi().alert('Avertissement', 'La publication automatique du formulaire a échoué. Vous devrez peut-être le faire manuellement. Détails : ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    // --- Suite de la fonction... ---
    const bdd = SpreadsheetApp.openById(getSystemIds().ID_BDD);
    const languesAInclure = _identifierLangues(bdd, config['Type_Test']);

    _construireQuestionsFormulaire(form, bdd, config, languesAInclure);
    Logger.log('[ÉTAPE 2] Construction des questions terminée.');

    try {
      form.setDestination(FormApp.DestinationType.SPREADSHEET, sheetId);
      Logger.log('[ÉTAPE 2] Liaison de la feuille de réponses réussie.');
    } catch (e) {
      Logger.log(`[ÉTAPE 2] AVERTISSEMENT : Liaison manuelle requise (${e.message})`);
      SpreadsheetApp.getUi().alert('Liaison manuelle requise', 'La liaison doit être faite manuellement.\n\nOuvrez le formulaire, onglet "Réponses" → Sélectionnez la destination.', SpreadsheetApp.getUi().ButtonSet.OK);
    }

    const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName('Paramètres Généraux');
    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    const colIndex = {};
    headers.forEach((header, i) => { if (header) colIndex[String(header).trim()] = i; });

    const urlHandler = getSystemIds().ID_WEBAPP_HANDLER;
    const payload = { rowIndex: rowIndex };
    const encryptedCode = CryptoJS.AES.encrypt(JSON.stringify(payload), 'FELIX QUI POTUIT RERUM COGNOCERE CAUSA VIC').toString();
    const urlFinaleChiffree = `${urlHandler}?code=${encodeURIComponent(encryptedCode)}`;

    configSheet.getRange(rowIndex, colIndex['Lien_Formulaire_Public'] + 1).setValue(urlFinaleChiffree);
    configSheet.getRange(rowIndex, colIndex['Accès Direct Formulaire'] + 1).setValue(form.getEditUrl());
    configSheet.getRange(rowIndex, colIndex['Statut'] + 1).setValue('Actif - Prêt');

    Logger.log(`[ÉTAPE 2] Déploiement pour la ligne ${rowIndex} terminé.`);

  } catch (e) {
    Logger.log(`ERREUR Critique (ÉTAPE 2, ligne ${rowIndex}) : ${e.toString()}\n${e.stack}`);
    throw e;
  }
}


/**
 * ÉTAPE 3 : Vérifie la cohérence du kit et écrit le rapport dans un onglet dédié.
 */
function etape3_verifierKit(rowIndex) {
  // ... (cette fonction reste inchangée)
  Logger.log(`Lancement de l'Étape 3 (Vérification) pour la ligne ${rowIndex}...`);
  const rapport = [];
  try {
    const config = getConfigurationFromRow(rowIndex);
    const formId = config['ID_Formulaire_Cible'];
    const sheetId = config['ID_Sheet_Cible'];
    const lienForm = config['Accès Direct Formulaire'];
    const lienPublic = config['Lien_Formulaire_Public'];
    const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName('Paramètres Généraux');
    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    const colIndex = {};
    headers.forEach((h, i) => { if (h) colIndex[String(h).trim()] = i; });

    try {
      const form = FormApp.openById(formId);
      rapport.push(`✅ Formulaire trouvé : ${form.getTitle()}\n   (ID: ${formId})`);
    } catch (e) {
      rapport.push(`❌ Formulaire introuvable (ID : ${formId})`);
    }

    try {
      const sheet = SpreadsheetApp.openById(sheetId);
      rapport.push(`✅ Feuille de réponses trouvée : ${sheet.getName()}\n   (ID: ${sheetId})`);
    } catch (e) {
      rapport.push(`❌ Feuille introuvable (ID : ${sheetId})`);
    }

    try {
      const form = FormApp.openById(formId);
      const dest = form.getDestinationId();
      if (dest === sheetId) {
        rapport.push('✅ Liaison Form → Sheet confirmée.');
      } else {
        rapport.push(`⚠️ Liaison Form → Sheet différente (actuelle : ${dest || 'aucune'}).`);
      }
    } catch (e) {
      rapport.push('❌ Impossible de vérifier la liaison Form → Sheet.');
    }

    if (lienForm && /https?:\/\/(docs|forms)\.google\.com\/forms\/.+\/edit/.test(lienForm)) {
      rapport.push(`✅ Lien d'édition valide :\n   ${lienForm}`);
    } else {
      rapport.push('⚠️ Lien d\'édition manquant ou invalide.');
    }
    if (lienPublic && String(lienPublic).includes('script.google.com')) {
      rapport.push(`✅ Lien public (macro) présent :\n   ${lienPublic}`);
    } else {
      rapport.push('⚠️ Lien public (macro) manquant ou invalide.');
    }

    const ok = rapport.every(r => r.startsWith('✅') || r.startsWith('⚠️'));
    configSheet.getRange(rowIndex, colIndex['Statut'] + 1).setValue(ok ? '✅ Vérifié' : '⚠️ À revoir');
    
    const rapportSheetName = 'Rapport de Vérification';
    const moteurSS = SpreadsheetApp.getActiveSpreadsheet();
    let rapportSheet = moteurSS.getSheetByName(rapportSheetName);

    if (!rapportSheet) {
      rapportSheet = moteurSS.insertSheet(rapportSheetName);
    }

    rapportSheet.clear();
    rapportSheet.getRange('A1').setValue(`Rapport de Vérification - Ligne ${rowIndex} (Généré le ${new Date().toLocaleString()})`).setFontWeight('bold');
    
    const outputData = rapport.map(item => [item]);
    rapportSheet.getRange(2, 1, outputData.length, 1).setValues(outputData).setWrap(true);
    
    rapportSheet.autoResizeColumn(1);
    moteurSS.setActiveSheet(rapportSheet);

    SpreadsheetApp.getUi().alert('Rapport Terminé', `Le rapport de vérification a été généré dans l'onglet "${rapportSheetName}".`, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (e) {
    Logger.log(`ERREUR Critique (ÉTAPE 3, ligne ${rowIndex}) : ${e.toString()}\n${e.stack}`);
    SpreadsheetApp.getUi().alert('❌ ERREUR', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}