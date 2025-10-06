// =================================================================================
// == PROJET [MOTEUR] - FICHIER LOGIQUE MÉTIER
// == VERSION : 12.1 (Migration API Forms REST)
// == RÔLE    : Contient la logique de déploiement manuel en deux étapes.
// =================================================================================

/**
 * ÉTAPE 1 : Crée les fichiers (Form & Sheet) et met à jour la ligne de configuration avec les IDs.
 */
function etape1_creerKit(rowIndex) {
  Logger.log(`Lancement de l'Étape 1 (Création) pour la ligne ${rowIndex}...`);
  try {
    const config = getConfigurationFromRow(rowIndex);
    const nomFichierComplet = `${new Date().getFullYear()}${('0' + (new Date().getMonth() + 1)).slice(-2)}${('0' + new Date().getDate()).slice(-2)}_${config['Type_Test'] || 'TypeInconnu'}_${config['Statut_Deploiement'] || 'alpha'}_${config['nbQuestions'] || 'Nq'}q`;

    const systemIds = getSystemIds();
    const dossierCible = config['ID_Dossier_Cible']
      ? DriveApp.getFolderById(config['ID_Dossier_Cible'])
      : DriveApp.getFolderById(systemIds.ID_DOSSIER_CIBLE_GEN);

    // 1. Création du Sheet
    const templateSheet = SpreadsheetApp.openById(systemIds.ID_TEMPLATE_TRAITEMENT_V2);
    const ssCopy = templateSheet.copy(nomFichierComplet);
    const sheetId = ssCopy.getId();

    // 2. Création du Form (copie Drive)
    const templateFormFile = DriveApp.getFileById(systemIds.ID_TEMPLATE_FORMULAIRE);
    const formFile = templateFormFile.makeCopy(nomFichierComplet, dossierCible);
    const formId = formFile.getId();
    Logger.log(`[ÉTAPE 1] Fichiers créés. ID Form: ${formId}, ID Sheet: ${sheetId}.`);

    // 3. Mise à jour de la feuille CONFIG
    const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName("Paramètres Généraux");
    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    const colIndex = {};
    headers.forEach((header, i) => { if (header) colIndex[header.trim()] = i; });

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
      throw new Error("ID de formulaire ou de feuille manquant sur la ligne " + rowIndex + ". Lancez d'abord l'Étape 1.");
    }

    // --- 1. CONFIGURATION VIA L'API FORMS REST ---
    Logger.log(`[ÉTAPE 2] Configuration du titre et des paramètres via l'API Forms...`);
    const requests = [
      {
        updateFormInfo: {
          info: {
            title: config['Titre_Formulaire_Utilisateur'] || "Formulaire sans titre",
            description: config['Sous-Titre_Formulaire'] || ""
          },
          updateMask: "title,description"
        }
      },
      {
        updateSettings: {
          settings: {
            quizSettings: { isQuiz: false }
          },
          updateMask: "quizSettings.isQuiz"
        }
      }
    ];

    const url = `https://forms.googleapis.com/v1/forms/${formId}:batchUpdate`;
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ requests }),
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    Logger.log(`[ÉTAPE 2] Réponse API : ${response.getResponseCode()} - ${response.getContentText()}`);

    if (response.getResponseCode() >= 300) {
      throw new Error("Erreur API Forms : " + response.getContentText());
    }

    Logger.log(`[ÉTAPE 2] Titre et paramètres de base mis à jour avec succès.`);

    // --- 2. CONSTRUCTION DES QUESTIONS (FormApp reste utilisable pour cela) ---
    const form = FormApp.openById(formId);
    const bdd = SpreadsheetApp.openById(getSystemIds().ID_BDD);
    const languesAInclure = _identifierLangues(bdd, config['Type_Test']);
    _construireQuestionsFormulaire(form, languesAInclure, config['nbQuestions']);
    Logger.log('[ÉTAPE 2] Construction des questions terminée.');

    // --- 3. LIAISON DE LA FEUILLE DE RÉPONSES ---
    try {
      form.setDestination(FormApp.DestinationType.SPREADSHEET, sheetId);
      Logger.log('[ÉTAPE 2] Liaison de la feuille de réponses réussie.');
    } catch (e) {
      Logger.log(`[ÉTAPE 2] AVERTISSEMENT : Liaison manuelle requise (${e.message})`);
      SpreadsheetApp.getUi().alert(
        "Liaison manuelle requise",
        "La configuration est presque terminée, mais la liaison à la feuille de réponses doit être faite manuellement.\n\nOuvrez le formulaire, onglet 'Réponses' → Sélectionnez la feuille de calcul de destination.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }

    // --- 4. MISE À JOUR FINALE DE LA FEUILLE CONFIG ---
    const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName("Paramètres Généraux");
    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    const colIndex = {};
    headers.forEach((header, i) => { if (header) colIndex[header.trim()] = i; });

    const urlHandler = getSystemIds().ID_WEBAPP_HANDLER;
    const payload = { rowIndex: rowIndex };
    const encryptedCode = CryptoJS.AES.encrypt(
      JSON.stringify(payload),
      "FELIX QUI POTUIT RERUM COGNOCERE CAUSA VIC"
    ).toString();
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
 * ÉTAPE 3 : Vérifie la cohérence du kit déployé.
 * Contrôle : Form existant, Sheet existante, lien actif, statut cohérent.
 */
function etape3_verifierKit(rowIndex) {
  Logger.log(`Lancement de l'Étape 3 (Vérification) pour la ligne ${rowIndex}...`);
  const rapport = [];

  try {
    const config = getConfigurationFromRow(rowIndex);
    const formId = config["ID_Formulaire_Cible"];
    const sheetId = config["ID_Sheet_Cible"];
    const statutCell = "Statut";
    const configSheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION).getSheetByName("Paramètres Généraux");
    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    const colIndex = {};
    headers.forEach((h, i) => { if (h) colIndex[h.trim()] = i; });

    // 1️⃣ Vérifier la présence du Formulaire
    try {
      const form = FormApp.openById(formId);
      rapport.push(`✅ Formulaire trouvé : ${form.getTitle()}`);
    } catch (e) {
      rapport.push(`❌ Formulaire introuvable ou non accessible (ID : ${formId})`);
    }

    // 2️⃣ Vérifier la présence de la Feuille
    try {
      const sheet = SpreadsheetApp.openById(sheetId);
      rapport.push(`✅ Feuille de réponses trouvée : ${sheet.getName()}`);
    } catch (e) {
      rapport.push(`❌ Feuille introuvable ou non accessible (ID : ${sheetId})`);
    }

    // 3️⃣ Vérifier la liaison Form → Sheet
    try {
      const form = FormApp.openById(formId);
      const dest = form.getDestinationId();
      if (dest === sheetId) {
        rapport.push("✅ Liaison Form → Sheet confirmée.");
      } else {
        rapport.push(`⚠️ Liaison Form → Sheet différente (actuelle : ${dest || "aucune"}).`);
      }
    } catch (e) {
      rapport.push("❌ Impossible de vérifier la liaison Form → Sheet.");
    }

    // 4️⃣ Vérifier les liens dans la feuille
    const lienForm = config["Accès Direct Formulaire"];
    const lienPublic = config["Lien_Formulaire_Public"];
    if (lienForm && /https?:\/\/(docs|forms)\.google\.com\/forms\/.+\/edit/.test(lienForm)) {
      rapport.push("✅ Lien d'édition du formulaire valide.");
    } else {
      rapport.push("⚠️ Lien d'édition du formulaire manquant ou invalide.");
    }
    if (lienPublic && lienPublic.includes("script.google.com")) {
      rapport.push("✅ Lien public (macro) présent.");
    } else {
      rapport.push("⚠️ Lien public (macro) manquant ou invalide.");
    }

    // 5️⃣ Mise à jour du statut et affichage
    const ok = rapport.every(r => r.startsWith("✅") || r.startsWith("⚠️"));
    configSheet.getRange(rowIndex, colIndex[statutCell] + 1).setValue(ok ? "✅ Vérifié" : "⚠️ À revoir");

    Logger.log("RAPPORT DE VÉRIFICATION :\n" + rapport.join("\n"));
    SpreadsheetApp.getUi().alert("Rapport de vérification", rapport.join("\n\n"), SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (e) {
    Logger.log(`ERREUR Critique (ÉTAPE 3, ligne ${rowIndex}) : ${e.toString()}\n${e.stack}`);
    SpreadsheetApp.getUi().alert("❌ ERREUR", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
