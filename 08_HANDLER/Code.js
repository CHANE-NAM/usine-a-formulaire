/**
 * Clé secrète pour le déchiffrement. DOIT être la même que dans le MOTEUR.
 */
const SECRET_KEY = "FELIX QUI POTUIT RERUM COGNOCERE CAUSA VIC";

/**
 * ID de la feuille de calcul de configuration [CONFIG].
 */
const ID_FEUILLE_CONFIG = '1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ';

/**
 * Fonction principale qui s'exécute lorsqu'un utilisateur accède à l'URL de l'application Web.
 * @param {Object} e - L'objet d'événement contenant les paramètres de la requête.
 */
function doGet(e) {
  try {
    // 1. Vérifier et déchiffrer le code
    if (!e.parameter.code) {
      return HtmlService.createHtmlOutput('<h1>Erreur : Paramètre d\'identification du test manquant.</h1>');
    }
    const encryptedCode = e.parameter.code.replace(/ /g, '+');
    
    let payload;
    try {
      const decryptedBytes = CryptoJS.AES.decrypt(encryptedCode, SECRET_KEY);
      const decryptedText = decryptedBytes.toString(CryptoJS.enc.Utf8);
      if (!decryptedText) throw new Error("Le déchiffrement a produit un texte vide.");
      payload = JSON.parse(decryptedText);
    } catch (err) {
      Logger.log('Erreur de déchiffrement : ' + err.message);
      return HtmlService.createHtmlOutput('<h1>Erreur : Le code du test est invalide ou a expiré.</h1>');
    }

    // 2. Récupérer la configuration complète du test
    const config = getConfigFromRow(payload.rowIndex);
    if (!config) {
      return HtmlService.createHtmlOutput('<h1>Erreur : Impossible de trouver la configuration pour ce test.</h1>');
    }

    // 3. Aiguillage en fonction du parcours utilisateur
    switch (config.Mode_Acces_Test) {
      case 'Payant':
        return afficherPagePaiement(config, encryptedCode); 
      
      case 'B2B':
        return afficherPageB2B(config, encryptedCode);

      case 'CLOS':
        const userEmail = Session.getActiveUser().getEmail();
        const isAdmin = estAdmin(userEmail);
        if (!isAdmin) {
          return HtmlService.createHtmlOutput('<h1>Ce test est actuellement fermé.</h1>');
        }
        return afficherPageIdentificationGratuite(config, encryptedCode);

      case 'Gratuit':
      default:
        return afficherPageIdentificationGratuite(config, encryptedCode);
    }

  } catch (error) {
    Logger.log("ERREUR FATALE dans doGet(e) : " + error.toString());
    return HtmlService.createHtmlOutput("<h1>Une erreur interne est survenue.</h1><p>Veuillez contacter l'administrateur.</p>");
  }
}

// ===============================================================
// == FONCTIONS D'AFFICHAGE DES PAGES
// ===============================================================

/**
 * Affiche la page d'identification standard pour un accès gratuit.
 */
function afficherPageIdentificationGratuite(config, encryptedCode) {
  const htmlTemplate = HtmlService.createTemplateFromFile('Index.html');
  htmlTemplate.config = config;
  htmlTemplate.encryptedCode = encryptedCode;
  return htmlTemplate.evaluate()
    .setTitle(config['Titre_Formulaire_Utilisateur'])
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Affiche la page d'identification pour un accès payant.
 */
function afficherPagePaiement(config, encryptedCode) {
  const template = HtmlService.createTemplateFromFile('Paiement.html');
  template.encryptedCode = encryptedCode;
  template.titreTest = config.Titre_Formulaire_Utilisateur;
  return template.evaluate().setTitle("Identification pour le test");
}

/**
 * Affiche la page d'identification avec code d'accès pour le B2B.
 */
function afficherPageB2B(config, encryptedCode) {
  const template = HtmlService.createTemplateFromFile('B2B.html');
  template.encryptedCode = encryptedCode;
  template.titreTest = config.Titre_Formulaire_Utilisateur;
  return template.evaluate().setTitle("Accès Entreprise");
}

// ===============================================================
// == FONCTIONS UTILITAIRES
// ===============================================================

/**
 * Récupère la configuration complète pour une ligne donnée.
 */
function getConfigFromRow(rowIndex) {
  const configSS = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
  const paramsSheet = configSS.getSheetByName('Paramètres Généraux');
  const headers = paramsSheet.getRange(1, 1, 1, paramsSheet.getLastColumn()).getValues()[0];
  const configData = paramsSheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

  const config = headers.reduce((obj, header, index) => {
    obj[header] = configData[index];
    return obj;
  }, {});
  return config;
}

/**
 * Vérifie si un utilisateur est un administrateur.
 */
function estAdmin(userEmail) {
  const configSS = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
  const optionsSheet = configSS.getSheetByName('sys_Options_Parametres');
  const adminEmailsRange = optionsSheet.getRange(2, 26, optionsSheet.getLastRow() - 1, 1).getValues();
  const ADMIN_EMAILS = adminEmailsRange.map(row => row[0].trim()).filter(String);
  return ADMIN_EMAILS.includes(userEmail);
}

/**
 * Traite les données soumises depuis le formulaire HTML d'identification.
 * @param {Object} formObject - Un objet représentant les données du formulaire (nom, email, etc.).
 * @returns {String} Une chaîne de caractères HTML à afficher à l'utilisateur.
 */
function processFormSubmission(formObject) {
  try {
    // --- ÉTAPE 1 : Valider et déchiffrer le code pour retrouver la configuration ---
    const encryptedCode = formObject.encryptedCode;
    if (!encryptedCode) {
      throw new Error("Le code de session est manquant. Impossible de continuer.");
    }
    
    let payload;
    try {
      const decryptedBytes = CryptoJS.AES.decrypt(encryptedCode.replace(/ /g, '+'), SECRET_KEY);
      const decryptedText = decryptedBytes.toString(CryptoJS.enc.Utf8);
      payload = JSON.parse(decryptedText);
    } catch (e) {
      throw new Error("Le code de session est invalide.");
    }
    const rowIndex = payload.rowIndex;
    const config = getConfigFromRow(rowIndex);

    // --- NOUVEAU : Gérer les logiques spécifiques par type de formulaire ---
    let detailsPourOrders = ""; // Variable pour stocker les infos supplémentaires

    if (formObject.formType === 'b2b') {
      const codeAcces = formObject.code; // 'code' est le "name" de notre champ dans B2B.html
      Logger.log(`Parcours B2B détecté. Code d'accès fourni : ${codeAcces}`);
      detailsPourOrders = `Code: ${codeAcces}`;
      // NOTE FUTURE : C'est ici que nous ajouterons la logique pour valider le code d'accès.
    } else if (formObject.formType === 'paiement') {
      Logger.log(`Parcours Payant détecté.`);
      detailsPourOrders = "Parcours Payant";
      // NOTE FUTURE : C'est ici qu'on lancerait la redirection vers un module de paiement.
    }

    // --- ÉTAPE 2 : Enregistrer les informations dans la feuille "Orders" ---
    const configSS = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
    let ordersSheet = configSS.getSheetByName("Orders");
    
    if (!ordersSheet) {
      ordersSheet = configSS.insertSheet("Orders");
      ordersSheet.appendRow(["Timestamp", "Nom", "Email", "Test Row Index", "Statut", "Détails"]);
    }
    
    ordersSheet.appendRow([
      new Date(),
      formObject.nom,
      formObject.email,
      rowIndex,
      "IDENTIFIED",
      detailsPourOrders // On ajoute les détails spécifiques
    ]);

// --- ÉTAPE 3 : Récupérer l'URL du formulaire de test public depuis la configuration ---
    const testFormUrl = config.Lien_Formulaire_Public;
    if (!testFormUrl) {
      throw new Error("L'URL du formulaire de test (Lien_Formulaire_Public) n'est pas configurée pour cette ligne.");
    }

    // --- ÉTAPE 4 : Renvoyer un objet de redirection au client ---
    // Le code côté navigateur (dans Index.html) interprétera cet objet 
    // et effectuera la redirection lui-même.
    return { redirectUrl: testFormUrl };
    


  } catch (error) {
    Logger.log("Erreur critique dans processFormSubmission: " + error.toString());
    throw new Error("Une erreur interne est survenue. Veuillez réessayer. Détails : " + error.message);
  }
}


/**
 * Fonction de compatibilité, non utilisée dans le flux principal.
 */
function handleFormSubmit(e) {
  try {
    const configSpreadsheet = SpreadsheetApp.openById("1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ");
    const ordersSheet = configSpreadsheet.getSheetByName("Orders");

    if (!ordersSheet) {
      Logger.log("ERREUR CRITIQUE: L'onglet 'Orders' est introuvable.");
      return;
    }

    const itemResponses = e.response.getItemResponses();
    let emailRepondant = '';
    let nomRepondant = '';

    for (let i = 0; i < itemResponses.length; i++) {
      const question = itemResponses[i].getItem().getTitle().toLowerCase();
      const reponse = itemResponses[i].getResponse();

      if (question.includes('mail')) {
        emailRepondant = reponse;
      } else if (question.includes('nom')) {
        nomRepondant = reponse;
      }
    }

    ordersSheet.appendRow([new Date(), emailRepondant, nomRepondant]);
    Logger.log('SUCCÈS (via handleFormSubmit) : Ligne ajoutée à l\'onglet Orders. Email: ' + emailRepondant + ', Nom: ' + nomRepondant);

  } catch (err) {
    Logger.log('ERREUR dans handleFormSubmit: ' + err.toString() + "\n" + err.stack);
  }
}